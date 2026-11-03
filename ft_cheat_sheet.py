"""
Wofford Men's Basketball - Free Throw Foul Cheat Sheet Generator

What it does
  1. Pulls Wofford's schedule from ESPN and every opponent's roster.
  2. For every player, gets free throw makes/attempts from ESPN:
       - this season's FT% once the player has MIN_FTA_SEASON attempts this season
       - otherwise career FT% (all schools, including transfers)
  3. Expected points per trip to the line:
       1-and-1 : p + p^2      (make the front end, then shoot the second)
       2 shots : 2p
  4. Scrapes TeamRankings points per game and possessions per game for every
     team; points per possession = PPG / possessions per game.
       - uses 2026-27 once the team has played MIN_GAMES_TEAM games
       - otherwise uses 2025-26 (final numbers as of 2026-04-07)
  5. Writes:
       Opponent_FT_Report.html/.pdf    - all opponents, one page each, worst shooters first
       teams/<Team>.html/.pdf          - one printable file per opponent
       index.html                      - landing page linking everything (the website)
       Generic_FT_Cheat_Sheet.html/.pdf - 40%..99% table, one printable page
       FT_Data.xlsx              - all the raw numbers

Rule of thumb printed on the sheets: fouling a shooter is favorable when their
expected points from the line are LOWER than their team's points per possession.

Run:  py ft_cheat_sheet.py
      py ft_cheat_sheet.py --min-fta 30 --min-games 10
"""

import argparse
import base64
import concurrent.futures as cf
import datetime as dt
import difflib
import html
import io
import re
import shutil
import subprocess
from urllib.parse import quote
import time
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

WOFFORD_ESPN_ID = "2747"
ESPN_SITE = "https://site.api.espn.com/apis/site/v2/sports/basketball/mens-college-basketball"
ESPN_ATHLETE = "https://site.web.api.espn.com/apis/common/v3/sports/basketball/mens-college-basketball/athletes/{id}/stats"
TR_STAT = "https://www.teamrankings.com/ncaa-basketball/stat/{stat}"
PRIOR_SEASON_DATE = "2026-04-07"   # final 2025-26 numbers
PRIOR_SEASON_LABEL = "2025-26"

# ESPN team id -> TeamRankings slug, for names the auto-matcher can't get.
TR_OVERRIDES = {
    "2511": "queens-university-of-charlotte-royals",
}

OUT_DIR = Path(__file__).resolve().parent     # change with --out

SESSION = requests.Session()
SESSION.headers["User-Agent"] = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/130.0 Safari/537.36"
)


def get(url, params=None, tries=3):
    for i in range(tries):
        try:
            r = SESSION.get(url, params=params, timeout=30)
            if r.status_code == 404:
                return None
            r.raise_for_status()
            return r
        except requests.RequestException:
            if i == tries - 1:
                raise
            time.sleep(2 * (i + 1))


# ----------------------------------------------------------------------------
# Free throw math
# ----------------------------------------------------------------------------
def one_and_one(p):
    return p + p * p


def two_shots(p):
    return 2 * p


# ----------------------------------------------------------------------------
# ESPN
# ----------------------------------------------------------------------------
def _eastern_tz():
    try:
        from zoneinfo import ZoneInfo
        return ZoneInfo("America/New_York")
    except Exception:          # no tz database: EST covers nearly all of the season
        return dt.timezone(dt.timedelta(hours=-5))


def to_eastern(espn_date):
    utc = dt.datetime.strptime(espn_date.replace("Z", ""), "%Y-%m-%dT%H:%M").replace(tzinfo=dt.timezone.utc)
    return utc.astimezone(_eastern_tz())


def now_eastern():
    return dt.datetime.now(_eastern_tz())


def get_schedule():
    d = get(f"{ESPN_SITE}/teams/{WOFFORD_ESPN_ID}/schedule").json()
    season_label = d["season"]["displayName"]          # e.g. "2026-27"
    season_year = d["season"]["year"]                  # e.g. 2027
    games = []
    for e in d.get("events", []):
        comp = e["competitions"][0]
        opp = next(c for c in comp["competitors"] if c["team"]["id"] != WOFFORD_ESPN_ID)
        home = next(c for c in comp["competitors"] if c["team"]["id"] == WOFFORD_ESPN_ID)["homeAway"]
        tip = to_eastern(e["date"])            # ESPN gives UTC; 7 PM ET shows as next day
        games.append({
            "date": tip.date().isoformat(),
            "tip": tip,
            "opp_id": opp["team"]["id"],
            "opp_name": opp["team"]["displayName"],
            "opp_short": opp["team"].get("location") or opp["team"]["displayName"],
            "site": "vs" if home == "home" else "at",
        })
    return season_label, season_year, games


def get_team_games_played(team_id):
    r = get(f"{ESPN_SITE}/teams/{team_id}")
    if r is None:
        return 0
    for item in r.json()["team"].get("record", {}).get("items", []):
        for s in item.get("stats", []):
            if s["name"] == "gamesPlayed":
                return int(s["value"])
    return 0


def get_roster(team_id):
    r = get(f"{ESPN_SITE}/teams/{team_id}/roster")
    if r is None:
        return []
    out = []
    for a in r.json().get("athletes", []):
        out.append({
            "id": a["id"],
            "name": a["displayName"],
            "jersey": a.get("jersey", ""),
            "pos": (a.get("position") or {}).get("abbreviation", ""),
            "class": (a.get("experience") or {}).get("abbreviation", ""),
        })
    return out


def _split(made_att):
    m, a = made_att.split("-")
    return int(float(m)), int(float(a))


def get_player_ft(athlete_id, season_label):
    """Return (season_ftm, season_fta, career_ftm, career_fta)."""
    r = get(ESPN_ATHLETE.format(id=athlete_id))
    if r is None:
        return 0, 0, 0, 0
    cats = {c["name"]: c for c in r.json().get("categories", [])}
    tot = cats.get("totals")
    if not tot:
        return 0, 0, 0, 0
    idx = tot["names"].index("freeThrowsMade-freeThrowsAttempted")
    s_m = s_a = 0
    for row in tot.get("statistics", []):
        if row.get("season", {}).get("displayName") == season_label:
            m, a = _split(row["stats"][idx])     # sum: mid-season transfers have 2 rows
            s_m += m
            s_a += a
    c_m, c_a = _split(tot["totals"][idx]) if tot.get("totals") else (0, 0)
    return s_m, s_a, c_m, c_a


# ----------------------------------------------------------------------------
# TeamRankings
# ----------------------------------------------------------------------------
def scrape_tr(stat, date=None):
    """Return DataFrame: slug, team, value  (value for the season shown on that page)."""
    params = {"date": date} if date else None
    soup = BeautifulSoup(get(TR_STAT.format(stat=stat), params=params).text, "lxml")
    rows = []
    for tr in soup.select("table.datatable tbody tr"):
        tds = tr.find_all("td")
        a = tds[1].find("a") if len(tds) > 2 else None
        if not a:
            continue
        m = re.search(r"/ncaa-basketball/team/([a-z0-9-]+)", a["href"])
        if not m:
            continue
        # tds[2] = first stat column = season of that page
        rows.append({"slug": m.group(1), "team": a.get_text(strip=True),
                     "value": pd.to_numeric(tds[2].get_text(strip=True), errors="coerce")})
    if len(rows) < 300:
        raise RuntimeError(f"TeamRankings page layout changed ({stat}): only {len(rows)} teams parsed")
    return pd.DataFrame(rows)


def _norm(name):
    s = name.lower().replace("&", "").replace("'", "").replace(".", "")
    toks = re.split(r"[^a-z0-9]+", s)
    repl = {"st": "state", "u": "university"}
    drop = {"university", "of", "the", ""}
    return " ".join(repl.get(t, t) for t in toks if repl.get(t, t) not in drop)


def match_tr_slug(espn_id, espn_display_name, tr_slugs):
    if espn_id in TR_OVERRIDES:
        return TR_OVERRIDES[espn_id]
    lookup = {_norm(s): s for s in tr_slugs}
    key = _norm(espn_display_name)
    if key in lookup:
        return lookup[key]
    close = difflib.get_close_matches(key, list(lookup), n=1, cutoff=0.85)
    return lookup[close[0]] if close else None


# ----------------------------------------------------------------------------
# Non-D1 opponents: Sidearm Sports school stats pages
# ----------------------------------------------------------------------------
# ESPN team id -> school stats page. Add a season with /<season> (e.g. .../stats/2025-26).
SIDEARM_SOURCES = {
    "2019": "https://auyellowjackets.com/sports/mens-basketball/stats",     # Allen
    "2679": "https://voorheesathletics.com/sports/mens-basketball/stats",   # Voorhees
}
SIDEARM_SEASONS_BACK = 4      # seasons at that school summed for "career"


def _flat(df):
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [b for a, b in df.columns]
    return df


def _first_last(name):
    """'Williams , Syris' -> 'Syris Williams'"""
    if "," in name:
        last, first = name.split(",", 1)
        name = f"{first.strip()} {last.strip()}"
    return " ".join(name.split())


def _name_key(name):
    name = name.strip()
    if "," in name:                                  # "Johnson, Jordahn" -> "Jordahn Johnson"
        last, first = name.split(",", 1)
        name = f"{first} {last}"
    toks = re.sub(r"[^a-z ]", "", name.lower()).split()
    return " ".join(t for t in toks if t not in {"jr", "sr", "ii", "iii", "iv"})


def sidearm_season(base, season):
    """Players' FT totals and team possession stats for one season, or None if no stats."""
    r = get(f"{base}/{season}")
    if r is None:
        return None
    try:
        tables = [_flat(t) for t in pd.read_html(io.StringIO(r.text))]
    except ValueError:                               # season page exists but has no stats yet
        return None
    ind = next((t for t in tables if {"#", "Player", "FTM", "FTA"} <= set(t.columns)), None)
    log = next((t for t in tables if {"Opponent", "Date", "FGM/A", "FTM/A", "OFF", "TO", "PTS"} <= set(t.columns)),
               None)
    if ind is None or log is None:
        return None
    players = {}
    for _, row in ind.iterrows():
        raw = str(row["Player"])
        if row["#"] in ("TM",) or pd.isna(row["#"]) or raw in ("Total", "Opponents"):
            continue
        display = _first_last(re.split(r"\s+\d{1,2}\s+", raw)[0])   # "Last, First  22 Last, First"
        # schools sometimes list one player on several rows (split stat entries) - add them up
        p = players.setdefault(_name_key(display), {"jersey": str(row["#"]).lstrip("0") or "0",
                                                    "name": display, "ftm": 0, "fta": 0})
        p["ftm"] += int(row["FTM"])
        p["fta"] += int(row["FTA"])
    games = log[log["Date"].notna()]
    tot = log[log["Date"].isna() & (log["Opponent"] != "Opponents")].iloc[0]
    fga, fta = int(str(tot["FGM/A"]).split("-")[1]), int(str(tot["FTM/A"]).split("-")[1])
    poss = fga - int(tot["OFF"]) + int(tot["TO"]) + 0.475 * fta
    team = {"season": season, "gp": len(games), "pts": int(tot["PTS"]),
            "ppp": round(int(tot["PTS"]) / poss, 3) if poss else None}
    return {"players": players, "team": team}


def load_sidearm(base, season_label):
    start = int(season_label[:4])
    seasons = [f"{y}-{str(y + 1)[2:]}" for y in range(start, start - SIDEARM_SEASONS_BACK, -1)]
    data = {s: sidearm_season(base, s) for s in seasons}
    return {"seasons": data, "team": {s: d["team"] for s, d in data.items() if d}}


def sidearm_roster_ft(espn_roster, sidearm, season_label):
    """Yield (player, (season_ftm, season_fta, career_ftm, career_fta)) for each player on the roster.
    Uses ESPN's roster, plus anyone in this season's school stats who isn't on it."""
    seasons = {s: d["players"] for s, d in sidearm["seasons"].items() if d}
    cur = seasons.get(season_label, {})
    roster = list(espn_roster)
    have = {_name_key(p["name"]) for p in roster}
    for key, sp in cur.items():
        if key not in have:
            roster.append({"id": "", "name": sp["name"], "jersey": sp["jersey"], "pos": "", "class": ""})
    for p in roster:
        key = _name_key(p["name"])
        s = cur.get(key, {"ftm": 0, "fta": 0})
        c_m = sum(d[key]["ftm"] for d in seasons.values() if key in d)
        c_a = sum(d[key]["fta"] for d in seasons.values() if key in d)
        yield p, (s["ftm"], s["fta"], c_m, c_a)


# ----------------------------------------------------------------------------
# Main build
# ----------------------------------------------------------------------------
def build(min_fta, min_games):
    season_label, season_year, games = get_schedule()
    print(f"Wofford {season_label}: {len(games)} games on schedule")

    print("Scraping TeamRankings ...")
    tr = {
        "ppg_now": scrape_tr("points-per-game"),
        "ppg_prior": scrape_tr("points-per-game", PRIOR_SEASON_DATE),
        "pace_now": scrape_tr("possessions-per-game"),
        "pace_prior": scrape_tr("possessions-per-game", PRIOR_SEASON_DATE),
    }
    tr_idx = {k: v.set_index("slug") for k, v in tr.items()}
    all_slugs = set(tr["pace_prior"].slug) | set(tr["pace_now"].slug)

    # Unique opponents in schedule order
    opponents = {}
    for g in games:
        opponents.setdefault(g["opp_id"], {"id": g["opp_id"], "name": g["opp_name"],
                                           "short": g["opp_short"], "games": []})
        opponents[g["opp_id"]]["games"].append(
            f'{g["site"]} {g["tip"].strftime("%a %b")} {g["tip"].day}')

    team_rows = []
    for opp in opponents.values():
        slug = match_tr_slug(opp["id"], opp["name"], all_slugs)
        gp = get_team_games_played(opp["id"])
        use_now = gp >= min_games

        def val(kind):
            df = tr_idx[f"{kind}_{'now' if use_now else 'prior'}"]
            if slug and slug in df.index and pd.notna(df.at[slug, "value"]):
                return float(df.at[slug, "value"])
            return None

        ppg, pace = val("ppg"), val("pace")
        opp.update({
            "tr_slug": slug,
            "tr_name": tr_idx["pace_prior"].at[slug, "team"] if slug in tr_idx["pace_prior"].index else "",
            "games_played": gp,
            "ppg": ppg,
            "pace": pace,
            "ppp": round(ppg / pace, 3) if ppg and pace else None,   # points per possession
            "ppp_season": season_label if use_now else PRIOR_SEASON_LABEL,
        })
        sidearm = None
        if opp["id"] in SIDEARM_SOURCES:      # non-D1: school's own stats site
            sidearm = load_sidearm(SIDEARM_SOURCES[opp["id"]], season_label)
            cur, prior = sidearm["team"].get(season_label), sidearm["team"].get(PRIOR_SEASON_LABEL)
            pick = cur if cur and cur["gp"] >= min_games else (prior or cur)
            if pick:
                opp.update({"ppp": pick["ppp"], "ppg": None, "pace": None, "games_played": cur["gp"] if cur else 0,
                            "ppp_season": pick["season"], "tr_name": "school site"})
        opp["sidearm"] = sidearm
        team_rows.append(opp)
        print(f'  {opp["name"]:<35} -> {opp["tr_name"] or "NOT FOUND":<18} '
              f'PPP {opp["ppp"] if opp["ppp"] is not None else "--"} ({opp["ppp_season"]}, {gp} GP)')

    print("Pulling rosters and free throw stats from ESPN ...")
    player_rows = []
    jobs = []
    with cf.ThreadPoolExecutor(max_workers=6) as pool:
        for opp in team_rows:
            roster = get_roster(opp["id"])
            if opp["sidearm"]:
                for p, ft in sidearm_roster_ft(roster, opp["sidearm"], season_label):
                    jobs.append((opp, p, ft))
            else:
                for p in roster:
                    jobs.append((opp, p, pool.submit(get_player_ft, p["id"], season_label)))
        for opp, p, fut in jobs:
            s_m, s_a, c_m, c_a = fut.result() if isinstance(fut, cf.Future) else fut
            if s_a >= min_fta:
                pct, basis, fta_used = s_m / s_a, season_label, s_a
            elif c_a > 0:
                pct, basis, fta_used = c_m / c_a, "Career", c_a
            else:
                pct, basis, fta_used = None, "No FT data", 0
            player_rows.append({
                "team_id": opp["id"], "team": opp["name"], **p,
                "season_ft": f"{s_m}-{s_a}", "season_pct": s_m / s_a if s_a else None,
                "career_ft": f"{c_m}-{c_a}", "career_pct": c_m / c_a if c_a else None,
                "ft_pct": pct, "basis": basis, "fta_used": fta_used,
                "epp_1and1": one_and_one(pct) if pct is not None else None,
                "epp_2shot": two_shots(pct) if pct is not None else None,
                "team_ppp": opp["ppp"],
            })
    print(f"  {len(player_rows)} players")
    return season_label, games, team_rows, player_rows


# ----------------------------------------------------------------------------
# Output
# ----------------------------------------------------------------------------
BLACK = "#111111"
GOLD = "#B9975B"        # Wofford old gold
GOLD_DARK = "#86693A"   # gold dark enough for text on white

CSS = f"""
*{{box-sizing:border-box;-webkit-print-color-adjust:exact;print-color-adjust:exact}}
body{{font-family:"Segoe UI",Roboto,Helvetica,Arial,sans-serif;margin:0;color:{BLACK};background:#fff}}
.sheet{{max-width:7.6in;margin:0 auto;padding:0.15in 0 0.1in}}
.band{{display:flex;align-items:center;justify-content:space-between;gap:14px;background:{BLACK};
  color:#fff;padding:12px 18px;border-bottom:5px solid {GOLD};border-radius:6px 6px 0 0}}
.band .logo{{width:62px;height:62px;border-radius:50%;background:#fff;display:flex;align-items:center;
  justify-content:center;flex:none;border:2px solid {GOLD}}}
.band .logo img{{width:48px;height:48px;object-fit:contain}}
.band .titles{{text-align:center;flex:1}}
.kicker{{color:{GOLD};font-size:10.5px;letter-spacing:2.5px;text-transform:uppercase;font-weight:700}}
.opp{{font-size:26px;font-weight:800;letter-spacing:.5px;text-transform:uppercase;margin:2px 0}}
.sub{{font-size:12px;color:#ddd}}
.tiles{{display:flex;justify-content:center;gap:12px;margin:14px 0 12px}}
.tile{{border:1.5px solid {GOLD};border-radius:6px;padding:6px 16px;text-align:center;min-width:1.6in}}
.tile .v{{font-size:22px;font-weight:800}}
.tile .l{{font-size:9.5px;letter-spacing:1.5px;text-transform:uppercase;color:{GOLD_DARK};font-weight:700}}
.tile .s{{font-size:9.5px;color:#888}}
table{{border-collapse:collapse;margin:0 auto;font-size:13.5px}}
th{{background:{BLACK};color:{GOLD};font-size:11px;letter-spacing:.8px;text-transform:uppercase;
  padding:8px 10px;border:1px solid {BLACK}}}
td{{padding:6.5px 10px;border:1px solid #ddd;text-align:center}}
tbody tr:nth-child(even) td{{background:#f7f4ee}}
td.l{{text-align:left;font-weight:600}}
td.num{{font-weight:800;color:{GOLD_DARK}}}
.dim{{color:#999;font-size:10.5px}}
tbody tr td.good{{background:#cdebd0;color:#14532d;font-weight:800}}
tbody tr td.bad{{background:#f5d0d0;color:#7f1d1d}}
.nodata{{max-width:6.8in;margin:10px auto 0;font-size:11px;color:#555;text-align:center}}
.legend{{display:flex;justify-content:center;flex-wrap:wrap;gap:6px 16px;margin:14px auto 0;font-size:10.5px;color:#444}}
.chip{{display:inline-block;width:12px;height:12px;border-radius:2px;vertical-align:-2px;margin-right:4px}}
.foot{{text-align:center;font-size:9.5px;color:#999;margin-top:8px;border-top:1px solid {GOLD};padding-top:5px}}
.cols{{display:flex;justify-content:center;gap:56px}}
.cols th{{font-size:12px;padding:9px 18px}}
.cols td{{font-size:15px;padding:0 24px;height:27.5px}}
.cols tr.ten td{{border-top:2px solid {GOLD}}}
@page{{size:letter;margin:0.4in}}
@media print{{.sheet{{padding:0;break-after:page}} .sheet:last-child{{break-after:auto}} .noprint{{display:none}}}}
"""

_LOGO_CACHE = {}

# Teams ESPN has no logo for (ESPN team id -> image URL)
LOGO_OVERRIDES = {
    "2679": "https://voorheesathletics.com/images/logos/site/site.png?width=120",   # Voorhees
}


def logo_url(team_id):
    return LOGO_OVERRIDES.get(team_id, f"https://a.espncdn.com/i/teamlogos/ncaa/500/{team_id}.png")


def logo_data_uri(team_id):
    """ESPN team logo as an embedded image (so PDFs never miss it). None if unavailable."""
    if team_id not in _LOGO_CACHE:
        uri = None
        try:
            r = get(LOGO_OVERRIDES.get(team_id) or
                    f"https://a.espncdn.com/combiner/i?img=/i/teamlogos/ncaa/500/{team_id}.png&h=120&w=120")
            ctype = r.headers.get("content-type", "").split(";")[0] if r is not None else ""
            if ctype.startswith("image"):
                uri = f"data:{ctype};base64," + base64.b64encode(r.content).decode()
        except Exception:
            pass
        _LOGO_CACHE[team_id] = uri
    return _LOGO_CACHE[team_id]


def _logo(team_id):
    uri = logo_data_uri(team_id)
    if not uri:
        return "<div class='logo' style='visibility:hidden'></div>"   # keeps the title centered
    return f"<div class='logo'><img src='{uri}'></div>"


def f3(x):
    return "--" if x is None or pd.isna(x) else f"{x:.3f}"


def pct(x):
    return "--" if x is None or pd.isna(x) else f"{x * 100:.1f}%"


def cmp_class(epp, ppp):
    if epp is None or ppp is None or pd.isna(epp):
        return ""
    return "good" if epp < ppp else "bad"


# Team sheets print landscape: stat boxes + key in a left column, roster table on the right.
LANDSCAPE_CSS = f"""
@page{{{{size:letter landscape;margin:0.35in}}}}
.sheet{{{{max-width:10.2in}}}}
.body{{{{display:flex;gap:20px;align-items:flex-start;justify-content:center;margin-top:14px}}}}
.side{{{{width:2.35in;flex:none}}}}
.side .tiles{{{{flex-direction:column;margin:0;gap:9px}}}}
.side .tile{{{{min-width:0;padding:5px 10px}}}}
.side .nodata{{{{text-align:left;margin:12px 0 0;font-size:10.5px}}}}
.side .legend{{{{flex-direction:column;align-items:flex-start;margin:12px 0 0;gap:5px}}}}
.main table td{{{{padding:0 12px;height:var(--rh,30px)}}}}
@media screen and (max-width:760px){{{{.body{{{{flex-direction:column;align-items:stretch}}}} .side{{{{width:auto}}}}
  .side .tiles{{{{flex-direction:row}}}} .main{{{{overflow-x:auto}}}}}}}}
""".replace("{{", "{").replace("}}", "}")


def _page(title, body, extra_css=""):
    return (f"<!DOCTYPE html><html><head><meta charset='utf-8'>"
            f"<meta name='viewport' content='width=device-width,initial-scale=1'>"
            f"<title>{html.escape(title)}</title><style>{CSS}{extra_css}</style></head><body>{body}</body></html>")


def _legend():
    return ("<div class='legend'>"
            "<span><span class='chip' style='background:#cdebd0;border:1px solid #14532d'></span>"
            "Below team PPP &mdash; fouling favorable</span>"
            "<span><span class='chip' style='background:#f5d0d0;border:1px solid #7f1d1d'></span>"
            "Above team PPP &mdash; don't foul</span>"
            "<span>* small sample (&lt;10 career FTA)</span></div>")


def _team_sheet(t, roster, season_label, updated):
    ppp = t["ppp"]
    have = sorted([p for p in roster if p["ft_pct"] is not None], key=lambda p: p["ft_pct"])
    none = [p for p in roster if p["ft_pct"] is None]
    n11 = sum(1 for p in have if cmp_class(p["epp_1and1"], ppp) == "good")
    n2 = sum(1 for p in have if cmp_class(p["epp_2shot"], ppp) == "good")

    out = ["<section class='sheet'>",
           f"<header class='band'>{_logo(WOFFORD_ESPN_ID)}<div class='titles'>"
           "<div class='kicker'>Wofford Basketball &middot; Free Throw Foul Sheet</div>"
           f"<div class='opp'>{html.escape(t['name'])}</div>"
           f"<div class='sub'>{' &nbsp;&middot;&nbsp; '.join(t['games'])}</div></div>{_logo(t['id'])}</header>",
           "<div class='body'><aside class='side'><div class='tiles'>"
           f"<div class='tile'><div class='l'>Points / Possession</div><div class='v'>{f3(ppp)}</div>"
           f"<div class='s'>{t['ppp_season'] if ppp else 'not available'}</div></div>"
           f"<div class='tile'><div class='l'>Foul on 1-and-1</div><div class='v'>{n11 if ppp else '--'}</div>"
           "<div class='s'>players</div></div>"
           f"<div class='tile'><div class='l'>Foul on 2 shots</div><div class='v'>{n2 if ppp else '--'}</div>"
           "<div class='s'>players</div></div></div>"]
    if none:
        names = ", ".join(f"#{html.escape(p['jersey'])} {html.escape(p['name'])}" for p in none)
        out.append(f"<div class='nodata'><b>No free throw data yet:</b> {names}</div>")
    out.append(_legend() + "</aside><div class='main'>")

    if have:
        rh = max(28, min(62, 545 // len(have)))          # px per row: short rosters get taller rows
        fs = 17 if rh >= 52 else 15.5 if rh >= 40 else 14.5 if rh >= 33 else 13.5
        out.append(f"<table style='--rh:{rh}px;font-size:{fs}px'><thead><tr><th>#</th><th>Player</th><th>Pos</th><th>Yr</th>"
                   f"<th>{season_label}</th><th>Career</th><th>FT%</th>"
                   "<th>1-and-1</th><th>2 Shots</th></tr></thead><tbody>")
        for p in have:
            flag = "*" if p["basis"] == "Career" and p["fta_used"] < 10 else ""
            out.append(
                f"<tr><td class='num'>{html.escape(p['jersey'])}</td><td class='l'>{html.escape(p['name'])}</td>"
                f"<td>{p['pos']}</td><td>{p['class']}</td>"
                f"<td>{p['season_ft']} <span class='dim'>{pct(p['season_pct']) if p['season_pct'] else ''}</span></td>"
                f"<td>{p['career_ft']} <span class='dim'>{pct(p['career_pct'])}</span></td>"
                f"<td><b>{pct(p['ft_pct'])}</b>{flag}</td>"
                f"<td class='{cmp_class(p['epp_1and1'], ppp)}'>{f3(p['epp_1and1'])}</td>"
                f"<td class='{cmp_class(p['epp_2shot'], ppp)}'>{f3(p['epp_2shot'])}</td></tr>")
        out.append("</tbody></table>")
    elif not roster:
        out.append("<p class='nodata'>No roster available on ESPN.</p>")
    out.append("</div></div>")
    out.append(f"<div class='foot'>Updated {updated}</div></section>")
    return "".join(out)


def _safe(name):
    return re.sub(r"[^A-Za-z0-9 ]+", "", name).strip().replace(" ", "_")


def write_report(season_label, teams, players, min_fta, min_games):
    """Combined HTML + one HTML per opponent. Returns (combined_path, [per-team paths])."""
    updated = now_eastern().strftime("%b %d, %Y %I:%M %p ET")
    by_team = {}
    for p in players:
        by_team.setdefault(p["team_id"], []).append(p)

    team_dir = OUT_DIR / "teams"
    team_dir.mkdir(exist_ok=True)
    sheets, team_paths = [], []
    for t in teams:
        sheet = _team_sheet(t, by_team.get(t["id"], []), season_label, updated)
        sheets.append(sheet)
        tp = team_dir / f"{_safe(t['name'])}.html"
        tp.write_text(_page(f"Wofford FT Sheet - {t['name']}", sheet, LANDSCAPE_CSS), encoding="utf-8")
        team_paths.append(tp)
    path = OUT_DIR / "Opponent_FT_Report.html"
    path.write_text(_page("Wofford FT Sheets - All Opponents", "".join(sheets), LANDSCAPE_CSS),
                    encoding="utf-8")
    return path, team_paths


BROWSERS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
]


def html_to_pdf(html_path):
    """Print an HTML file to PDF (same folder, same name) with headless Edge/Chrome."""
    browser = next((b for b in BROWSERS if Path(b).exists()), None) or next(
        (shutil.which(n) for n in ("google-chrome", "chromium", "chromium-browser", "msedge")
         if shutil.which(n)), None)
    if not browser:
        raise RuntimeError("Edge or Chrome is needed to make PDFs")
    pdf = html_path.with_suffix(".pdf")
    subprocess.run([browser, "--headless", "--disable-gpu", "--no-sandbox", "--no-pdf-header-footer",
                    "--print-to-pdf-no-header", f"--print-to-pdf={pdf}", html_path.as_uri()],
                   check=True, capture_output=True, timeout=120)
    if not pdf.exists():
        raise RuntimeError(f"PDF was not created for {html_path.name}")
    return pdf


def generic_rows():
    return [(p, one_and_one(p / 100), two_shots(p / 100)) for p in range(40, 100)]


def write_generic():
    rows = generic_rows()
    half = len(rows) // 2

    def table(chunk):
        t = ["<table><thead><tr><th>FT%</th><th>1-and-1</th><th>2 Shots</th></tr></thead><tbody>"]
        for p, a, b in chunk:
            cls = " class='ten'" if p % 10 == 0 and p not in (40, 70) else ""
            t.append(f"<tr{cls}><td class='num'>{p}%</td><td>{a:.3f}</td><td>{b:.3f}</td></tr>")
        t.append("</tbody></table>")
        return "".join(t)

    body = ("<section class='sheet'>"
            f"<header class='band'>{_logo(WOFFORD_ESPN_ID)}<div class='titles'>"
            "<div class='kicker'>Wofford Basketball &middot; Free Throw Cheat Sheet</div>"
            "<div class='opp'>Expected Points per Trip</div>"
            "<div class='sub'>Foul when the shooter's number is below the opponent's points per possession</div>"
            f"</div>{_logo(WOFFORD_ESPN_ID)}</header>"
            "<div style='height:14px'></div>"
            f"<div class='cols'>{table(rows[:half])}{table(rows[half:])}</div></section>")
    path = OUT_DIR / "Generic_FT_Cheat_Sheet.html"
    path.write_text(_page("Wofford FT Cheat Sheet", body), encoding="utf-8")
    return path


INDEX_CSS = f"""
.wrap{{max-width:820px;margin:0 auto;padding:16px}}
.buttons{{display:flex;justify-content:center;flex-wrap:wrap;gap:10px;margin:16px 0}}
.buttons a{{background:{BLACK};color:{GOLD};text-decoration:none;font-weight:700;padding:10px 18px;
  border-radius:6px;border:2px solid {GOLD}}}
.sched{{width:100%}}
.sched td{{font-size:14px;padding:7px 8px}}
.sched img{{width:26px;height:26px;object-fit:contain;vertical-align:middle;margin-right:8px}}
.sched a{{color:{GOLD_DARK};font-weight:800;text-decoration:none}}
.sched tr.past td{{opacity:.45}}
.sched tr.next td{{background:#fff4dc !important;font-weight:700}}
.sched .note{{font-size:11.5px;font-style:italic;font-weight:400;color:{GOLD_DARK};margin:2px 0 0 34px}}
"""

# Notes shown under an opponent's name on the schedule page (ESPN team id -> text).
TEAM_NOTES = {
    "3229": "NCCAA - stats not available until closer to game day",   # Virginia-Lynchburg (site blocks automated access)
}


def write_index(season_label, games, teams, team_paths):
    """Landing page: schedule in order with links to each opponent's sheet."""
    link = {t["id"]: p for t, p in zip(teams, team_paths)}
    ppp = {t["id"]: t["ppp"] for t in teams}
    today = now_eastern().date().isoformat()
    next_date = next((g["date"] for g in games if g["date"] >= today), None)
    rows = []
    for g in games:
        href = quote(f"teams/{link[g['opp_id']].stem}")
        cls = "past" if g["date"] < today else ("next" if g["date"] == next_date else "")
        rows.append(
            f"<tr class='{cls}'><td>{g['tip'].strftime('%a %b')} {g['tip'].day}</td><td>{g['site']}</td>"
            f"<td class='l'><img src='{logo_url(g['opp_id'])}' alt='' onerror=\"this.style.visibility='hidden'\">"
            f"{html.escape(g['opp_name'])}"
            + (f"<div class='note'>{html.escape(TEAM_NOTES[g['opp_id']])}</div>" if g["opp_id"] in TEAM_NOTES else "")
            + f"</td><td>{f3(ppp[g['opp_id']])}</td>"
            f"<td><a href='{href}.pdf'>PDF</a> &nbsp;&middot;&nbsp; <a href='{href}.html'>View</a></td></tr>")
    body = (f"<style>{INDEX_CSS}</style><div class='wrap'>"
            f"<header class='band'>{_logo(WOFFORD_ESPN_ID)}<div class='titles'>"
            f"<div class='kicker'>Wofford Men's Basketball &middot; {season_label}</div>"
            "<div class='opp'>Free Throw Foul Sheets</div>"
            f"<div class='sub'>Updated {now_eastern().strftime('%a %b %d, %I:%M %p')} ET &middot; refreshes every morning</div>"
            f"</div>{_logo(WOFFORD_ESPN_ID)}</header>"
            "<div class='buttons'><a href='Generic_FT_Cheat_Sheet.pdf'>Generic 40&ndash;99% Sheet</a>"
            "<a href='Opponent_FT_Report.pdf'>All Opponents (PDF)</a></div>"
            "<table class='sched'><thead><tr><th>Date</th><th></th><th>Opponent</th><th>PPP</th>"
            "<th>Sheet</th></tr></thead><tbody>" + "".join(rows) + "</tbody></table></div>")
    path = OUT_DIR / "index.html"
    path.write_text(_page("Wofford FT Foul Sheets", body), encoding="utf-8")
    return path


def write_excel(teams, players):
    path = OUT_DIR / "FT_Data.xlsx"
    with pd.ExcelWriter(path) as xw:
        pd.DataFrame(generic_rows(), columns=["FT%", "1-and-1 pts/poss", "2-shot pts/poss"]) \
            .to_excel(xw, sheet_name="Generic", index=False)
        pd.DataFrame([{
            "Team": t["name"], "TeamRankings name": t["tr_name"], "Games vs Wofford": ", ".join(t["games"]),
            "PPG": t["ppg"], "Poss/G": t["pace"], "PPP": t["ppp"], "PPP season": t["ppp_season"], "Games played": t["games_played"],
        } for t in teams]).to_excel(xw, sheet_name="Teams", index=False)
        cols = ["team", "jersey", "name", "pos", "class", "season_ft", "season_pct", "career_ft",
                "career_pct", "ft_pct", "basis", "epp_1and1", "epp_2shot", "team_ppp"]
        pd.DataFrame(players)[cols].sort_values(["team", "ft_pct"]) \
            .to_excel(xw, sheet_name="Players", index=False)
    return path


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--min-fta", type=int, default=20,
                    help="FTA this season before using season FT%% instead of career (default 20)")
    ap.add_argument("--min-games", type=int, default=8,
                    help="team games this season before using this season's PPP (default 8)")
    ap.add_argument("--no-excel", action="store_true", help="skip FT_Data.xlsx (used for the website)")
    ap.add_argument("--out", help="folder to write everything to (default: this script's folder)")
    args = ap.parse_args()

    global OUT_DIR
    if args.out:
        OUT_DIR = Path(args.out).resolve()
        OUT_DIR.mkdir(parents=True, exist_ok=True)

    print("Wrote", write_generic())
    season_label, games, teams, players = build(args.min_fta, args.min_games)
    combined, team_paths = write_report(season_label, teams, players, args.min_fta, args.min_games)
    print("Wrote", combined)
    print("Making PDFs ...")
    for p in [OUT_DIR / "Generic_FT_Cheat_Sheet.html", combined, *team_paths]:
        try:
            print("  Wrote", html_to_pdf(p))
        except Exception as e:  # keep going; HTML is still usable
            print(f"  PDF failed for {p.name}: {e}")
    try:
        if not args.no_excel:
            print("Wrote", write_excel(teams, players))
    except PermissionError:
        print("Could not write FT_Data.xlsx - close it in Excel and re-run.")
    print("Wrote", write_index(season_label, games, teams, team_paths))


if __name__ == "__main__":
    main()
