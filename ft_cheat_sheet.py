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
import concurrent.futures as cf
import datetime as dt
import difflib
import html
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
def get_schedule():
    d = get(f"{ESPN_SITE}/teams/{WOFFORD_ESPN_ID}/schedule").json()
    season_label = d["season"]["displayName"]          # e.g. "2026-27"
    season_year = d["season"]["year"]                  # e.g. 2027
    games = []
    for e in d.get("events", []):
        comp = e["competitions"][0]
        opp = next(c for c in comp["competitors"] if c["team"]["id"] != WOFFORD_ESPN_ID)
        home = next(c for c in comp["competitors"] if c["team"]["id"] == WOFFORD_ESPN_ID)["homeAway"]
        games.append({
            "date": e["date"][:10],
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
        opponents[g["opp_id"]]["games"].append(f'{g["site"]} {g["date"]}')

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
        team_rows.append(opp)
        print(f'  {opp["name"]:<35} -> TR: {opp["tr_name"] or "NOT FOUND":<18} '
              f'PPP {opp["ppp"] if opp["ppp"] is not None else "--"} ({opp["ppp_season"]}, {gp} GP)')

    print("Pulling rosters and free throw stats from ESPN ...")
    player_rows = []
    jobs = []
    with cf.ThreadPoolExecutor(max_workers=6) as pool:
        for opp in team_rows:
            for p in get_roster(opp["id"]):
                jobs.append((opp, p, pool.submit(get_player_ft, p["id"], season_label)))
        for opp, p, fut in jobs:
            s_m, s_a, c_m, c_a = fut.result()
            if s_a >= min_fta:
                pct, basis, fta_used = s_m / s_a, season_label, s_a
            elif c_a > 0:
                pct, basis, fta_used = c_m / c_a, "Career", c_a
            else:
                pct, basis, fta_used = None, "No D1 FT data", 0
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
CSS = """
body{font-family:Arial,Helvetica,sans-serif;margin:24px;color:#111}
h1{margin:0 0 4px;color:#000}
h2{margin:28px 0 4px;padding:6px 10px;background:#111;color:#c8a464;font-size:18px}
.meta{font-size:12px;color:#444;margin-bottom:8px}
table{border-collapse:collapse;font-size:12px}
th,td{border:1px solid #bbb;padding:3px 7px;text-align:center}
th{background:#eee}
td.l{text-align:left}
.good{background:#c9efc9;font-weight:bold}   /* fouling favorable */
.bad{background:#f6c8c8}                     /* fouling unfavorable */
.small{color:#888}
.key{font-size:12px;margin:8px 0 16px}
.key span{padding:2px 8px;margin-right:8px;border:1px solid #bbb}
.cols{display:flex;gap:24px}
@page{size:letter;margin:10mm}
@media print{body{margin:0} h2{break-before:page} h1+div+div+div+h2{break-before:avoid}
  .noprint{display:none} *{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
"""


def f3(x):
    return "--" if x is None or pd.isna(x) else f"{x:.3f}"


def pct(x):
    return "--" if x is None or pd.isna(x) else f"{x * 100:.1f}%"


def cmp_class(epp, ppp):
    if epp is None or ppp is None or pd.isna(epp):
        return ""
    return "good" if epp < ppp else "bad"


def _team_section(t, roster, season_label):
    ppp = t["ppp"]
    roster = sorted(roster, key=lambda p: (p["ft_pct"] is None, p["ft_pct"] if p["ft_pct"] is not None else 0))
    out = [f"<h2>{html.escape(t['name'])}</h2>",
           f"<div class='meta'>Games: {', '.join(t['games'])} &nbsp;|&nbsp; "
           f"<b>Points per possession {f3(ppp)}</b> ({t['ppp_season']})</div>"]
    if not roster:
        out.append("<p class='small'>No roster on ESPN.</p>")
        return out
    out.append("<table><tr><th>#</th><th>Player</th><th>Pos</th><th>Yr</th>"
               f"<th>{season_label} FT</th><th>Career FT</th><th>FT% used</th><th>Basis</th>"
               "<th>1-and-1 pts</th><th>2-shot pts</th></tr>")
    for p in roster:
        flag = "*" if p["basis"] == "Career" and p["fta_used"] < 10 else ""
        out.append(
            f"<tr><td>{html.escape(p['jersey'])}</td><td class='l'>{html.escape(p['name'])}</td>"
            f"<td>{p['pos']}</td><td>{p['class']}</td>"
            f"<td>{p['season_ft']} <span class='small'>{pct(p['season_pct'])}</span></td>"
            f"<td>{p['career_ft']} <span class='small'>{pct(p['career_pct'])}</span></td>"
            f"<td><b>{pct(p['ft_pct'])}</b>{flag}</td><td>{p['basis']}</td>"
            f"<td class='{cmp_class(p['epp_1and1'], ppp)}'>{f3(p['epp_1and1'])}</td>"
            f"<td class='{cmp_class(p['epp_2shot'], ppp)}'>{f3(p['epp_2shot'])}</td></tr>")
    out.append("</table>")
    return out


def _report_header(title, season_label, min_fta, min_games):
    now = dt.datetime.now().strftime("%Y-%m-%d %I:%M %p")
    return [f"<!DOCTYPE html><html><head><meta charset='utf-8'><title>{html.escape(title)}</title>"
            f"<style>{CSS}</style></head><body>",
            f"<h1>{html.escape(title)}</h1>",
            f"<div class='meta'>Updated {now}</div>",
            "<div class='key'><span class='good'>Green</span>expected points from the line are below team PPP "
            "(fouling is favorable) <span class='bad'>Red</span>above team PPP (don't foul)</div>",
            "<div class='key'>* = fewer than 10 career FTA (small sample) &nbsp;&nbsp; "
            "No D1 FT data = ESPN has no college stats for the player yet (usually a newcomer)</div>"]


def _safe(name):
    return re.sub(r"[^A-Za-z0-9 ]+", "", name).strip().replace(" ", "_")


def write_report(season_label, teams, players, min_fta, min_games):
    """Combined HTML + one HTML per opponent. Returns (combined_path, [per-team paths])."""
    by_team = {}
    for p in players:
        by_team.setdefault(p["team_id"], []).append(p)

    combined = _report_header("Wofford Men's Basketball - Opponent Free Throw Report",
                              season_label, min_fta, min_games)
    team_dir = OUT_DIR / "teams"
    team_dir.mkdir(exist_ok=True)
    team_paths = []
    for t in teams:
        section = _team_section(t, by_team.get(t["id"], []), season_label)
        combined += section
        doc = (_report_header(f"Wofford FT Report - {t['name']}", season_label, min_fta, min_games)
               + section + ["</body></html>"])
        tp = team_dir / f"{_safe(t['name'])}.html"
        tp.write_text("\n".join(doc), encoding="utf-8")
        team_paths.append(tp)
    combined.append("</body></html>")
    path = OUT_DIR / "Opponent_FT_Report.html"
    path.write_text("\n".join(combined), encoding="utf-8")
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
        t = ["<table><tr><th>FT%</th><th>1-and-1<br>pts/poss</th><th>2 shots<br>pts/poss</th></tr>"]
        for p, a, b in chunk:
            t.append(f"<tr><td><b>{p}%</b></td><td>{a:.3f}</td><td>{b:.3f}</td></tr>")
        t.append("</table>")
        return "".join(t)

    doc = (f"<!DOCTYPE html><html><head><meta charset='utf-8'><title>FT Cheat Sheet</title>"
           f"<style>{CSS} td,th{{font-size:13px;padding:2px 10px}}</style></head><body>"
           "<h1>Free Throw Expected Points per Possession</h1>"
           "<div class='meta'>Foul is favorable when the shooter's number is below the opponent's "
           "points per possession.</div>"
           f"<div class='cols'>{table(rows[:half])}{table(rows[half:])}</div></body></html>")
    path = OUT_DIR / "Generic_FT_Cheat_Sheet.html"
    path.write_text(doc, encoding="utf-8")
    return path


def now_eastern():
    try:
        from zoneinfo import ZoneInfo
        return dt.datetime.now(ZoneInfo("America/New_York"))
    except Exception:          # Windows without tzdata: local clock is Eastern anyway
        return dt.datetime.now()


def write_index(season_label, games, teams, team_paths):
    """Landing page: schedule in order with links to each opponent's report."""
    link = {t["id"]: p for t, p in zip(teams, team_paths)}
    ppp = {t["id"]: t["ppp"] for t in teams}
    today = now_eastern().date().isoformat()
    rows = []
    for g in games:
        p = link[g["opp_id"]]
        href = quote(f"teams/{p.stem}")
        cls = " class='past'" if g["date"] < today else ""
        rows.append(f"<tr{cls}><td>{g['date']}</td><td>{g['site']}</td>"
                    f"<td class='l'>{html.escape(g['opp_name'])}</td><td>{f3(ppp[g['opp_id']])}</td>"
                    f"<td><a href='{href}.pdf'>PDF</a> &middot; <a href='{href}.html'>View</a></td></tr>")
    doc = (f"<!DOCTYPE html><html><head><meta charset='utf-8'>"
           f"<meta name='viewport' content='width=device-width,initial-scale=1'>"
           f"<title>Wofford FT Cheat Sheets</title><style>{CSS}"
           " a{color:#7a5c1e;font-weight:bold} tr.past td{color:#999} tr.past a{color:#b9a27a}"
           " .big a{display:inline-block;margin:4px 8px 4px 0;padding:8px 14px;background:#111;color:#c8a464;"
           "text-decoration:none;border-radius:4px} td,th{font-size:14px;padding:5px 9px}</style></head><body>"
           f"<h1>Wofford Men's Basketball &mdash; Free Throw Cheat Sheets</h1>"
           f"<div class='meta'>{season_label} &middot; Updated automatically every morning. "
           f"Last update: {now_eastern().strftime('%a %b %d, %Y %I:%M %p')} ET</div>"
           "<div class='big'><a href='Generic_FT_Cheat_Sheet.pdf'>Generic 40&ndash;99% sheet (PDF)</a>"
           "<a href='Opponent_FT_Report.pdf'>All opponents (PDF)</a></div>"
           "<h2>Schedule</h2><table><tr><th>Date</th><th></th><th>Opponent</th><th>Team PPP</th>"
           "<th>Report</th></tr>" + "".join(rows) + "</table></body></html>")
    path = OUT_DIR / "index.html"
    path.write_text(doc, encoding="utf-8")
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
