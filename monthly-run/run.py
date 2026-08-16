#!/usr/bin/env python3
"""
run.py - Monthly Sprint Room orchestrator.

Runs one carline's monthly build as a sequence of stages, calling the skill's own
scripts with the right arguments and per-carline config. It deliberately STOPS at the
bullet-menu approval gate and never auto-applies the manual OOXML fixes - those are
surfaced as explicit checklists (see `fixskeleton`). Nothing here does arithmetic on
the data; the skill scripts do.

Conventions
-----------
Inputs  : monthly-run/inputs/<CARLINE>/<YYYY-MM>/
            last_month.pptx  raw.xlsx  chart_volume.xlsx  chart_country.xlsx
            chart_platform.xlsx  chart_sentiment.xlsx
Work    : monthly-run/work/<CARLINE>/<YYYY-MM>/   (skeleton/, menu.md, output.pptx)

Typical cycle
-------------
    python run.py newmonth  --carline Kuga --month 2026-07
    # (drop the six exports into the inputs folder it printed)
    python run.py doctor    --carline Kuga --month 2026-07
    python run.py validate  --carline Kuga --month 2026-07
    python run.py menu      --carline Kuga --month 2026-07
    #   ---- HARD GATE: review & approve work/.../menu.md ----
    python run.py skeleton  --carline Kuga --month 2026-07
    python run.py fixskeleton --carline Kuga --month 2026-07   # manual checklist
    python run.py roles     --carline Kuga --month 2026-07     # confirm chart mapping
    python run.py charts    --carline Kuga --month 2026-07
    python run.py bullets   --carline Kuga --month 2026-07
    python run.py embed     --carline Kuga --month 2026-07
    python run.py pack      --carline Kuga --month 2026-07

Add --dry-run to any stage to print the commands without executing them.
Override the skill location with  SKILL_DIR=/mnt/skills/user/sprint-room-monthly-update
"""
import argparse, json, os, re, subprocess, sys, zipfile
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
SKILL_DIR = Path(os.environ.get(
    "SKILL_DIR", REPO / "skills" / "sprint-room-monthly-update"))
SCRIPTS = SKILL_DIR / "scripts"
CONFIG = json.loads((HERE / "config" / "carlines.json").read_text())

INPUT_FILES = ["last_month.pptx", "raw.xlsx", "chart_volume.xlsx",
               "chart_country.xlsx", "chart_platform.xlsx", "chart_sentiment.xlsx"]
MONTHS = ["", "January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

C = {"g": "\033[32m", "y": "\033[33m", "r": "\033[31m", "b": "\033[1m", "x": "\033[0m"}
def say(msg, col=None): print(f"{C.get(col,'')}{msg}{C['x'] if col else ''}")


# ---------- helpers ----------
def cfg(carline):
    if carline not in CONFIG:
        sys.exit(f"Unknown carline {carline!r}. Known: {', '.join(k for k in CONFIG if not k.startswith('_'))}")
    return CONFIG[carline]

def paths(carline, month):
    inp = HERE / "inputs" / carline / month
    work = HERE / "work" / carline / month
    return inp, work

def month_parts(month):
    m = re.fullmatch(r"(\d{4})-(\d{2})", month)
    if not m:
        sys.exit(f"--month must be YYYY-MM, got {month!r}")
    yr, mm = int(m.group(1)), int(m.group(2))
    prior = 12 if mm == 1 else mm - 1
    return yr, mm, prior

def run(cmd, dry):
    printable = " ".join(str(c) for c in cmd)
    if dry:
        say("  DRY  " + printable, "y"); return 0
    say("  RUN  " + printable, "b")
    return subprocess.call([str(c) for c in cmd])

def need(path, what):
    if not Path(path).exists():
        sys.exit(f"Missing {what}: {path}\n  (run earlier stages first, or `doctor` to check inputs)")


# ---------- stages ----------
def stage_newmonth(a):
    inp, work = paths(a.carline, a.month)
    inp.mkdir(parents=True, exist_ok=True); work.mkdir(parents=True, exist_ok=True)
    (inp / "PUT_THESE_SIX_FILES_HERE.txt").write_text(
        "Drop these exact filenames into this folder:\n" +
        "\n".join(f"  - {f}" for f in INPUT_FILES) +
        "\n\nThen: python run.py doctor --carline %s --month %s\n" % (a.carline, a.month))
    c = cfg(a.carline)
    say(f"Scaffolded {a.carline} {a.month}", "g")
    say(f"  inputs : {inp}")
    say(f"  work   : {work}")
    say(f"  segment: {c['segment']}   competitors ({len(c['competitors'])}): {', '.join(c['competitors'])}")
    if c.get("verify_chart_roles"):
        say("  NOTE   : chart roles for this carline need verification - run `roles` after `skeleton`.", "y")
    say("Next: drop the six exports in, then `doctor`.")

def stage_doctor(a):
    inp, _ = paths(a.carline, a.month)
    ok = True
    say(f"Checking inputs for {a.carline} {a.month} in {inp}")
    for f in INPUT_FILES:
        p = inp / f
        mark = f"{C['g']}OK {C['x']}" if p.exists() else f"{C['r']}MISS{C['x']}"
        size = f"{p.stat().st_size/1024:.0f} KB" if p.exists() else ""
        print(f"  [{mark}] {f} {size}")
        ok = ok and p.exists()
    say("All inputs present." if ok else "Some inputs missing - add them before `validate`.",
        "g" if ok else "r")
    return 0 if ok else 1

def stage_validate(a):
    inp, _ = paths(a.carline, a.month)
    cmd = ["python3", SCRIPTS / "validate_inputs.py", "--carline", a.carline,
           "--raw", inp / "raw.xlsx", "--vol", inp / "chart_volume.xlsx",
           "--country", inp / "chart_country.xlsx", "--platform", inp / "chart_platform.xlsx",
           "--sent", inp / "chart_sentiment.xlsx", "--last-month", inp / "last_month.pptx"]
    return run(cmd, a.dry_run)

def stage_menu(a):
    inp, work = paths(a.carline, a.month)
    work.mkdir(parents=True, exist_ok=True)
    _, mm, prior = month_parts(a.month)
    menu = work / "menu.md"
    cmd = ["python3", SCRIPTS / "extract_drivers.py", "--carline", a.carline,
           "--raw", inp / "raw.xlsx", "--new-month", mm, "--prior-month", prior, "--out", menu]
    rc = run(cmd, a.dry_run)
    if rc == 0 and not a.dry_run:
        say("\n" + "="*70, "y")
        say("HARD GATE: open the menu, check clusters/counts/URLs, edit, and APPROVE", "y")
        say(f"   {menu}", "y")
        say("Nothing gets built into the deck until you're happy with it. Then run `skeleton`.", "y")
        say("="*70, "y")
    return rc

def stage_skeleton(a):
    inp, work = paths(a.carline, a.month)
    yr, mm, _ = month_parts(a.month)
    report_date = a.report_date or date.today().strftime("%d.%m.%Y")
    out = work / "skeleton"
    cmd = ["python3", SCRIPTS / "build_skeleton.py", "--last-month", inp / "last_month.pptx",
           "--carline", a.carline, "--new-month", f"{MONTHS[mm]} {yr}",
           "--report-date", report_date, "--out", out]
    rc = run(cmd, a.dry_run)
    if rc == 0 and not a.dry_run:
        say("\nNow run `fixskeleton` - build_skeleton has known per-cycle bugs to fix by hand.", "y")
    return rc

def stage_fixskeleton(a):
    c = cfg(a.carline)
    say("Manual build_skeleton fixes (see docs/tooling-patterns.md):", "b")
    items = [
        "Remove the orphaned <Relationship> for the dropped older slide from "
        "ppt/_rels/presentation.xml.rels",
        "Remove stale [Content_Types].xml overrides for the deleted chart files + slide XML",
        "Recreate the new slide's notesSlide: notesSlideN.xml, its .rels, the "
        "[Content_Types].xml override, and re-insert the rId2 relationship in the slide's rels",
    ]
    if c["segment"] == "EU5":
        items.append(f"EU5 carline: build_skeleton hardcodes 'EU7' in the subtitle regex - "
                     f"set the subtitle to '{a.carline} - Social Media ({c['segment']} "
                     f"{MONTHS[month_parts(a.month)[1]]} {month_parts(a.month)[0]})' manually")
    if a.carline == "Kuga":
        items.append("Kuga: sentiment series names are Capitalised ('Positive Pre'/'Negative Pre')")
    for i, it in enumerate(items, 1):
        print(f"  {i}. {it}")
    say("These are not auto-applied on purpose - verify each in the unpacked skeleton.", "y")
    return 0

def _chart_categories(chart_xml):
    """Return the category strings of the first series in a chart (for role detection)."""
    try:
        data = Path(chart_xml).read_bytes()
    except FileNotFoundError:
        return []
    cats = re.findall(rb"<c:pt[^>]*>\s*<c:v>([^<]*)</c:v>", data)
    return [x.decode("utf-8", "replace") for x in cats]

def _infer_role(cats):
    joined = " ".join(cats).lower()
    if any(w in joined for w in ["positive", "negative", "neutral"]): return "sentiment"
    if any(w in joined for w in ["news", "forum", "blog", "tumblr", "review", " x "]): return "platform"
    if re.search(r"\b(20\d\d|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|month)", joined):
        return "volume"
    if any(w in joined for w in ["germany", "united kingdom", "spain", "france", "italy",
                                 "netherlands", "norway"]): return "country"
    return "unknown"

def stage_roles(a):
    _, work = paths(a.carline, a.month)
    charts = work / "skeleton" / "unpacked" / "ppt" / "charts"
    if not charts.exists():
        sys.exit(f"No unpacked skeleton at {charts} - run `skeleton` first.")
    c = cfg(a.carline)
    say(f"Chart-role detection for {a.carline} (config vs. what's in the skeleton):", "b")
    inferred = {}
    for xml in sorted(charts.glob("chart*.xml"),
                      key=lambda p: int(re.search(r"chart(\d+)", p.name).group(1))):
        n = xml.stem
        role = _infer_role(_chart_categories(xml))
        inferred[n] = role
    mismatch = False
    for role, chartn in c["chart_roles"].items():
        got = inferred.get(chartn, "(missing)")
        ok = (got == role)
        mismatch = mismatch or not ok
        mark = f"{C['g']}OK {C['x']}" if ok else f"{C['r']}!! {C['x']}"
        print(f"  [{mark}] config says {chartn} = {role:9s} | skeleton looks like: {got}")
    if mismatch or c.get("verify_chart_roles"):
        say("Mismatch or verification flagged. Update config/carlines.json chart_roles "
            "for this carline before `charts`.", "y")
    else:
        say("Config matches the skeleton. Safe to run `charts`.", "g")
    return 0

def stage_charts(a):
    _, work = paths(a.carline, a.month)
    inp, _ = paths(a.carline, a.month)
    c = cfg(a.carline)
    if c.get("verify_chart_roles") and not a.i_verified_roles:
        sys.exit(f"{a.carline} chart roles are flagged for verification. Run `roles`, confirm/fix "
                 f"config/carlines.json, then re-run `charts --i-verified-roles`.")
    ch = work / "skeleton" / "unpacked" / "ppt" / "charts"
    r = c["chart_roles"]
    cmds = [
        ["python3", SCRIPTS / "update_volume_chart.py",   ch / f"{r['volume']}.xml",   "--vol", inp / "chart_volume.xlsx"],
        ["python3", SCRIPTS / "update_country_chart.py",  ch / f"{r['country']}.xml",  "--country", inp / "chart_country.xlsx"],
        ["python3", SCRIPTS / "update_platform_chart.py", ch / f"{r['platform']}.xml", "--platform", inp / "chart_platform.xlsx"],
        ["python3", SCRIPTS / "update_sentiment_chart.py", ch / f"{r['sentiment']}.xml",
         "--new-month-sent", inp / "chart_sentiment.xlsx", "--prior-from-last-month", inp / "last_month.pptx"],
    ]
    rc = 0
    for cmd in cmds:
        rc = run(cmd, a.dry_run) or rc
    return rc

def stage_bullets(a):
    _, work = paths(a.carline, a.month)
    yr, mm, _ = month_parts(a.month)
    # prefer the analyst-approved menu (write_bullets format) if present
    menu = work / "menu_confirmed.md"
    if not menu.exists():
        menu = work / "menu.md"
    if not a.dry_run:
        need(menu, "approved bullet menu (menu_confirmed.md or menu.md)")
    # auto-detect the new-month slide (raw-only layouts vary); fall back to --slide
    up = work / "skeleton" / "unpacked"
    slide = up / "ppt" / "slides" / a.slide
    if not a.dry_run and up.exists():
        want = f"{MONTHS[mm]} {yr}".lower()
        import glob as _g
        for s in _g.glob(str(up / "ppt" / "slides" / "slide*.xml")):
            txt = Path(s).read_text(encoding="utf-8", errors="replace").lower()
            if want in txt and "volume" in txt:
                slide = Path(s); break
    cmd = ["python3", SCRIPTS / "write_bullets.py", slide, "--menu", menu]
    return run(cmd, a.dry_run)

def stage_embed(a):
    _, work = paths(a.carline, a.month)
    say("Reminder: patch embed_backing_xlsx.py's chart mapping for this carline first "
        "(chart_roles in config/carlines.json). See docs/slide-configs.md.", "y")
    cmd = ["python3", SCRIPTS / "embed_backing_xlsx.py", work / "skeleton" / "unpacked"]
    return run(cmd, a.dry_run)

def stage_pack(a):
    inp, work = paths(a.carline, a.month)
    yr, mm, _ = month_parts(a.month)
    out = work / f"{a.carline}_Social_Media_{MONTHS[mm]}_{yr}_Sprint_Room.pptx"
    cmd = ["python3", SCRIPTS / "pack_and_qa.py", work / "skeleton" / "unpacked", out]
    if (inp / "last_month.pptx").exists():
        cmd += ["--original", inp / "last_month.pptx"]
    rc = run(cmd, a.dry_run)
    if rc == 0 and not a.dry_run:
        say(f"\nBuilt: {out}", "g")
        say("QA it (render to PDF), then deliver per docs/workflow-phases.md Phase 6.", "g")
    return rc


# ---------- raw-only path (raw.xlsx + last_month.pptx only; no manual chart exports) ----------
def _carline_query(a):
    return getattr(a, "carline_query", None) or f"Ford {a.carline}"

def stage_clean(a):
    """Clean the raw feed (drop false/off-topic, retag passing) -> cleaning_menu.md for sign-off."""
    inp, work = paths(a.carline, a.month); work.mkdir(parents=True, exist_ok=True)
    cmd = ["python3", SCRIPTS / "clean_raw.py", "--carline", a.carline, "--month", a.month,
           "--raw", inp / "raw.xlsx", "--out", work / "cleaning_menu.md"]
    if a.also: cmd += ["--also", *a.also]
    if a.retag_passing: cmd += ["--retag-passing"]
    return run(cmd, a.dry_run)

def stage_derive(a):
    """Derive the 4 chart tables from cleaned raw + last_month deck (writes chart_*.xlsx to inputs)."""
    inp, _ = paths(a.carline, a.month)
    cmd = ["python3", SCRIPTS / "derive_charts.py", "--carline", a.carline, "--month", a.month,
           "--raw", inp / "raw.xlsx", "--last-month", inp / "last_month.pptx",
           "--carline-query", _carline_query(a), "--outdir", inp]
    if a.also: cmd += ["--also", *a.also]
    if a.retag_passing: cmd += ["--retag-passing"]
    return run(cmd, a.dry_run)

def stage_rawmenu(a):
    """Bullet menu from CLEANED raw (raw-only replacement for `menu`). HARD GATE."""
    inp, work = paths(a.carline, a.month); work.mkdir(parents=True, exist_ok=True)
    menu = work / "menu.md"
    cmd = ["python3", SCRIPTS / "build_menu.py", "--carline", a.carline, "--month", a.month,
           "--raw", inp / "raw.xlsx", "--carline-query", _carline_query(a), "--out", menu]
    if a.also: cmd += ["--also", *a.also]
    rc = run(cmd, a.dry_run)
    if rc == 0 and not a.dry_run:
        say("\nHARD GATE: review & approve menu.md, then `skeletonraw`.", "y")
    return rc

def stage_skeletonraw(a):
    """Layout-agnostic skeleton (clones the reported slide + its charts, drops the oldest)."""
    inp, work = paths(a.carline, a.month)
    yr, mm, _ = month_parts(a.month)
    report_date = a.report_date or date.today().strftime("%d.%m.%Y")
    c = cfg(a.carline)
    cmd = ["python3", SCRIPTS / "build_skeleton_raw.py", "--last-month", inp / "last_month.pptx",
           "--carline", a.carline, "--segment", c["segment"], "--new-month", f"{MONTHS[mm]} {yr}",
           "--report-date", report_date, "--out", work / "skeleton"]
    return run(cmd, a.dry_run)

def _newslide_roles(unpacked, new_month_name):
    """Return {role: chartpath} for the skeleton's NEW-month slide (by subtitle month)."""
    import glob as _g
    up = str(unpacked)
    target = None
    for s in _g.glob(f"{up}/ppt/slides/slide*.xml"):
        if new_month_name.lower() in Path(s).read_text(encoding="utf-8", errors="replace").lower():
            rels = Path(f"{up}/ppt/slides/_rels/{Path(s).name}.rels")
            if rels.exists() and "charts/chart" in rels.read_text():
                target = rels; break
    if not target: return {}
    roles = {}
    for cn in re.findall(r'charts/(chart\d+\.xml)', target.read_text()):
        x = Path(f"{up}/ppt/charts/{cn}").read_text(encoding="utf-8", errors="replace")
        types = re.findall(r'<c:(barChart|lineChart|pieChart|scatterChart)>', x)
        sn = [t.lower() for t in re.findall(r'<c:tx>.*?<c:v>([^<]*)</c:v>', x, re.S)]
        cm = re.search(r'<c:cat>(.*?)</c:cat>', x, re.S)
        cats = re.findall(r'<c:v>([^<]*)</c:v>', cm.group(1)) if cm else []
        if "scatterChart" in types or any(t in ("positive","neutral","negative") for t in sn): role="sentiment"
        elif "lineChart" in types: role="volume"
        elif "pieChart" in types:
            j=" ".join(cats).lower()
            role = "platform" if (any(w in j for w in ["news","forum","blog","tumblr","review"]) or " x " in f" {j} ") else "country"
        else: role="unknown"
        roles[role] = f"{up}/ppt/charts/{cn}"
    return roles

def stage_chartsraw(a):
    """Update the new slide's 4 charts from derived tables (role-detected; sentiment reads prior by role)."""
    inp, work = paths(a.carline, a.month)
    yr, mm, prior = month_parts(a.month)
    up = work / "skeleton" / "unpacked"
    roles = _newslide_roles(up, f"{MONTHS[mm]} {yr}")
    need_roles = {"volume","country","platform","sentiment"}
    if not need_roles <= set(roles):
        sys.exit(f"Could not detect all chart roles on the new slide (got {roles}). Run `skeletonraw` first.")
    say(f"New-slide chart roles: { {k: Path(v).name for k,v in roles.items()} }", "b")
    cmds = [
        ["python3", SCRIPTS/"update_volume_chart.py",  roles["volume"],  "--vol", inp/"chart_volume.xlsx"],
        ["python3", SCRIPTS/"update_country_chart.py", roles["country"], "--country", inp/"chart_country.xlsx"],
        ["python3", SCRIPTS/"update_platform_chart.py",roles["platform"],"--platform", inp/"chart_platform.xlsx"],
        ["python3", SCRIPTS/"update_sentiment_chart.py", roles["sentiment"],
         "--new-month-sent", inp/"chart_sentiment.xlsx", "--prior-from-last-month", inp/"last_month.pptx",
         "--prior-month-name", MONTHS[prior]],
    ]
    rc = 0
    for cmd in cmds: rc = run(cmd, a.dry_run) or rc
    return rc


STAGES = {
    "newmonth": stage_newmonth, "doctor": stage_doctor, "validate": stage_validate,
    "menu": stage_menu, "skeleton": stage_skeleton, "fixskeleton": stage_fixskeleton,
    "roles": stage_roles, "charts": stage_charts, "bullets": stage_bullets,
    "embed": stage_embed, "pack": stage_pack,
    # raw-only path:
    "clean": stage_clean, "derive": stage_derive, "rawmenu": stage_rawmenu,
    "skeletonraw": stage_skeletonraw, "chartsraw": stage_chartsraw,
}


def main():
    ap = argparse.ArgumentParser(description="Monthly Sprint Room orchestrator")
    ap.add_argument("stage", choices=list(STAGES))
    ap.add_argument("--carline", required=True)
    ap.add_argument("--month", required=True, help="YYYY-MM, e.g. 2026-07")
    ap.add_argument("--report-date", help="DD.MM.YYYY for the title slide (default: today)")
    ap.add_argument("--slide", default="slide4.xml", help="slide XML for bullets (default slide4.xml)")
    ap.add_argument("--i-verified-roles", action="store_true",
                    help="assert chart roles were confirmed for a verify-flagged carline")
    ap.add_argument("--carline-query", default=None,
                    help="raw-only: the raw Query Name for the carline (default 'Ford <carline>')")
    ap.add_argument("--also", nargs="*", default=[],
                    help="raw-only: extra raw export(s) to merge (adds only queries missing from --raw)")
    ap.add_argument("--retag-passing", action="store_true",
                    help="raw-only clean: promote carline news/blog passing pos/neg to neutral (else REVIEW)")
    ap.add_argument("--dry-run", action="store_true", help="print commands without running them")
    a = ap.parse_args()
    cfg(a.carline)  # validate carline early
    if not SCRIPTS.exists() and a.stage not in ("newmonth", "doctor", "fixskeleton"):
        say(f"Skill scripts not found at {SCRIPTS}. Set SKILL_DIR to the skill location.", "r")
    sys.exit(STAGES[a.stage](a) or 0)


if __name__ == "__main__":
    main()
