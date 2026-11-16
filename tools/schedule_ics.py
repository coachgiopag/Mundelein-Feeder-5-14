"""Build calendar feeds (docs/cal/*.ics) from docs/assets/schedule.json.

One feed per team (b4r.ics = 4th grade boys Red, g6b.ics = 6th grade girls Black, b8.ics = 8th grade boys,
...) plus all.ics. 8th grade has one boys team and one girls team, no Red/Black.
Families subscribe from mundyball.com/schedule, so their calendars pick up changes on their own.
Run by seo.py; can also be run directly: python3 tools/schedule_ics.py
"""
import datetime, hashlib, json, os, re
from zoneinfo import ZoneInfo

HERE = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(HERE, "..", "docs")
CT = ZoneInfo("America/Chicago")
GRADES = ["4", "5", "6", "7", "8"]
ORD = {"4": "4th", "5": "5th", "6": "6th", "7": "7th", "8": "8th"}


def expand(codes):
    """'B6' -> {B6R, B6B}; '6' -> all four 6th grade teams; '*' -> every team."""
    out = set()
    for c in (codes or "*").split():
        gs = [c[0]] if c[:1] in ("B", "G") else ["B", "G"]
        c = c[1:] if c[:1] in ("B", "G") else c
        gr = [c[0]] if c[:1] in GRADES else GRADES
        c = c[1:] if c[:1] in GRADES else c
        cs = [c] if c in ("R", "B") else ["R", "B"]
        out |= {g + n + ('' if n == '8' else k) for g in gs for n in gr for k in cs}  # 8th grade: one team, no Red/Black
    return out


def parse_range(day, text):
    a, b = re.split(r"\s*[–-]\s*", text.replace("*", "").strip())
    mb = re.search(r"(AM|PM)", b, re.I)
    if not re.search(r"(AM|PM)", a, re.I):
        a += " " + mb.group(1)

    def at(t):
        h, m, ap = re.match(r"(\d{1,2}):(\d{2})\s*(AM|PM)", t.strip(), re.I).groups()
        h = int(h) % 12 + (12 if ap.upper() == "PM" else 0)
        return datetime.datetime.combine(day, datetime.time(h, int(m)), CT)

    return at(a), at(b)


def esc(s):
    return s.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def fold(line):
    b = line.encode("utf-8")
    if len(b) <= 74:
        return line
    parts, cur = [], b""
    for ch in line:
        e = ch.encode("utf-8")
        if len(cur) + len(e) > 74:
            parts.append(cur.decode("utf-8"))
            cur = b" "
        cur += e
    parts.append(cur.decode("utf-8"))
    return "\r\n".join(parts)


def utc(d):
    return d.astimezone(datetime.timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def team_name(code):
    name = "%s Grade %s" % (ORD[code[1]], "Boys" if code[0] == "B" else "Girls")
    return name if len(code) == 2 else name + (" Red" if code[2] == "R" else " Black")


def build():
    data = json.load(open(os.path.join(DOCS, "assets", "schedule.json"), encoding="utf-8"))
    loc = data["location"]
    events = []  # (teams, vevent lines)
    for s in data["sessions"]:
        day = datetime.date.fromisoformat(s["date"])
        for slot in s.get("slots") or []:
            time, who, codes = slot[0], slot[1], slot[2]
            kind = slot[3] if len(slot) > 3 else "Practice"
            start, end = parse_range(day, time)
            uid = hashlib.sha1(("%s|%s|%s|%s" % (s["date"], time, who, s["gym"])).encode()).hexdigest()[:20] + "@mundyball.com"
            lines = [
                "BEGIN:VEVENT",
                "UID:" + uid,
                "DTSTAMP:" + utc(start),
                "DTSTART:" + utc(start),
                "DTEND:" + utc(end),
                "SUMMARY:" + esc("Jr. Mustangs %s: %s" % (kind, who)),
                "LOCATION:" + esc("%s, %s" % (s["gym"], loc)),
                "DESCRIPTION:" + esc("%s · %s\nTimes and gyms may change: https://www.mundyball.com/schedule" % (s["gym"], who)),
                "URL:https://www.mundyball.com/schedule",
                "END:VEVENT",
            ]
            events.append((expand(codes), lines))

    outdir = os.path.join(DOCS, "cal")
    os.makedirs(outdir, exist_ok=True)
    for old in os.listdir(outdir):  # drop feeds for teams that no longer exist
        if old.endswith(".ics"):
            os.remove(os.path.join(outdir, old))
    teams = sorted({g + n + ("" if n == "8" else k) for g in "BG" for n in GRADES for k in "RB"})
    feeds = {t.lower(): ("Jr. Mustangs · " + team_name(t), {t}) for t in teams}
    feeds["all"] = ("Jr. Mustangs · All Teams", set(teams))
    for slug, (name, want) in feeds.items():
        out = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//MundyBall//Schedule//EN", "CALSCALE:GREGORIAN",
               "METHOD:PUBLISH", "X-WR-CALNAME:" + esc(name), "X-WR-TIMEZONE:America/Chicago",
               "REFRESH-INTERVAL;VALUE=DURATION:PT6H", "X-PUBLISHED-TTL:PT6H"]
        for tset, lines in events:
            if tset & want:
                out += lines
        out.append("END:VCALENDAR")
        with open(os.path.join(outdir, slug + ".ics"), "w", encoding="utf-8", newline="") as f:
            f.write("\r\n".join(fold(l) for l in out) + "\r\n")
    return len(feeds), len(events)


if __name__ == "__main__":
    n, e = build()
    print("calendar feeds: %d files, %d events" % (n, e))
