#!/usr/bin/env python3
"""Kinmo follow-up agent.

Reads the outreach tracker (an Excel file), finds LinkedIn contacts who are due a
follow-up, and drafts a message for each one. It never sends anything: you review
each draft and send it yourself on LinkedIn, then record it with `sent`.

Usage:
  python3 followup.py                 draft today's follow-ups (writes drafts/YYYY-MM-DD.html)
  python3 followup.py --open          ...and open the drafts page in your browser
  python3 followup.py list            show everyone in the tracker and their follow-up status
  python3 followup.py sent "Name"     record that you sent a follow-up to this contact today
"""
import argparse
import datetime as dt
import html
import json
import os
import re
import subprocess
import sys
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))

DEFAULT_CONFIG = {
    "tracker_path": "sample/network-tracker-sample.xlsx",
    "sheet": "KS",
    "channels": ["LinkedIn"],
    "days_between": 5,
    "max_followups": 2,
    "sender_name": "Your name",
    "sender_line": "Kinmo Solutions · building factories across Africa",
    "closed_words": ["secured", "signed", "declined", "not interested", "closed", "lost", "no longer"],
    "log_path": "followups-log.json",
    "drafts_dir": "drafts",
    "use_claude": False,
    "claude_model": "claude-sonnet-5-5",
}

# Column names in the tracker. Optional columns are used when present.
COL = {
    "company": "Client", "type": "Type", "name": "Contact Name", "role": "Contact Role",
    "purpose": "Purpose", "channel": "Contact Method", "comments": "Comments",
    "contract": "Contract", "production": "Production",
    # optional
    "last": "Last Contacted", "count": "Follow-ups Sent", "replied": "Replied", "url": "LinkedIn URL",
}


# ---------------------------------------------------------------- config and data
def load_config(path=None):
    cfg = dict(DEFAULT_CONFIG)
    path = path or os.path.join(HERE, "config.json")
    if os.path.exists(path):
        with open(path) as f:
            cfg.update(json.load(f))
    return cfg


def resolve(cfg, key):
    p = os.path.expanduser(cfg[key])
    return p if os.path.isabs(p) else os.path.join(HERE, p)


def load_templates():
    with open(os.path.join(HERE, "templates.json")) as f:
        return json.load(f)


def read_tracker(path, sheet):
    try:
        import openpyxl
    except ImportError:
        sys.exit("openpyxl is needed: python3 -m pip install --user openpyxl")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet in wb.sheetnames else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    out = []
    for r in rows[1:]:
        rec = {h: v for h, v in zip(header, r) if h}
        if any(v not in (None, "") for v in rec.values()):
            out.append(rec)
    return out


def load_log(path):
    if os.path.exists(path):
        with open(path) as f:
            return json.load(f)
    return []


# ---------------------------------------------------------------- reading the tracker
DATE_RE = re.compile(r"\b(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\b")


def comment_dates(text, today):
    """Dates written in comments like 'Proposal sent 22/10'. Day/month order (UK style).
    Returns (dates, future) where future lists dates that look like they're ahead of today."""
    dates, future = [], []
    for d, m, y in DATE_RE.findall(text or ""):
        try:
            year = int(y) + (2000 if y and len(y) == 2 else 0) if y else today.year
            when = dt.date(year, int(m), int(d))
        except ValueError:
            continue
        (future if when > today else dates).append(when)
    return dates, future


def as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str) and v.strip():
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
            try:
                return dt.datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


TYPE_LABEL = {"PE Firm": "private equity firm", "Manufacturer": "manufacturer", "Exporter": "exporter",
              "Importer": "importer", "Supplier": "supplier", "Individual": "investor"}


def with_article(word):
    return ("an " if word[:1].lower() in "aeiou" else "a ") + word


def first_name(full):
    return (full or "").strip().split(" ")[0] or "there"


def stage_of(rec):
    c = (rec.get(COL["comments"]) or "").lower()
    if "expensive" in c or "budget" in c or "cost" in c:
        return "price"
    if "proposal" in c or "sow" in c or "statement of work" in c:
        return "proposal"
    purpose = (rec.get(COL["purpose"]) or "").strip().lower()
    return purpose if purpose in ("sales", "partnership", "network") else "sales"


def proposal_date(rec, today):
    m = re.search(r"proposal sent\s+(\d{1,2}/\d{1,2}(?:/\d{2,4})?)", rec.get(COL["comments"]) or "", re.I)
    if not m:
        return ""
    ds, fut = comment_dates(m.group(1), today)
    d = (ds or fut or [None])[0]
    return d.strftime("%-d %B") if d else m.group(1)


def assess(rec, log, cfg, today):
    """Works out where a contact stands. Returns a dict with status and the reason."""
    name = (rec.get(COL["name"]) or "").strip()
    company = (rec.get(COL["company"]) or "").strip()
    comments = rec.get(COL["comments"]) or ""
    a = {"name": name, "company": company, "type": (rec.get(COL["type"]) or "").strip(),
         "role": (rec.get(COL["role"]) or "").strip(), "purpose": (rec.get(COL["purpose"]) or "").strip(),
         "channel": (rec.get(COL["channel"]) or "").strip(), "comments": comments.strip(),
         "url": (rec.get(COL["url"]) or "").strip() if isinstance(rec.get(COL["url"]), str) else "",
         "stage": stage_of(rec)}
    if not name:
        return {**a, "status": "skip", "reason": "No contact name"}
    if a["channel"] not in cfg["channels"]:
        return {**a, "status": "skip", "reason": f"Contact method is {a['channel'] or 'blank'}, not LinkedIn"}
    if rec.get(COL["contract"]) or rec.get(COL["production"]):
        return {**a, "status": "skip", "reason": "Already at contract or production"}
    low = comments.lower()
    if any(w in low for w in cfg["closed_words"]):
        return {**a, "status": "skip", "reason": "Comments say this is closed"}
    if str(rec.get(COL["replied"]) or "").strip().lower() in ("yes", "y", "true", "1"):
        return {**a, "status": "skip", "reason": "Marked as replied"}

    sent = [e for e in log if e["name"].lower() == name.lower()]
    count = int(rec.get(COL["count"]) or 0) + len(sent)
    if count >= cfg["max_followups"]:
        return {**a, "status": "done", "count": count, "reason": f"Already had {count} follow-ups"}

    dates, future = comment_dates(comments, today)
    col_date = as_date(rec.get(COL["last"]))
    if col_date:
        dates.append(col_date)
    dates += [dt.date.fromisoformat(e["date"]) for e in sent]
    if not dates:
        why = f"Comment date {future[0].strftime('%-d/%-m')} looks like it's in the future" if future else "No date to go on"
        return {**a, "status": "nodate", "count": count, "reason": why + ". Add a Last Contacted date."}
    last = max(dates)
    days = (today - last).days
    a.update(count=count, last=last.isoformat(), days=days, number=count + 1)
    if days < cfg["days_between"]:
        return {**a, "status": "wait", "reason": f"Last contact {days} day{'s' if days != 1 else ''} ago; due in {cfg['days_between'] - days} day{'s' if cfg['days_between'] - days != 1 else ''}"}
    return {**a, "status": "due", "reason": f"No reply for {days} days since {last.strftime('%-d %b')} · follow-up {count + 1} of {cfg['max_followups']}"}


# ---------------------------------------------------------------- drafting
def draft(a, cfg, templates, today):
    t = templates
    kind = a["type"] if a["type"] in t["icp"] else "default"
    stage = a["stage"] if a["stage"] in t["messages"] else "sales"
    step = "final" if a["number"] >= cfg["max_followups"] else "first"
    fields = {
        "first": first_name(a["name"]),
        "company": a["company"] or "your team",
        "icp": t["icp"][kind]["pitch"],
        "short": t["icp"][kind]["short"],
        "phase": t["icp"][kind]["phase"],
        "type": TYPE_LABEL.get(a["type"], (a["type"] or "business").lower()),
        "proposal_date": proposal_date({COL["comments"]: a["comments"]}, today) or "recently",
        "sender": cfg["sender_name"],
        "sender_line": cfg["sender_line"],
    }
    fields["a_type"] = with_article(fields["type"])
    body = t["messages"][stage][step].format(**fields)
    return body + "\n\n" + t["signature"].format(**fields)


def polish_with_claude(text, a, cfg):
    """Optional: asks Claude to tailor the draft using the tracker comments. Needs ANTHROPIC_API_KEY."""
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not (cfg.get("use_claude") and key):
        return text
    prompt = (
        "You are helping the founder of Kinmo Solutions, a Nigerian consultancy that helps businesses build "
        "and invest in factories, write a LinkedIn follow-up. Improve the draft below so it reads naturally and "
        "uses the context, but keep it under 90 words, warm and professional, with one clear ask. Do not invent "
        "facts, prices or dates. Keep the signature lines exactly. Reply with the message only.\n\n"
        f"Contact: {a['name']}, {a['role']} at {a['company']} ({a['type']}). Purpose: {a['purpose']}.\n"
        f"Tracker notes: {a['comments'] or 'none'}\nThis is follow-up {a['number']} of {cfg['max_followups']}.\n\n"
        f"Draft:\n{text}"
    )
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps({"model": cfg["claude_model"], "max_tokens": 400,
                         "messages": [{"role": "user", "content": prompt}]}).encode(),
        headers={"x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            out = json.load(r)
        return "".join(b.get("text", "") for b in out.get("content", [])).strip() or text
    except Exception as e:  # keep the template draft if the API is unavailable
        print(f"  (Claude unavailable for {a['name']}: {e}; using the template)")
        return text


# ---------------------------------------------------------------- output
def linkedin_link(a):
    if a.get("url", "").startswith("http"):
        return a["url"], "Open LinkedIn profile"
    q = urllib.parse.quote(f"{a['name']} {a['company']}".strip())
    return f"https://www.linkedin.com/search/results/people/?keywords={q}", "Find on LinkedIn"


def write_page(path, due, others, today, cfg):
    e = html.escape
    cards = []
    for a in due:
        link, label = linkedin_link(a)
        cmd = f'python3 followup.py sent "{a["name"]}"'
        cards.append(f"""
<article class="card">
  <div class="top"><div><h2>{e(a['name'])}</h2><p class="meta">{e(a['role'] + ', ' if a['role'] else '')}{e(a['company'])} · {e(a['type'] or 'Other')} · {e(a['purpose'] or 'Outreach')}</p></div>
  <span class="pill">Follow-up {a['number']} of {cfg['max_followups']}</span></div>
  <p class="why">{e(a['reason'])}{(' · Notes: ' + e(a['comments'].replace(chr(10), ' / '))) if a['comments'] else ''}</p>
  <textarea rows="8" aria-label="Message to {e(a['name'])}">{e(a['message'])}</textarea>
  <div class="actions"><button type="button" class="copy">Copy message</button><a href="{e(link)}" target="_blank" rel="noopener">{label} ↗</a></div>
  <p class="after">After sending, record it: <code>{e(cmd)}</code></p>
</article>""")
    rows = "".join(f"<tr><td>{e(a['name'])}</td><td>{e(a['company'])}</td><td>{e({'nodate': 'Needs a date', 'wait': 'Not due yet', 'done': 'Follow-ups done', 'skip': 'Skipped'}[a['status']])}</td><td>{e(a['reason'])}</td></tr>" for a in others)
    page = f"""<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Kinmo follow-ups · {today.strftime('%-d %b %Y')}</title>
<style>
:root{{--bg:#F4F8F9;--card:#fff;--ink:#0F2A3A;--muted:#4F5D6B;--line:#D5DEE4;--aqua:#0E7C86;--soft:#DDF1F2}}
@media (prefers-color-scheme:dark){{:root{{--bg:#0B1A24;--card:#12242F;--ink:#E6EEF2;--muted:#9FB0BC;--line:#24394A;--aqua:#2BB3BD;--soft:#0F3B40}}}}
*{{box-sizing:border-box}}body{{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif}}
main{{max-width:860px;margin:0 auto;padding:28px 16px 60px;display:flex;flex-direction:column;gap:16px}}
h1{{margin:0;font-size:26px}}.sub{{color:var(--muted);margin:4px 0 0}}
.card{{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:18px;display:flex;flex-direction:column;gap:10px}}
.top{{display:flex;justify-content:space-between;gap:12px;align-items:flex-start;flex-wrap:wrap}}h2{{margin:0;font-size:18px}}
.meta,.why,.after{{margin:0;color:var(--muted);font-size:13px}}.pill{{background:var(--soft);color:var(--aqua);font-weight:600;font-size:12px;border-radius:10px;padding:3px 10px;white-space:nowrap}}
textarea{{width:100%;font:inherit;color:inherit;background:var(--bg);border:1px solid var(--line);border-radius:8px;padding:10px;resize:vertical}}
.actions{{display:flex;gap:14px;align-items:center;flex-wrap:wrap}}button{{font:inherit;font-weight:600;background:var(--aqua);color:#fff;border:0;border-radius:6px;padding:9px 14px;cursor:pointer;min-height:40px}}
a{{color:var(--aqua)}}code{{background:var(--bg);border:1px solid var(--line);border-radius:4px;padding:1px 5px;font-size:12px;word-break:break-all}}
table{{width:100%;border-collapse:collapse;font-size:13px}}th,td{{text-align:left;padding:7px 6px;border-bottom:1px solid var(--line);vertical-align:top}}th{{color:var(--muted);font-weight:600}}
.empty{{color:var(--muted)}}.wrap{{overflow-x:auto}}
</style></head><body><main>
<div><h1>LinkedIn follow-ups · {today.strftime('%-d %B %Y')}</h1><p class="sub">{len(due)} due today. Review each draft, edit it if you like, copy it and send it yourself on LinkedIn.</p></div>
{''.join(cards) or '<p class="empty">Nobody is due a follow-up today.</p>'}
<section class="card"><h2>Everyone else</h2><div class="wrap"><table><thead><tr><th>Contact</th><th>Company</th><th>Status</th><th>Why</th></tr></thead><tbody>{rows}</tbody></table></div></section>
</main>
<script>
document.querySelectorAll(".copy").forEach(function(b){{b.addEventListener("click",function(){{
  var t=b.closest(".card").querySelector("textarea");
  function done(){{b.textContent="Copied";setTimeout(function(){{b.textContent="Copy message"}},1800)}}
  try{{navigator.clipboard.writeText(t.value).then(done,function(){{t.select();document.execCommand("copy");done()}})}}catch(e){{t.select();document.execCommand("copy");done()}}
}})}});
</script></body></html>"""
    with open(path, "w") as f:
        f.write(page)


# ---------------------------------------------------------------- commands
def run(cfg, today, open_page=False, notify=False):
    tracker = resolve(cfg, "tracker_path")
    if not os.path.exists(tracker):
        sys.exit(f"Tracker not found: {tracker}\nSet tracker_path in config.json.")
    log = load_log(resolve(cfg, "log_path"))
    templates = load_templates()
    people = [assess(r, log, cfg, today) for r in read_tracker(tracker, cfg["sheet"])]
    due = [a for a in people if a["status"] == "due"]
    for a in due:
        a["message"] = polish_with_claude(draft(a, cfg, templates, today), a, cfg)
    order = {"nodate": 0, "wait": 1, "done": 2, "skip": 3}
    others = sorted([a for a in people if a["status"] != "due"], key=lambda a: (order[a["status"]], a["name"]))
    out_dir = resolve(cfg, "drafts_dir")
    os.makedirs(out_dir, exist_ok=True)
    page = os.path.join(out_dir, f"{today.isoformat()}.html")
    write_page(page, due, others, today, cfg)
    nodate = sum(1 for a in others if a["status"] == "nodate")
    print(f"{len(due)} follow-up(s) due today · {nodate} contact(s) need a date · drafts: {page}")
    for a in due:
        print(f"  - {a['name']} ({a['company']}): {a['reason']}")
    if notify and sys.platform == "darwin":
        msg = f"{len(due)} LinkedIn follow-up(s) ready to review" if due else "No LinkedIn follow-ups due today"
        subprocess.run(["osascript", "-e", f'display notification "{msg}" with title "Kinmo follow-ups"'], check=False)
    if open_page and due:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.run([opener, page], check=False)
    return due, others


def cmd_sent(cfg, name, today):
    log_path = resolve(cfg, "log_path")
    log = load_log(log_path)
    names = {(r.get(COL["name"]) or "").strip() for r in read_tracker(resolve(cfg, "tracker_path"), cfg["sheet"])}
    match = next((n for n in names if n.lower() == name.strip().lower()), None)
    if not match:
        sys.exit(f'"{name}" is not in the tracker. Check the spelling (it must match Contact Name).')
    log.append({"name": match, "date": today.isoformat()})
    with open(log_path, "w") as f:
        json.dump(log, f, indent=2)
    print(f"Recorded a follow-up to {match} on {today.strftime('%-d %b %Y')}.")


def cmd_list(cfg, today):
    log = load_log(resolve(cfg, "log_path"))
    labels = {"due": "DUE", "nodate": "NEEDS DATE", "wait": "waiting", "done": "done", "skip": "skipped"}
    for r in read_tracker(resolve(cfg, "tracker_path"), cfg["sheet"]):
        a = assess(r, log, cfg, today)
        if a["name"]:
            print(f"{labels[a['status']]:<11} {a['name']:<24} {a['company'][:22]:<22} {a['reason']}")


def main(argv=None):
    p = argparse.ArgumentParser(description="Drafts LinkedIn follow-ups from your outreach tracker.")
    p.add_argument("command", nargs="?", default="draft", choices=["draft", "list", "sent"])
    p.add_argument("name", nargs="?", help="contact name, for the sent command")
    p.add_argument("--config", help="path to config.json")
    p.add_argument("--open", action="store_true", help="open the drafts page when follow-ups are due")
    p.add_argument("--notify", action="store_true", help="show a Mac notification")
    p.add_argument("--today", help="pretend today is this date (YYYY-MM-DD), for testing")
    args = p.parse_args(argv)
    cfg = load_config(args.config)
    today = dt.date.fromisoformat(args.today) if args.today else dt.date.today()
    if args.command == "sent":
        if not args.name:
            sys.exit('Give the contact\'s name, e.g. python3 followup.py sent "Amanda Etuk"')
        cmd_sent(cfg, args.name, today)
    elif args.command == "list":
        cmd_list(cfg, today)
    else:
        run(cfg, today, args.open, args.notify)


if __name__ == "__main__":
    main()
