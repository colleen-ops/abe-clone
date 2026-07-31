"""
build_audit_v5.py - audit sheet builder, Abe's official criteria edition.

Per lead:
  1. Find the LAST USER WHO COMMUNICATED (latest outbound call/SMS/email) - judged user
  2. Pull that user's activity (90d) + all merchant inbounds + their all-time max call
  3. Compute FACTS (mechanical fields) + Claude writes the summary
  4. Output xlsx ready for draft_v3 / test_e2e

Usage:
    python3 build_audit_v5.py all46_ids.txt
    python3 build_audit_v5.py some_sheet.xlsx      # reuses its Lead ID column
    -> writes <input>_built.xlsx

Needs CLOSE_API_KEY + ANTHROPIC_API_KEY in .env.
"""
import os, sys, re, requests, openpyxl
from datetime import datetime, timedelta, timezone
from dotenv import load_dotenv
import anthropic

load_dotenv()
KEY = os.environ["CLOSE_API_KEY"]
claude = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
REV_FIELD = "custom.cf_Rlkx5bEUjkavr5q7iS7j7LxArX6K08ucklALia9XUw3"  # QB Rep Reported Monthly Revenue
CUTOFF = (datetime.now(timezone.utc) - timedelta(days=90)).isoformat()

_users = {}


def close_get(path, **params):
    import time as _t
    for attempt in range(4):
        try:
            r = requests.get(f"https://api.close.com/api/v1{path}", params=params, auth=(KEY, ""), timeout=45)
            r.raise_for_status()
            return r.json()
        except Exception:
            if attempt == 3:
                raise
            _t.sleep(3 * (attempt + 1))


def user_name(uid):
    if uid and uid not in _users:
        try:
            _users[uid] = close_get(f"/user/{uid}/").get("display_name", uid)
        except Exception:
            _users[uid] = uid
    return _users.get(uid, "Unknown")


def comm_activity(lead_id, forced_uid=None, asof_dt=None):
    """Find (or pin) the judged user; build their lines + inbounds; FACTS. asof_dt drops later activity."""
    calls = close_get("/activity/call/", lead_id=lead_id, date_created__gt=CUTOFF, _limit=40,
                      _fields="id,user_id,direction,date_created,duration,note,disposition").get("data", [])
    smss = close_get("/activity/sms/", lead_id=lead_id, date_created__gt=CUTOFF, _limit=40).get("data", [])
    emails = close_get("/activity/email/", lead_id=lead_id, date_created__gt=CUTOFF, _limit=15,
                       _fields="id,user_id,direction,date_created,subject,body_text").get("data", [])
    if asof_dt:
        calls = [c for c in calls if c["date_created"][:19] <= asof_dt]
        smss = [x for x in smss if x["date_created"][:19] <= asof_dt]
        emails = [e for e in emails if e["date_created"][:19] <= asof_dt]

    if forced_uid:
        judged_uid = forced_uid
    else:
        outbound = ([c for c in calls if c.get("direction") != "inbound"] +
                    [s for s in smss if s.get("direction") == "outbound"] +
                    [e for e in emails if e.get("direction") == "outgoing"])
        outbound.sort(key=lambda a: a["date_created"], reverse=True)
        judged_uid = outbound[0].get("user_id") if outbound else None
    judged_user = user_name(judged_uid) if judged_uid else "NONE"

    facts = {"judged_user": judged_user, "max_call_sec_90d": 0, "answered_2min_90d": 0,
             "last_2min_connect": "", "dials_since_connect": 0,
             "inbound_msgs_90d": 0, "last_inbound": "", "channel_pref": "",
             "max_call_ever_sec": 0}
    def call_ai_summary(call_id):
        """Call transcript (hidden REST field - must be requested via _fields) or rep note."""
        try:
            full = close_get(f"/activity/call/{call_id}/", _fields="id,note,transcript")
        except Exception:
            return ""
        txt = full.get("transcript") or full.get("note") or ""
        return " ".join(str(txt).split())[:1800]

    lines = []
    for c in calls:
        if c.get("user_id") == judged_uid:
            dur = c.get("duration", 0) or 0
            facts["max_call_sec_90d"] = max(facts["max_call_sec_90d"], dur)
            if dur >= 120:
                facts["answered_2min_90d"] += 1
                facts["last_2min_connect"] = max(facts["last_2min_connect"], c["date_created"][:10])
            note = (c.get("note") or "")[:150]
            summ = call_ai_summary(c["id"]) if dur >= 60 else ""
            extra = f" - {note}" if note else ""
            if summ:
                extra += f" | CALL TRANSCRIPT: {summ}"
            lines.append(f"CALL {c['date_created'][:10]} {c.get('direction','')} {dur}s{extra}")
    for s in smss:
        if s.get("direction") == "inbound":
            facts["inbound_msgs_90d"] += 1
            facts["last_inbound"] = max(facts["last_inbound"], s["date_created"][:10])
            low = (s.get("text") or "").lower()
            if "text me" in low or "please text" in low:
                facts["channel_pref"] = "TEXT (merchant stated)"
            elif "email me" in low or "please email" in low:
                facts["channel_pref"] = "EMAIL (merchant stated)"
            lines.append(f"INBOUND SMS {s['date_created'][:10]}: {(s.get('text') or '')[:150]}")
        elif s.get("user_id") == judged_uid:
            lines.append(f"SMS {s['date_created'][:10]} out: {(s.get('text') or '')[:150]}")
    for e in emails:
        if e.get("user_id") == judged_uid or e.get("direction") == "incoming":
            body = " ".join((e.get("body_text") or "").split())[:300]
            tag = "INBOUND EMAIL" if e.get("direction") == "incoming" else "EMAIL out"
            lines.append(f"{tag} {e['date_created'][:10]}: {(e.get('subject') or '')[:80]}"
                         + (f" | {body}" if body else ""))

    # all-time max call for judged user (criterion 1 has no window)
    for c in close_get("/activity/call/", lead_id=lead_id, _limit=100).get("data", []):
        if c.get("user_id") == judged_uid and (not asof_dt or c["date_created"][:19] <= asof_dt):
            facts["max_call_ever_sec"] = max(facts["max_call_ever_sec"], c.get("duration", 0) or 0)

    lines = sorted(set(lines), reverse=True)[:40]
    for ln in lines:
        if ln.startswith("CALL") and facts["last_2min_connect"] and ln[5:15] > facts["last_2min_connect"] and " 0s" in ln:
            facts["dials_since_connect"] += 1
    return lines, facts, judged_uid, judged_user, call_ai_summary


def status_at(lead_id, current_status, asof_dt=None):
    """(status_label, days_in_status) — reconstructed at asof when given."""
    try:
        chs = close_get("/activity/status_change/lead/", lead_id=lead_id, _limit=25).get("data", [])
    except Exception:
        chs = []
    ref = datetime.fromisoformat(asof_dt).replace(tzinfo=timezone.utc) if asof_dt else datetime.now(timezone.utc)
    label, changed = current_status, None
    for ch in chs:  # newest first
        d = ch["date_created"][:19]
        if asof_dt and d > asof_dt:
            label = ch.get("old_status_label") or label
            continue
        changed = datetime.fromisoformat(ch["date_created"].replace("Z", "+00:00"))
        if asof_dt:
            label = ch.get("new_status_label") or label
        break
    days = (ref - changed).days if changed else -1
    return label, days


def best_call_ever(lead_id, summary_fn=None):
    calls = close_get("/activity/call/", lead_id=lead_id, _limit=100).get("data", [])
    if not calls:
        return "No calls on record"
    best = max(calls, key=lambda c: c.get("duration", 0) or 0)
    who = user_name(best.get("user_id"))
    out = (f"Best call ever: {round((best.get('duration') or 0)/60,1)}m on {best['date_created'][:10]} by {who}"
           + (f" - {(best.get('note') or '')[:150]}" if best.get("note") else ""))
    if summary_fn and (best.get("duration") or 0) >= 60:
        summ = summary_fn(best["id"])
        if summ:
            out += f" | CALL TRANSCRIPT: {summ}"
    return out


def facts_line(facts, rev_val, task_by_user, open_task, status_days):
    rev_filled = "Yes" if (rev_val and not str(rev_val).strip().startswith("0.")) else "No"
    return (f"FACTS: judged_user={facts['judged_user']} | "
            f"max_call_ever_sec={facts['max_call_ever_sec']} | "
            f"max_call_sec_90d={facts['max_call_sec_90d']} | "
            f"answered_2min_90d={facts['answered_2min_90d']} | "
            f"last_2min_connect={facts['last_2min_connect'] or 'never'} | "
            f"dials_since_connect={facts['dials_since_connect']} | "
            f"inbound_msgs_90d={facts['inbound_msgs_90d']} | "
            f"last_inbound={facts['last_inbound'] or 'none'} | "
            f"days_in_current_status={status_days} | "
            f"channel_pref={facts['channel_pref'] or 'none stated'} | "
            f"rev_filled={rev_filled} | task_by_last_user={'Yes' if task_by_user else 'No'} | "
            f"open_task={'Yes' if open_task else 'No'}")


JUDGE_PROMPT = """You audit MCA sales activity. A "meaningful touch" = merchant showed interest AND the rep
captured the business picture (what the company does, revenue) and emailed the merchant.
Dials, voicemails, one-word replies, and STOP opt-outs are NOT meaningful.

Lead: {name} | Status: {status}
Judged user (last to communicate): {owner}
Their activity (last 90 days):
{activity}
Historical context: {best}

Reply in EXACTLY this format, nothing else:
SUMMARY: <2-3 sentences: what the judged user actually did and what the merchant said>
MEANINGFUL TOUCH: <YES / NO / PARTIAL> - <one-line reason with date/quote>"""


def judge(name, status, owner, activity, best):
    msg = claude.messages.create(
        model="claude-sonnet-4-6", max_tokens=400,
        messages=[{"role": "user", "content": JUDGE_PROMPT.format(
            name=name, status=status, owner=owner,
            activity="\n".join(activity) or "(no activity in 90 days)", best=best)}])
    txt = msg.content[0].text.strip()
    summary = touch = ""
    for line in txt.splitlines():
        if line.startswith("SUMMARY:"):
            summary = line.split(":", 1)[1].strip()
        elif line.startswith("MEANINGFUL TOUCH:"):
            touch = line.split(":", 1)[1].strip()
    return summary, touch


def load_lead_ids(path):
    if path.endswith(".xlsx"):
        ids = []
        for ws in openpyxl.load_workbook(path).worksheets:
            headers = [c.value for c in ws[1]]
            if "Lead ID" not in headers:
                continue
            col = headers.index("Lead ID") + 1
            for row in range(2, ws.max_row + 1):
                v = ws.cell(row=row, column=col).value
                if v and str(v).startswith("lead_"):
                    ids.append(str(v).strip())
        return list(dict.fromkeys(ids))
    return [l.strip() for l in open(path) if l.strip().startswith("lead_")]


def main(path, forced_uid=None, asof=None):
    ids = load_lead_ids(path)
    asof_dt = (asof + "T23:59:59") if asof else None
    global CUTOFF
    if asof:
        from datetime import datetime as _dt
        CUTOFF = (_dt.fromisoformat(asof).replace(tzinfo=timezone.utc) - timedelta(days=90)).isoformat()
    print(f"{len(ids)} leads to build" + (f" | judged user pinned: {forced_uid}" if forced_uid else "")
          + (f" | as of {asof}" if asof else "") + "\n")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Audit"
    ws.append(["Lead", "Lead ID", "Status", "Judged User (last communicator)",
               "Owner Activity Summary (90d)", "Meaningful touch?", "Historical Best Call"])

    for lid in ids:
        try:
            lead = close_get(f"/lead/{lid}/", _fields=f"id,display_name,status_label,{REV_FIELD}")
            name = lead.get("display_name", lid)
            status = lead.get("status_label", "")

            acts, facts, judged_uid, judged_user, summary_fn = comm_activity(lid, forced_uid, asof_dt)
            all_tasks = close_get("/task/", lead_id=lid, _limit=25).get("data", [])
            if asof_dt:
                all_tasks = [t for t in all_tasks if t.get("date_created", "")[:19] <= asof_dt]
            open_tasks = [t for t in all_tasks if not t.get("is_complete")]
            task_by_user = any(t.get("assigned_to") == judged_uid for t in open_tasks)  # OPEN tasks only
            open_task = bool(open_tasks)
            status, sdays = status_at(lid, status, asof_dt)
            fline = facts_line(facts, lead.get(REV_FIELD), task_by_user, open_task, sdays)
            best = best_call_ever(lid, summary_fn)
            summary, touch = judge(name, status, judged_user, acts, best)
            summary = fline + "\n" + summary

            ws.append([name, lid, status, judged_user, summary, touch, best])
            print(f"  \u2713 {name} | {judged_user} | {touch[:60]}")
        except Exception as e:
            ws.append(["ERROR", lid, "", "", str(e)[:200], "", ""])
            print(f"  \u2717 {lid}: {e}")

    out = (path.rsplit(".", 1)[0]) + "_built.xlsx"
    wb.save(out)
    print(f"\nSaved: {out}")


USERS = {"makar": "user_QGYi4cGXAsQyGxQukFfqfmy3YkVMkeQkNZpss0YPDcd",
         "calder": "user_umyJYCUlaz8CAerNp9xojOtnJVFKveonTYeWqfhU8Vy"}

if __name__ == "__main__":
    forced_uid = asof = None
    args = sys.argv[1:]
    if "--user" in args:
        v = args[args.index("--user") + 1]
        forced_uid = USERS.get(v.lower(), v)
    if "--asof" in args:
        asof = args[args.index("--asof") + 1]  # YYYY-MM-DD
    main(args[0], forced_uid, asof)
