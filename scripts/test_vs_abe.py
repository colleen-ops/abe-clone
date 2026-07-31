"""
test_vs_abe.py - THE test: can the AI reproduce Abe's judgment?

Leave-one-out over every validated row in Supabase:
  for each row -> hide it, retrieve the most similar OTHER rows,
  have Claude fill Abe's 8-field form, compare to Abe's real answers.

Usage:
    python3 test_vs_abe.py            # all rows
    python3 test_vs_abe.py --limit 10 # quick pass

Prints per-lead diffs + a final agreement scorecard.
Needs the same .env as everything else. Run embed_backfill.py FIRST
(rows without embeddings are skipped).
"""
import sys, json, time
from voyageai.error import RateLimitError
from common import sb, claude, embed, CLAUDE_MODEL


def embed_throttled(texts, kind):
    """Free-tier safe: retry on 3 RPM limit."""
    for _ in range(6):
        try:
            return embed(texts, kind=kind)
        except RateLimitError:
            print("       ...rate limited, waiting 21s")
            time.sleep(21)
    raise SystemExit("Voyage rate limit persisted - add a card or rerun later.")

ALLOWED = {
    "correct_status": ["Yes", "No"],
    "convo_2min": ["Yes", "No"],
    "rep_error": ["", "No connect to owner", "Status-too long",
                  "Sales Rep Conversation Skill",
                  "Didnt collect information and was possible"],
    "task_verdict": ["No - Not Needed", "Yes - Necessary",
                     "Yes - Unnecessary", "No - Should Have"],
    "meaningful_touch": ["Yes", "No"],
    "next_steps": ["Follow up", "Get App", "Get Banks", "Remove from positive",
                   "Adjust follow up date or status",
                   "Sales rep conversation training"],
    "rev_filled": ["Yes", "No"],
}

FIELD_DB = {  # test field -> abe_notes column
    "correct_status": "abe_correct_status",
    "convo_2min": "abe_owner_convo_2min",
    "rep_error": "abe_rep_error_category",
    "task_verdict": "abe_task_verdict",
    "meaningful_touch": "abe_meaningful_touch",
    "next_steps": "abe_next_steps",
    "rev_filled": "abe_rev_filled",
}

PROMPT = """You are Abe, founder of an MCA brokerage, auditing how a rep handled a lead.

JUDGMENT PRINCIPLES:
- Status must match the merchant's stated words/timeline, never rep habit.
- A live ask, buying signal, or unresolved promise NEVER sits in 14/40 Day -> status is No.
- Merchant-set timelines honored correctly = status Yes; don't punish good restraint.

USE THE FACTS LINE FIRST. If a FACTS line is present, these fields are MECHANICAL:
- convo_2min = Yes iff owner_max_call_sec >= 120.
- rev_filled = copy rev_filled from FACTS.
- task_verdict: if open_task=Yes and lead is live -> "Yes - Necessary"; open_task=Yes on a dead lead -> "Yes - Unnecessary"; open_task=No + an explicit merchant commitment (appointment/docs promise/callback window) -> "No - Should Have"; otherwise "No - Not Needed".
(No FACTS line: infer these conservatively from the summary text.)

CALIBRATION - Abe is FORGIVING on rep_error (blank on most leads) and STRICT on meaningful_touch:
- rep_error: BLANK is the default and correct answer for most leads. Assign an error ONLY when it is the dominant story of the lead:
    * "No connect to owner" - owner never reached the decision-maker live (FACTS last_2min_connect=never). Overrides all other errors.
    * "Status-too long" - a LIVE deal/signal parked in a slow bucket for weeks (use days_in_current_status).
    * "Sales Rep Conversation Skill" - a real convo happened and the rep clearly botched it: ignored a stated channel_pref, broke a merchant-agreed timeline, or failed to ask for app/time on a live ask.
    * "Didnt collect information and was possible" - substantive convo happened, key info (revenue/amount/timeline) was clearly gettable, none captured.
  Normal imperfect selling, unanswered chases, merchant ghosting = BLANK. When torn, choose BLANK.
- meaningful_touch: Yes ONLY if the MERCHANT engaged substantively in the window: a 2min+ live convo, or a written exchange with real content (amounts, docs, timing). Reps sending things = No. Merchant one-liners ("ok","text me","who is this?","stop") = No. When torn, choose No.
- correct_status: judge the STATUS, not the rep. A dead/DQ lead correctly parked in 40 Day = Yes even if the lead is worthless. Only No when the CURRENT bucket contradicts the lead's actual state (live thing in slow bucket, dead thing in positive/fast bucket).
- next_steps: pick EXACTLY ONE (two only when both truly required):
    * live deal or honored merchant timeline -> "Follow up"
    * live ask, no app yet -> "Get App"
    * docs/statements promised or needed next -> "Get Banks"
    * dead / DQ / opted-out / wrong bucket needs demotion -> "Remove from positive"
    * lead alive but bucket/date wrong -> "Adjust follow up date or status"
    * rep botched a live convo and needs coaching -> "Sales rep conversation training"
  Never stack "Remove from positive" with anything else.

Here is how you judged the most similar past situations (your validated reviews):
{examples}

New lead to judge:
Status: {status}
Owner activity (90d): {activity}
Owner meaningful touch: {touch}

Return ONLY JSON, no markdown fences:
{{"correct_status":"...","convo_2min":"...","rep_error":"...","task_verdict":"...",
 "meaningful_touch":"...","next_steps":"...","rev_filled":"...","note":"..."}}"""



def norm(v):
    return (str(v).strip() if v is not None else "")


def next_steps_match(ai, abe):
    """Set-overlap: full match = 1, partial (shares a step) = 0.5, else 0."""
    a = {p.strip() for p in norm(ai).split(",") if p.strip()}
    b = {p.strip() for p in norm(abe).split(",") if p.strip()}
    if a == b:
        return 1.0
    return 0.5 if a & b else 0.0


def main():
    limit = None
    if "--limit" in sys.argv:
        limit = int(sys.argv[sys.argv.index("--limit") + 1])

    rows = (sb.table("abe_notes")
            .select("id, lead_name, status_before, owner_activity, owner_meaningful_touch, "
                    + ", ".join(FIELD_DB.values()))
            .not_.is_("embedding", "null")
            .not_.is_("abe_correct_status", "null")
            .execute().data)
    if limit:
        rows = rows[:limit]
    print(f"Testing {len(rows)} validated+embedded rows (leave-one-out)\n")

    scores = {k: 0.0 for k in FIELD_DB}
    low_conf = 0
    n = 0

    for r in rows:
        situation = (f"Status: {r['status_before'] or ''}\n"
                     f"Owner activity (90d): {r['owner_activity'] or ''}\n"
                     f"Owner meaningful touch: {r['owner_meaningful_touch'] or ''}")
        q = embed_throttled([situation], kind="query")[0]
        res = sb.rpc("match_abe_notes", {"query_embedding": q, "match_count": 7}).execute()
        matches = [m for m in res.data if m["id"] != r["id"]][:6]  # leave self out
        top_sim = matches[0]["similarity"] if matches else 0
        if top_sim < 0.75:
            low_conf += 1

        examples = "\n".join(
            f"- {m['status_before']} | {(m['owner_activity'] or '')[:200]}\n"
            f"  YOUR VERDICT -> Correct status: {m['abe_correct_status']} | "
            f"Rep error: {m['abe_rep_error_category'] or 'none'} | "
            f"Task: {m['abe_task_verdict']} | Meaningful: {m['abe_meaningful_touch']} | "
            f"Next: {m['abe_next_steps']}"
            for m in matches)

        msg = claude.messages.create(
            model=CLAUDE_MODEL, max_tokens=400,
            messages=[{"role": "user", "content": PROMPT.format(
                examples=examples, status=r["status_before"],
                activity=r["owner_activity"], touch=r["owner_meaningful_touch"])}])
        txt = msg.content[0].text.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
        try:
            d = json.loads(txt)
        except json.JSONDecodeError:
            print(f"  !! {r['lead_name']}: bad JSON, skipped")
            continue

        time.sleep(21)  # free tier: 3 RPM (delete this line if billing card added)
        n += 1
        diffs = []
        for k, col in FIELD_DB.items():
            ai, abe = norm(d.get(k)), norm(r.get(col))
            if k == "next_steps":
                s = next_steps_match(ai, abe)
            else:
                s = 1.0 if ai == abe else 0.0
            scores[k] += s
            if s < 1.0:
                diffs.append(f"{k}: AI '{ai}' vs ABE '{abe}'")

        mark = "OK " if not diffs else "DIFF"
        print(f"[{mark}] {r['lead_name']} (top sim {top_sim:.2f})")
        for dd in diffs:
            print(f"       {dd}")

    print("\n" + "=" * 52)
    print(f"SCORECARD ({n} leads)")
    for k in FIELD_DB:
        pct = 100 * scores[k] / n if n else 0
        print(f"  {k:<18} {pct:5.1f}%")
    overall = 100 * sum(scores.values()) / (len(scores) * n) if n else 0
    print(f"  {'OVERALL':<18} {overall:5.1f}%")
    print(f"  low-confidence retrievals (top sim <0.75): {low_conf}/{n}")
    print("\nRead: correct_status is the headline. 75%+ = ready for note-only mode.")
    print("Check every DIFF - some will be the AI being wrong, some will be rows")
    print("where reasonable judges split; only the first kind matters.")


if __name__ == "__main__":
    main()
