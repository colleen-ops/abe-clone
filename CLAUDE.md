# Hunter Mode AI — Claude Code Playbook

You are operating Getty Advance's sales-audit pipeline ("Hunter Mode AI").
It judges how reps handled leads, in Abe's (the founder's) voice, and posts
draft review notes to Close CRM.

## HARD GUARDRAILS
- NOTE-ONLY MODE: never change lead statuses, tasks, or contacts in Close.
  You may only post notes prefixed "[HUNTER MODE]".
- Never SMS/email merchants. Never touch leads with a STOP/opt-out on record.
- Judgment criteria live in criteria/abe_rules.md — they are LAW. If a case
  doesn't fit them, mark the note "LOW CONFIDENCE - needs Abe" instead of guessing.
- Secrets come from .env (never commit it, never print values).

## DAILY ROUTINE (in order)
1. Build: `python3 scripts/build_audit_v5.py data/today_ids.txt`
   (today_ids.txt = lead IDs with fresh activity; produced upstream or via n8n)
2. Draft: `python3 scripts/draft_v3.py data/today_ids_built.xlsx`
3. Review the drafts yourself: flag any verdict that contradicts criteria/abe_rules.md.
4. Post: `python3 scripts/post.py data/today_ids_drafts.xlsx --send`
5. Summarize the batch (counts, low-confidence leads, compliance flags) — one
   short message; if a Slack webhook is configured in .env, post it there.

## WHEN THINGS BREAK
- Voyage RateLimitError → wait and retry; scripts already throttle.
- Close 5xx/timeout → retry the single lead, don't kill the batch.
- Anthropic credit error → STOP and report; do not degrade to guessing.
- A lead errors twice → skip it, list it in the batch summary.

## CORRECTION LOOP
When Abe edits a [HUNTER MODE] note or the review sheet, those rows are
re-ingested via scripts/ingest_v2.py and embedded via scripts/embed_backfill.py.
Never modify abe_* verdict columns in Supabase yourself — they are Abe's ground truth.

## STACK
Close CRM (leads, activity, transcripts via _fields=transcript) ·
Supabase pgvector `abe_notes` + `match_abe_notes` (project tnmdtrfkrmldvkzulyxm) ·
Voyage voyage-3.5 embeddings (situation text only) · Anthropic API (judge).
