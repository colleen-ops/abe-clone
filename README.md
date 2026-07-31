# Hunter Mode AI — Getty Advance sales-audit pipeline

RAG system that audits rep lead-handling in Abe's voice and posts [HUNTER MODE]
draft notes to Close. See CLAUDE.md for the operating playbook and
criteria/abe_rules.md for the judging law.

## Setup
1. `cp .env.example .env` and fill keys (never commit .env)
2. `pip install -r requirements.txt`
3. Daily: see CLAUDE.md "DAILY ROUTINE" — or let Claude Code run it:
   `claude -p "Run the daily Hunter Mode routine per CLAUDE.md"`

## Scripts
- scripts/build_audit_v5.py — pull leads from Close → summaries + FACTS (+ --user/--asof test flags)
- scripts/draft_v3.py — judge (RAG vs Supabase cabinet) → 8 verdicts + Abe-voice note
- scripts/post.py — post [HUNTER MODE] notes to Close (--send)
- scripts/ingest_v2.py / embed_backfill.py — grow the cabinet from Abe-validated sheets
- scripts/test_e2e.py / test_vs_abe.py — accuracy tests vs Abe's stored verdicts
