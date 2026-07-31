"""
embed_backfill.py — compute Voyage embeddings for rows that don't have one yet.

The data is ALREADY in Supabase (Claude loaded all 27 validated rows via MCP).
This fills in the embedding column — the only step that needs your Voyage key.

Usage:
    python embed_backfill.py

Safe to re-run anytime: only touches rows where embedding is null.
Needs SUPABASE_URL, SUPABASE_SERVICE_KEY, VOYAGE_API_KEY in .env (same as before).
"""
import time
from voyageai.error import RateLimitError
from common import sb, embed


def embed_throttled(texts, kind):
    for _ in range(8):
        try:
            return embed(texts, kind=kind)
        except RateLimitError:
            print("    ...rate limited, waiting 21s")
            time.sleep(21)
    raise SystemExit("Rate limit persisted - add a card at dashboard.voyageai.com or rerun later.")

rows = (sb.table("abe_notes")
        .select("id, lead_name, status_before, owner_activity, owner_meaningful_touch")
        .is_("embedding", "null")
        .execute().data)

print(f"{len(rows)} rows need embeddings")

for r in rows:
    situation = (
        f"Status: {r.get('status_before') or ''}\n"
        f"Owner activity (90d): {r.get('owner_activity') or ''}\n"
        f"Owner meaningful touch: {r.get('owner_meaningful_touch') or ''}"
    )
    vec = embed_throttled([situation], kind="document")[0]
    time.sleep(21)  # free tier pacing - delete this line if billing card added
    sb.table("abe_notes").update({"embedding": vec}).eq("id", r["id"]).execute()
    print(f"  ✓ {r['lead_name']}")

n = sb.table("abe_notes").select("id", count="exact").is_("embedding", "null").execute().count
print(f"\nDone. Rows still missing embeddings: {n}")
