"""Shared clients + helpers. Requires .env (see .env.example)."""
import os
from dotenv import load_dotenv

load_dotenv()

SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_KEY = os.environ["SUPABASE_SERVICE_KEY"]
VOYAGE_KEY = os.environ["VOYAGE_API_KEY"]
ANTHROPIC_KEY = os.environ["ANTHROPIC_API_KEY"]
CLOSE_KEY = os.environ.get("CLOSE_API_KEY", "")

EMBED_MODEL = "voyage-3.5"   # 1024 dims
CLAUDE_MODEL = "claude-sonnet-4-6"

from supabase import create_client
import voyageai
import anthropic

sb = create_client(SUPABASE_URL, SUPABASE_KEY)
vo = voyageai.Client(api_key=VOYAGE_KEY)
claude = anthropic.Anthropic(api_key=ANTHROPIC_KEY)


def embed(texts, kind="document"):
    """kind: 'document' when saving, 'query' when searching."""
    return vo.embed(texts, model=EMBED_MODEL, input_type=kind).embeddings


def lead_context(row):
    """The 'situation' string we embed — status + activity + touch signal.
    Same recipe for saving AND searching so matches line up."""
    return (
        f"Status: {row.get('Status','')}\n"
        f"Activity: {row.get('Activity','')}\n"
        f"Meaningful touch: {row.get('Meaningful touch?','')}"
    )
