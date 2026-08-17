"""
credentials.example.py — Credential template for Automation Domination

Copy this file to credentials.py and fill in your real values.
credentials.py is NOT committed to git (see .gitignore) — keep it local only.

Placeholders only in this file. It IS committed to a PUBLIC repo, so never paste a
real host, username, password, or API key here.
"""

# ── Oracle (iasWorld CAMA database) ───────────────────────────────────────────
# DSN format is host:port/service_name — get the real values from the internal
# setup notes, not from this file.
ORACLE_USER     = "YOUR_ORACLE_USERNAME"
ORACLE_PASSWORD = "YOUR_ORACLE_PASSWORD"
ORACLE_DSN      = "DB_HOST:PORT/SERVICE_NAME"

# ── MLS Matrix / MLS Now (Auth0 SSO) ──────────────────────────────────────────
MLS_USERNAME = "YOUR_MLS_MEMBER_ID"
MLS_PASSWORD = "YOUR_MLS_PASSWORD"

# ── iasWorld web UI (future-year record creation via Playwright) ──────────────
# Two environments: CAST for testing, PROD for real record creation.
IASWORLD_CAST_URL = "https://CAST_HOST/iasworld/main/Login.aspx"
IASWORLD_PROD_URL = "https://PROD_HOST/iasworld/main/Login.aspx"
IASWORLD_USERNAME = "YOUR_IASWORLD_USERNAME"
IASWORLD_PASSWORD = "YOUR_IASWORLD_PASSWORD"

# ── Anthropic (portal grade/condition suggestions) ────────────────────────────
# Only used by local tooling. The portal itself NEVER embeds a key — staff enter
# it at runtime via the portal's gear button, because GitHub Pages is public.
ANTHROPIC_API_KEY = "sk-ant-YOUR-KEY-HERE"
