"""
credentials.example.py — Credential template for Automation Domination

Copy this file to credentials.py and fill in your real values.
credentials.py is NOT committed to git (see .gitignore) — keep it local only.

Placeholders only in this file. It IS committed to a PUBLIC repo, so never paste a
real host, username, password, or API key here.
"""

# ── Operator (whoever runs the weekly pipeline on this machine) ───────────────
# Stamped on SALETAB (USER1) and MassEntrance (appraiser). No fallback in the
# code — a missing value stops the run rather than stamping someone else.
APPRAISER_INITIALS  = "ABC"
APPRAISER_FULL_NAME = "First Last"
# Name of YOUR custom export in MLS Matrix (custom exports are per account).
# Its columns must match the list in ONBOARDING.md.
MLS_EXPORT_NAME = "Portal Builder"

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
# Tyler CLOUD (Sept 2026): a single production environment. IASWORLD_URL is the
# cloud app home; unauthenticated it redirects to Tyler Portico, which federates
# to Microsoft Entra ID. So USERNAME is your Microsoft/Portico sign-in EMAIL and
# PASSWORD is your Microsoft password (NOT the old "stark\" domain login).
IASWORLD_URL      = "https://YOUR_TENANT.atentcloud.tylerapp.com/app/home#/dashboard"
IASWORLD_USERNAME = "you@YOUR_COUNTY.gov"
IASWORLD_PASSWORD = "YOUR_MICROSOFT_PASSWORD"

# ── Anthropic (portal grade/condition suggestions) ────────────────────────────
# Only used by local tooling. The portal itself NEVER embeds a key — staff enter
# it at runtime via the portal's gear button, because GitHub Pages is public.
ANTHROPIC_API_KEY = "sk-ant-YOUR-KEY-HERE"
