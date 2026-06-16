# Stark County Auditor — MLS/CAMA Review Portal
## Claude Code Project Memory

---

## What This Project Is
A staff review portal for the Stark County Auditor's Office, Assessment Division.
Staff use it to compare MLS listing data against CAMA (iasWorld) property records
and review Grade/Condition mismatches weekly.

**Live URL:** https://starkcountyohio.github.io/cama-review-portal/
**Portal URL:** https://starkcountyohio.github.io/cama-review-portal/review_portal.html
**GitHub repo:** https://github.com/Starkcountyohio/cama-review-portal (public)

---

## File Structure
```
Portal Builder/
├── CLAUDE.md                        ← this file (Claude Code memory)
├── build_portal.py                  ← Monday deploy script (run this weekly)
├── review_portal_template.html      ← React app template (never upload this directly)
├── review_portal.html               ← OUTPUT: built by build_portal.py, upload to GitHub
├── index.html                       ← Landing page (upload to GitHub once, rarely changes)
├── data/
│   ├── Grade_Con_Compara.xlsx       ← Appraiser before/after decisions (monthly update)
│   └── weekly_json/                 ← Archive of weekly JSON files (2026-W12.json etc.)
└── prompt_training/
    └── update_prompt.py             ← Monthly retraining script
```

---

## Monday Weekly Workflow
1. Run `python build_portal.py` — GUI picks Excel files + photo folder
2. Script outputs `review_portal.html` (~20-40MB with photos)
3. Upload `review_portal.html` to GitHub repo (replaces previous week)
4. Staff refresh the URL — new week loads with green "Data Pre-Loaded" banner

## Monthly Prompt Retraining Workflow
1. Export latest appraiser decisions from iasWorld → save as `data/Grade_Con_Compara.xlsx`
2. Save latest weekly JSON files to `data/weekly_json/`
3. Run `python prompt_training/update_prompt.py`
4. Script analyzes decisions, updates the AI prompt in `review_portal_template.html`
5. Run `python build_portal.py` to rebuild with new prompt
6. Upload to GitHub

---

## Tech Stack
- **Frontend:** React 18 + Babel + SheetJS — **SELF-HOSTED in `lib/`** (NOT from CDN: county network blocks unpkg.com & cdn.sheetjs.com → blank dark-navy screen). In-browser Babel transpiles JSX. Google Fonts still external (cosmetic only). `lib/` must stay committed.
- **Build:** Python 3, openpyxl, Pillow
- **Backend:** None — fully static site
- **Hosting:** GitHub Pages (public repo required for free Pages)
- **AI:** Anthropic Claude API (claude-sonnet-4-6), called direct from browser — model ID must be a current, non-retired model (claude-sonnet-4-20250514 was retired 2026-06-15)
- **API Key:** Set via ⚙️ button in portal bottom-right, stored in memory only — NEVER embed in HTML

---

## Key Domain Concepts
- **CAMA:** Computer-Assisted Mass Appraisal — iasWorld is the CAMA system
- **CDU:** Condition, Desirability, Utility — EX/VG/GD/AV/FR/PR (best to worst)
- **GRADE:** 2024 letter scale — X+/X/X- → A+/A/A- → B+/B/B- → C+/C/C- → D → E
- **EFFYR:** Effective Year Built — used for depreciation calculation
- **YRBLT:** Actual year built
- **SFLA:** Above-grade finished living area (sq ft)
- **FIXBATH/FIXHALF:** Full baths / half baths
- **SALEKEY:** iasWorld sale record identifier
- **NOPAR:** Number of parcels in the sale
- **WindowId:** iasWorld URL parameter needed to open parcel links
- **iasWorld URL pattern:** https://iasworld.starkcountyohio.gov/iasworld/Maintain/Transact.aspx?txtMaskedPin={parcelId}&windowId={windowId}

---

## AI Grade & Condition Prompt — Training History

### What the prompt is trained on:
The AI prompt inside `review_portal_template.html` is calibrated on ACTUAL
Stark County appraiser decisions — not just general appraisal theory.

### Training data used so far:
- **355 parcels** across weeks W12–W16 (2026)
- **Grade_Con_Compara.xlsx** — actual appraiser before/after decisions

### Key findings baked into the prompt:
- CDU changed on 35% of parcels, Grade on 12%
- Most common CDU changes: AV→GD (46x), GD→AV (31x), VG→GD (14x), AV→FR (13x)
- "opportunity" → strongest single downgrade word (43% of CDU downgrades, 1.5x baseline)
- "investor" → 17% of downgrades (2.4x baseline)
- "as-is" → 8% of downgrades (4x baseline)
- Hard negatives (estate sale/cash only/vacant) → 88% change rate
- "remodeled" → 28% of upgrades (2x baseline) — strongest upgrade word
- "renovated" → 23% of upgrades (3x baseline)
- "new furnace/windows" → 2x baseline upgrade rate
- "hardwood" alone → below baseline, confirmed noise
- Location/school fluff → zero predictive value, ignore completely
- Photo always outweighs remarks when they conflict

### Evidence weighting hierarchy (in prompt):
1. Photo (highest)
2. Hard negative language
3. Specific named system updates
4. Soft negative language
5. Positive generic language (lowest — realtor marketing)
6. Cosmetic only terms (ignore for CDU)

---

## Known Issues / Decisions Made
- GitHub Pages requires PUBLIC repo for free tier — key must NEVER be embedded in HTML
- API key is set manually via ⚙️ button each Monday session
- Chrome requires explicit window dimensions in window.open() to open new window vs tab
- iasWorld and Zillow open in named windows: "iasWorldWindow" and "zillowWindow"
- SharePoint cannot render HTML files directly — GitHub Pages is the solution
- localStorage cannot share state across remote users — no shared persistence currently
- Photos embed as base64 in HTML — 60-100 photos ≈ 20-40MB output file

---

## Staff
- 7 remote staff reviewers
- Office: Stark County Auditor's Office, Assessment Division, Canton OH
- Admin/builder: Jason Jeffries (jmjeffri)
- CAMA system: iasWorld (Stark County Ohio)
