# Stark County Auditor — Monday Pipeline Onboarding

How to run the weekly MLS/CAMA Review Portal pipeline, grounded in the run of
**Monday August 10, 2026 (week 2026-W33)**. Follow top to bottom. Commands are run
from the project root: `Portal Builder/`.

> Paths below are written relative to the project root (`<portal-root>`) because
> that's where the pipeline scripts run. A fresh PowerShell window does **not**
> open there — it opens in your user profile — so a bare `.\run_zillow_photos.bat`
> fails with `CommandNotFoundException`. Use the full quoted path with the `&`
> call operator (quotes for the spaces in the path, `&` to actually execute it
> rather than just echo the string back):
> `& "<portal-root>\run_zillow_photos.bat"`

- **Live portal:** https://starkcountyohio.github.io/cama-review-portal/review_portal.html
- **Repo:** https://github.com/Starkcountyohio/cama-review-portal (public — never commit keys)
- **Cadence:** Jason runs this Monday. Appraisers review Tue–Fri. Field-visit flags assigned the following Tuesday.

---

## Prerequisites (once per machine)
- `Automation Domination/credentials.py` present. **Not** committed and never will
  be — copy `credentials.example.py` to `credentials.py` and fill in the real
  values from the internal setup notes. Every other pipeline script *is* tracked,
  so a fresh clone plus this one file is a working build machine.
- Network access to the CAMA database and to MLS Now from the build machine.
  If step 1 fails, check credentials and the database service before assuming a
  connectivity problem — that is rarely the cause on a configured build machine.
- Python 3 with openpyxl, Pillow, pandas, oracledb, playwright.

---

## The pipeline, in order

There are **two ways** to run it — both do the same 8 steps in the same order:

- **All-in-one (default):** one command, three human gates. Use when Jason runs
  the whole thing himself at the keyboard. → **§1–§6 below.**
- **Split (assisted):** the automated halves are separated from the interactive
  middle, so the data pull and the build can be driven by an assistant (e.g.
  Claude Code) while Jason handles the Zillow CAPTCHAs and the photo review. Same
  outputs, same order — just cut at the human gates. → **see "Split run" next.**

### Split run (data-pull / Zillow / build) — assisted mode
Used the week of 2026-W31. Two helper scripts do the non-interactive halves;
they call the exact same functions as `run_weekly.py`, and both default to
**today's date**, so run them on the Monday.

**A. Data pull — steps 1–3 (CAMA + MLS + compare), stops before Zillow:**
```powershell
python "Automation Domination\run_datapull.py"
```
Writes the 4 Excel files to `MLSvsCAMA\<week>\` and stops. Only gate is the MLS
Auth0 login. Sanity-check the comparison summary (see §1). Note: mismatch **rows
≠ parcels** — a parcel gets one row per mismatched field, and a relisted property
can appear in *both* the mismatch and perfect files. The portal count is unique
parcels, deduped across both files (mismatch wins, so the discrepancy still shows).
W33: 109 CAMA → 81 mismatch rows / 47 parcels + 21 perfects, 0 in both =
**68 unique parcels**. (W32: 173 CAMA → 134 unique parcels. W31: 174 → 122.)

Two summary numbers look alarming and are not:
- **"Missing in CAMA: 2,759"** — the MLS saved search returns the whole
  active/sold set (2,827 listings in W33), not just last week's sales. Nearly all
  of it is simply not last week's Stark County sales. Normal.
- **Weekly volume swings of ±40%.** 2026 has ranged 97–174 CAMA sales/week
  (mean ≈133). W33's 109 followed W31/W32 at 174/173, the year's two highest.
  Before suspecting the extract, check that (a) all five business days appear in
  the window — sales only record Mon–Fri — and (b) the drop is spread across
  cities rather than one city going to zero. In W33 both held, so the week was
  genuinely light, not broken. A single city collapsing to zero *would* point at
  `TAXDIST_TO_ZONE` or a filter. Closings also cluster at month-end, so a window
  covering the first week of a month runs light.

> **If step 2 fails with "✗ Matrix icon not found — cannot continue":** read
> `Automation Domination\matrix_icon_missing.html` *before* touching any selector.
> If it ends in `<!-- ends eula -->` and contains `doEulaUpdate()`, nothing is
> broken — MLS Now is serving its **End User License Agreement** and it sits in
> front of the portal, so there's no Matrix icon to find. The error message blames
> the wrong subsystem. Fix: log in at https://mdweb.mmsi2.com/mlsnow/ in a normal
> browser, accept the agreement, re-run. Acceptance is stored server-side, so it's
> once per agreement version. **Don't script the "I agree" checkbox** — that's a
> license acceptance on the county's behalf, not an automation's call. First hit
> 2026-08-03. A related tell: `⚠ still on login portal after submit (attempt 1)`
> is the same redirect confusing the login check, not a credentials problem.

**B. Zillow photos — Jason, solves CAPTCHAs:**
```powershell
.\run_zillow_photos.bat
```
Runs the downloader on value_mismatches then perfect_matches into `Photos_New`
(and builds `Photos_New_Portal` + the upload CSV). **Edit the two dates inside
the .bat each week** — they're hardcoded to one week's paths. Auto-skips parcels
that already have a photo, so it's safe to re-run after any interruption. For a
few stragglers, build a tiny retry `.xlsx` of just those rows and pass it to
`ZillowPhotos\download_zillow_photos.py <retry.xlsx> MLSvsCAMA\<week>\Photos_New`.

**C. Photo review** — fix bad exteriors in BOTH folders. → **§2.**

**D. Regenerate CSV + build — steps 6–7, NO push:**
```powershell
python "Automation Domination\run_build.py"
```
Runs the **photo gate** (see §2a) first, then regenerates the photo-upload CSV,
then builds `review_portal.html` + `version.json` + `archive/YYYY-WNN.json`.
Does **not** commit or push. If the gate finds a hard failure it exits 1 and
builds nothing — fix the photos and re-run.

**E. Sanity-check, then push** — verify the built HTML (see §4), then:
```powershell
git add review_portal.html version.json archive/<YYYY-WNN>.json
git commit -m "Portal update — Week of <Month DD, YYYY>"
git push origin main
```

After build + push, continue with the **future-year records (§5)** and the
**manual iasWorld steps (§6)** — identical in both modes.

### 1. Run the weekly pipeline (all-in-one)
```powershell
python "Automation Domination\run_weekly.py"
```
This runs 8 steps: CAMA extract → MLS export → Compare → Zillow photos → photo
verify → photo-upload CSV → build portal → git push. It **pauses** at three
human gates:

- **MLS export** — Playwright login (Auth0 SSO). Log in when prompted.
- **Zillow photos** — solve CAPTCHAs as they appear.
- **Photo-review gate** — a hard stop to eyeball every exterior. Aborts if there's
  no interactive terminal, so run it in a real console.

Sanity-check the Step 3 comparison summary (mismatches / perfects / missing)
before continuing. This week: 109 CAMA sales → **47 mismatch + 21 perfect
parcels** (68 unique), 41 missing in MLS.

Ignore these known non-issues:
- Auction/Kiko sales showing "missing in MLS" = posting lag, self-corrects.
- `likely_parcel_mismatches` file = advisory only, glance and dismiss.

### 2. Fix bad photos at the review gate
Zillow sometimes returns a **realtor headshot** or wrong house instead of an
exterior. For each bad one, look up its MLS number (match the parcel ID against
`value_mismatches_*.xlsx` / `perfect_matches_*.xlsx`, column `Listing_Number`),
pull the correct exterior, and save it as `<parcel>-1.jpg`.

**Save each swap ONCE, into either folder** — as of W34 the build syncs them:
- `MLSvsCAMA/<week>/Photos_New/` → feeds the iasWorld photo-upload CSV
- `MLSvsCAMA/<week>/Photos_New_Portal/` → feeds the portal build

`photo_checks.sync_primaries()` runs at the top of `run_build.py` (and right after
the review gate in `run_weekly.py`), propagates any `<parcel>-1.jpg` that exists in
only one folder or differs between them, and prints every file it moved. Direction
is by mtime, so the copy you just saved wins. Saving into both by hand is still
fine — identical bytes are a no-op.

**The photo gate will NOT catch a one-off headshot.** `photo_checks.py` flags a
headshot only when the *same image* is the lead on 2+ parcels. A single unique
headshot passes every hard check and ships to staff and to iasWorld. In W33,
`5218513` was a headshot the gate rated clean — it was caught only by eyeballing
the two parcels the geometry warning flagged as portrait/square. **Always open
the lead photos the geometry warning lists**; that warning is the only automated
hint you get. (W33's other two flagged leads, `4307614` at 320x320 and the
portrait phone shots, were genuine exteriors — the warning is a hint, not proof,
in both directions.)

W33 photo work, for shape: 1 headshot replaced (`5218513`), 4 primaries pulled
manually (`106726`, `2204333`, `211596`, `213306`), and 1 promoted from a
secondary (`5218304`, whose `-2` was a clean front exterior). Coverage went
63/68 → **68/68**.

Promoting a secondary to `-1` is a judgment call, not a default: it makes that
image the official primary in iasWorld. In W33 only `5218304` qualified —
`211596`'s secondary was a rear/side view and `213306`'s was an interior bedroom,
so both got manual pulls instead.

**Why a parcel shows up with secondaries but no primary** (fixed in W34).
`download_zillow_photos.py` used to reject photo 1 whenever it was portrait
(`h > w`), on the theory that headshots are portrait and houses are landscape.
When that fires, `-1` is written to *neither* folder while `-2/-3/-4` still land in
`Photos_New_Portal` — the "3 secondaries, no primary" signature, visible in the
mapping CSV as `Success` with `Photos = 3`. The rule was wrong both ways: in W34 it
**rejected two genuine houses** (`202027`, `209854` — 576x768 portrait phone shots
that then had to be pulled by hand) while **missing the actual headshot** on
`247844`, which was 240x240 — *square*, so `h > w` was false. The recurring W30/W32
headshot was the same 240x240.

The screen now requires **small AND near-square** (`< 20 KB or < 300 px on the long
edge`, aspect 0.9–1.1), which is the real signature of the headshot family.
Validated against all 166 lead photos across W33+W34: zero genuine exteriors
rejected, both headshot profiles caught.

**The thresholds are thin on purpose — do not loosen them casually.** The two
closest real data points are only:

| | dimensions | bytes | verdict |
|---|---|---|---|
| `247844` (W34) headshot | 240x240 | 8,860 | reject |
| `4307614` (W33) genuine house | 320x320 | 38,127 | keep |

A 40 KB floor — the obvious first guess — drops that real house. Because the margin
is this narrow, the **photo-review gate remains the actual defense**; this screen
only trims the obvious cases.

**Root cause, found in W34 — earlier versions of this guide had it wrong.** This
was described for six weeks (W22, W24, W30, W32, W33, W34) as a copy that
"silently does not take". There was no such copy. **No code path writes one photo
folder without the other**, so nothing existed to fail intermittently:
`download_zillow_photos.py:521-527` downloads photo 1 into `Photos_New_Portal` and
then *derives* `Photos_New` from it with `shutil.copy2`. The real data flow is
**Portal → Photos_New**, the reverse of what this guide used to imply. A manual fix
was therefore propagated only by the operator remembering to write the same file
into two directories, with nothing but a printed reminder enforcing it — and
`Photos_New`, which holds only `-1.jpg` files, reads as "the primaries folder"
where a replacement primary naturally belongs. Six weeks out of six is a process
gap, not a filesystem gremlin. `sync_primaries()` now closes it.

The verification below is still worth knowing, but it is now a **check that the
sync worked**, not the thing standing between a missed swap and the live portal:

```bash
# from MLSvsCAMA/<week>/
for f in Photos_New/*.jpg; do b=$(basename "$f"); \
  cmp -s "$f" "Photos_New_Portal/$b" || echo "MISSED: $b"; done
```
PowerShell equivalent (what actually gets used on the build machine — it also
reports files missing from the portal folder entirely, which the `cmp` loop
does not):

```powershell
$w = "<portal-root>\MLSvsCAMA\<week>"
Get-ChildItem "$w\Photos_New" -Filter *.jpg -File | ForEach-Object {
  $c = "$w\Photos_New_Portal\$($_.Name)"
  if (-not (Test-Path $c)) { "MISSING: $($_.Name)" }
  elseif ((Get-FileHash $_.FullName -Algorithm MD5).Hash -ne
          (Get-FileHash $c -Algorithm MD5).Hash) { "DIFFERS: $($_.Name)" }
}
```

Any line printed = a swap that only hit `Photos_New`; copy it into
`Photos_New_Portal` and re-check. (The build resizes at load time, so full-size
source photos in the portal folder are correct — no pre-resize needed.)

Expected steady state: `Photos_New` holds one `-1.jpg` per covered parcel plus
the downloader's `mapping_*.csv` files (those are *not* copied to the portal
folder — 2 "missing" CSVs is correct, not a fault). `Photos_New_Portal` holds
the same primaries plus `-2/-3/-4` secondaries. W33: 68 primaries + 2 CSVs in
`Photos_New`, 259 files in `Photos_New_Portal`, 0 differing.

### 3. Regenerate the photo-upload CSV after any manual photo work
```powershell
python "Automation Domination\make_photo_upload.py" --week <M-DD-YY>
```
It rescans `Photos_New`, so late downloads and manual swaps are picked up, and
stamps `Taxyr = create_future_year.TARGET_TAX_YEAR` (**2027**). `run_build.py`
runs it automatically as step 6, so a separate call is only needed if you fixed
photos after a build. W33: **68 rows @ Taxyr=2027**, one per covered parcel.

### 4. Build & push the portal
`run_weekly.py` does this automatically as steps 7–8. If the pipeline was paused
and you fixed photos out of band, build directly (mirrors steps 7–8): build
`review_portal.html`, write `version.json` with the **same** `generatedAt`
baked into the HTML (mismatched = permanent false reload banner), save
`archive/YYYY-WNN.json`, then commit + push those three files to `main`
(GitHub Pages source branch).

Before pushing, sanity-check the built HTML — a broken template ships a blank
navy screen. What to confirm, all of which has failed at least once:

- `https://api.anthropic.com/v1/messages` present and model id `claude-sonnet-4-6`
  (no dated snapshot like `claude-*-20250514` — those retire and silently break
  the Suggest button)
- `generatedAt` in the HTML **matches** `version.json` exactly, or staff get a
  permanent false reload banner
- `EFF_AGE_BASE_YEAR` is the current year
- no `unpkg.com` / `cdn.sheetjs.com` refs — libs must stay self-hosted in `lib/`
- no real API key. Note `sk-ant-` **does** legitimately appear twice, as the
  settings input placeholder and an error-message hint, so a bare grep for
  `sk-ant-` is a false alarm. Match `sk-ant-[A-Za-z0-9_\-]{6,}` instead; that
  should return nothing.

W33: 7.3 MB, 68 parcels, 259 photos, all checks clean, pushed as `bd0b18b`.

> If you built out of band, **press `Q`** in the paused `run_weekly.py` terminal
> so it doesn't build/push a duplicate commit.

### 5. Create the future-year (2027) records — BEFORE the manual iasWorld steps
All manual iasWorld steps below apply to the **2027** record, so the 2027 layer
must exist first.

```bash
# Bash tool, from Automation Domination/ :
echo YES | python create_future_year.py --env prod
```
Runs ~15–20 min, one unattended pass. **Don't touch the Chromium window it
opens.** W33: 74 parcels → **73 created, 1 `exists`, 0 errors.** The log lands at
`MLSvsCAMA/<week>/future_year_log_<date>.csv`.

The parcel count here is normally *higher* than the portal parcel count (W33: 74
vs 68) — multi-parcel sales need a 2027 record for each parcel, not one per sale.
An occasional `exists` is fine: iasWorld blocks a duplicate future-year layer, so
that parcel already had one.

Gotchas (both learned the hard way):
- Pipe the confirmation with **`echo YES | python …` in Bash**, NOT PowerShell's
  `Write-Output "YES" | python …` — the PowerShell pipe doesn't reach `input()`
  and the script aborts (non-destructive, but nothing happens).
- Piped stdout is **block-buffered** — a background run shows no interim progress
  until it exits. The log CSV is written all at once at the end. A quiet output
  file is NOT a hang.
- Re-runs are non-destructive: iasWorld blocks duplicate future-year layers and
  logs them as `exists`.

### 6. Manual iasWorld steps (on the 2027 records)
1. **Sale Tab Mass Update** — `SALETAB_MassUpdate_MMDDYY.xlsx`
2. **MassEntrance** — `MassEntranceMMDDYY.csv`
3. **Photo upload** — Document Loader with `Photo Upload M-D-YYYY.csv`
   (`Taxyr=2027`)

Sale Tab and MassEntrance have no tax-year column, so they attach to 2027
automatically. The photo-upload CSV is the only year-sensitive file — confirm
`Taxyr=2027` (step 3 handles this).

---

## Annual rollover (each January)
Bump `TARGET_TAX_YEAR` in `Automation Domination/create_future_year.py`
(single source of truth — `make_photo_upload.py` imports it) and roll
`EFF_AGE_BASE_YEAR` forward in both the template and the built portal.

---

## This week's result (2026-W33) — reference
- CAMA: 109 sales (Aug 1–7) · MLS: 2,827 listings, clean login, no EULA gate
- Portal: 47 mismatch + 21 perfect = **68 unique parcels** · 41 missing in MLS
- Photos: 259 embedded · coverage 63/68 → **68/68** after 1 headshot replaced,
  4 primaries pulled manually, 1 promoted from a secondary
- Portal-folder sync failed again (4 missing + 1 stale) — caught by hash, fixed
- Upload CSV: 68 rows @ Taxyr=2027
- Built 7.3 MB, pushed `bd0b18b`, live
- Future-year: 73 created / 1 exists / 0 errors of 74 parcels
- Volume note: 109 vs 173/174 the prior two weeks — verified genuine (all 5
  business days present, decline spread across all cities), not a pipeline fault
