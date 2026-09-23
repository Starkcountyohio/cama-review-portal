# Cloud Cutover — Dress Rehearsal Checklist

**Rehearsal: Friday, October 9, 2026.  First real cloud run: Monday, October 12, 2026.**

Goal: run the entire weekly chain against **Tyler cloud** once, before it counts, so
Oct 12 is a repeat and not a debut. This layers the *cloud-specific* concerns on top
of the normal pipeline — for the unchanged mechanics (gates, photo review, sanity
checks) follow `ONBOARDING.md`; this file only calls out what's different on cloud.

Local working doc — not committed. Check boxes as you go.

---

## 0. Prerequisites — verify BEFORE starting (the rehearsal is pointless without these)

- [ ] **Oracle/SQL cloud account in `credentials.py`.** The CAMA extract (pipeline
      step 1) connects to the **Oracle DB**, which is a *different* credential than the
      web UI. Update `ORACLE_USER` / `ORACLE_PASSWORD` / `ORACLE_DSN` to the cloud DB
      and confirm the **schema names are unchanged** (see `reference_iasworld_schema_queries`).
      Quick check (name only, no values): the extract should connect without error in step 1.
- [ ] **Updated (cloud) Document Loader installed** — for the photo upload step. If it
      hasn't arrived, see §4 contingency; the rest of the rehearsal can still proceed.
- [ ] **Cloud login still works.** The saved `cloud_session.json` (from 2026-09-21) may
      have expired. Confirm with:
      `python create_future_year.py --login-only`
      If it fails, refresh once: `python create_future_year.py --login-only --relogin`
      (or `--manual-login` if Microsoft throws a verification). A ✓ = you're set.
- [ ] **Decide the portal-flip push timing** (see §1-E). Default: validate locally Oct 9,
      push the flipped portal at the start of Oct 12 (or Fri evening if no staff will
      use the portal over the weekend).

---

## 1. The rehearsal run — full chain, rehearsal-safe

Run from the project root. Use the split/assisted mode so you can stop at each gate.

**A. Data pull (steps 1–3) — the #1 thing to prove.**
```powershell
python "Automation Domination\run_datapull.py"
```
- [ ] Step 1 **CAMA extract connects to the cloud Oracle** and returns a sane sale count
      (2026 weeks have run 97–174; see `ONBOARDING.md` §1 for what's normal vs alarming).
- [ ] Step 2 MLS export logs in (unchanged — Auth0; watch for the EULA gate per ONBOARDING).
- [ ] Step 3 compare summary looks reasonable (mismatch/perfect/missing).
- [ ] The 4 Excel files land in `MLSvsCAMA\<week>\`.

> If step 1 fails: it's the Oracle credentials or the cloud DB endpoint, **not** a
> generic connectivity problem. Fix `credentials.py` before anything else — this is the
> hard dependency for Oct 12.

**B. Zillow photos + C. Photo review** — unchanged from `ONBOARDING.md` §2. You can
abbreviate the review for the rehearsal, but at least run the downloader so the CSV/photo
folders exist for the build and the upload test.

**D. Build the portal WITH THE CLOUD FLIP ON — but do NOT push yet.**
- [ ] In `review_portal_template.html`, set `IASWORLD_CLOUD_DEFAULT = true`.
- [ ] Build (no push):
      `python "Automation Domination\run_build.py"`
- [ ] Sanity-check the built HTML per `ONBOARDING.md` §4 (Anthropic block, model id,
      `generatedAt` vs `version.json`, `EFF_AGE_BASE_YEAR`, no CDN refs, no real key).
- [ ] **Confirm the built portal points at cloud:** it should contain
      `atentcloud.tylerapp.com` in the parcel-link base and NOT the old
      `iasworld.starkcountyohio.gov`.

**E. Portal button validation (local, no push)** — serve locally and click through:
```powershell
# from the folder holding the freshly built review_portal.html:
python -m http.server 8765 --bind 127.0.0.1
```
Open `http://127.0.0.1:8765/review_portal.html` in Chrome (serve over http, NOT file://
— see `reference_local_portal_preview`).
- [ ] **iasWorld** button on a parcel opens the cloud Transact page and the parcel renders.
- [ ] **Zillow** button opens the right listing (allow pop-ups if blocked).
- [ ] Decide push timing (§0). Do **not** leave a cloud-URL portal live for staff before
      Oct 12 — their buttons would be dead until cloud go-live.

**F. Future-year records — validate, don't mass-create.**
```powershell
python create_future_year.py --dry-run --limit 1 --diag
```
- [ ] Login → hamburger → Transaction → Multi-Use (new window) → parcel loads →
      `dry_run_would_create`. (A real `--limit 1` create is already proven from 2026-09-22;
      save the full 70-parcel run for Oct 12 so you don't create records prematurely.)

---

## 2. Cloud-specific validations (the new / risky surface — the heart of the rehearsal)

- [ ] **Oracle extract on cloud** — proven in §1-A. Note the sale count for comparison Oct 12.
- [ ] **Portal cloud iasWorld button** — proven in §1-E (real parcel renders).
- [ ] **create_future_year login + nav + Multi-Use** — proven in §1-F.
- [ ] **Photo mass-import via the new Document Loader** — load the week's
      `Photo Upload M-D-YYYY.csv` (`Taxyr=2027`) through the *updated* Document Loader for
      **one or two test parcels**. Confirm the photo attaches to the **2027** record.
      This is the piece we know least about — give it real attention.
- [x] **SALETAB Mass Update on cloud** — VERIFIED WORKING (Jason tested 2026-09-23).
      No action needed at the rehearsal beyond the normal run.
- [x] **MassEntrance on cloud** — VERIFIED WORKING (Jason tested 2026-09-23).
      No action needed at the rehearsal beyond the normal run.
- [ ] **windowId behavior** — note whether the cloud windowId rotates daily or per-login,
      so the portal help-text wording can be made specific if needed.

---

## 3. Go / No-Go for Oct 12

**GO if:** Oracle extract works · portal cloud buttons render parcels · future-year
login/nav works · Document Loader accepts the photo CSV. (SALETAB Mass Update +
MassEntrance already verified working 2026-09-23; future-year login/nav/create already
proven 2026-09-22.) Then Oct 12 = run it for real (§5).

The only genuinely open items left for Oct 12: the **Oracle account** (coming) and the
**photo Document Loader** (pending the updated tool). Everything else is proven.

**NO-GO / partial** if any of the above fails — fix before Oct 12, or use the §4
contingency for that piece and flag it as a known manual fallback for the first run.

---

## 4. If a prerequisite isn't ready by Oct 9

- **No Oracle account yet** → the data pull can't run; the rehearsal is blocked on the
  one thing that matters most. Escalate immediately — this is the Oct 12 critical path.
- **No updated Document Loader yet** → rehearse everything else; for photos, the fallback
  is per-parcel manual import in the cloud UI (severe regression at 70/wk, so chase the
  bulk tool). Re-point `make_photo_upload.py` column names only if the new loader's CSV
  differs.
- **Cloud login throws MFA/verification** → `python create_future_year.py --manual-login`
  once, complete it by hand, session is saved for the rest of the day.

---

## 5. Oct 12 (first real run) — what changes from the rehearsal

Same chain, with the rehearsal's "safe" swaps turned real:
- [ ] Full pipeline (`run_weekly.py` or split), real data pull.
- [ ] Portal built with the flip ON and **pushed live** (staff now on cloud):
      `git add review_portal.html version.json archive/<YYYY-WNN>.json`
      `git commit -m "Portal update — Week of October 12, 2026 (cloud cutover)"`
      `git push origin main`  ·  also commit the flipped `review_portal_template.html`.
- [ ] **Full** future-year run: `echo YES | python create_future_year.py` (Bash), all parcels.
- [ ] Manual iasWorld steps on the 2027 records (SALETAB → MassEntrance → photo upload),
      per `ONBOARDING.md` §6, using whatever was validated in §2.
