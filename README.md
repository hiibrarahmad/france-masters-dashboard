# France Master's 2027 – program dashboard

An interactive dashboard of **999 master's / engineering / Mastère Spécialisé programs** from the
[Études en France](https://etudesenfrance.diplomatie.gouv.fr/catalogue-formations) catalogue that fit an
embedded-systems / biomedical-device engineering profile, for the **September 2027** intake.

**Open `index.html` in any browser** – no install needed. Private online version: https://claude.ai/artifact/Dpmc2uaXaxRTcK5Pr3RigK (shared by invitation only).

## What it does

- **Find programs** – one-click presets (*Best for me*, *English-taught & affordable*, *Biomedical / EEG*,
  *Embedded & electronics*), an **English only** checkbox, search, and filters for eligibility, language, fee, field,
  relevance and institution type.
- **Warnings on anything suspicious** – ⛔ serious, ⚠ check, ℹ note (e.g. fee understated, entry year closed,
  English listed but French required, website down, apprenticeship-only, outdated catalogue text).
- **My shortlist** – star programs, track your application stage and notes (saved in your browser), export to CSV.
- **Overview** – clickable charts by field, fee band, city and warning type.
- **Fees & exemptions** – verified 2026-27 fee rules and each university's exemption policy, with sources.
- **How to apply** and **What was checked** (the full verification report).

## How the data was checked (9 Oct 2026)

- All 999 listings re-downloaded from the live Études en France catalogue (998 succeeded; 1 server error, flagged).
- National fees confirmed on service-public.gouv.fr: master €3,950 non-EU / €255 EU (2026-27);
  Decree 2026-385 caps exemptions at 30% → **25% for 2027-28** → 20%.
- Corrections to the original spreadsheet:
  - EEF "entry level" is the level of the year you join (M1 = Bac+4, M2 = Bac+5), so 260 programs were re-classed
    from "needs a Master" to "M2 entry possible", and years-to-finish were recalculated.
  - 33 fee estimates raised from the EU rate (€6xx) to the non-EU €3,950.
  - Exemption policy updates for Université Paris Cité (2027-28 change), Grenoble Alpes and Tours.

## Refresh the data

```bash
pip install openpyxl
python scripts/build.py --refresh --check-web   # re-download from EEF + re-test websites (~5–10 min)
```

This regenerates `data/programs.js`. Fee rules and policy updates live at the top of `scripts/build.py`.

> Fees and policies change every year. Always confirm on the school's own website before paying anything.
