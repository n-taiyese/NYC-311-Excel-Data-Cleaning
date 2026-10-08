# 🧹 NYC 311 Data Cleaning Project — Excel

**Tech Stack:** Microsoft Excel (formulas, built-in tools, Data Validation, Conditional Formatting)

**Dataset:** 100,000 real NYC 311 service request records — noise complaints, potholes, missed trash pickups, and more, filed with the City of New York.

**Source:** [NYC Open Data — 311 Service Requests from 2010 to Present](https://data.cityofnewyork.us/Social-Services/311-Service-Requests-from-2010-to-Present/erm2-nwe9/about_data) (public domain, updated daily)

---

## 📊 Business Problem

New York City receives millions of 311 service requests, and the raw export is exactly what an analyst meets in the real world: blank fields, inconsistent text, mixed date formats, and values that can't possibly be right. This project takes the raw file through a full cleaning workflow — 10 techniques, applied in order — to produce an analysis-ready dataset.

The point of this project is the *process*, not just the result. Each technique below is a core data-cleaning skill every data analyst is expected to know.

---

## 🗂️ Workbook Structure

| Sheet | Purpose |
|---|---|
| `Raw` | Untouched source data (100,000 rows). Never edited — this is the audit trail. |
| `Working` | Helper columns showing each fix step by step, so the cleaning is fully traceable. |
| `Clean` | Final cleaned dataset — output only, ready for analysis. |
| `Cleaning Log` | One row per technique: issue found, records affected, decision made. |
| `Data Dictionary` | Every column defined, with its expected format and data type. |

**Rule of thumb:** always work on a copy. `Raw` stays pristine; all cleaning happens in `Working`; `Clean` holds the result.

---

## 🧼 The 10 Cleaning Techniques (applied in order)

The order matters. Each step protects the next — for example, you standardize text *before* deduplicating, because `"BROOKLYN "` and `"BROOKLYN"` look like different records until the spaces are gone.

### 1. Profiling and blank audit (basic)
**What it is:** Before changing anything, measure the mess. Count blanks per column with `COUNTA` against the total row count, and use Data > Filter to scan distinct values in categorical columns.

**How it works:** Blanks are information. In this dataset, `closed_date` is blank for every open complaint — that's meaningful, not broken. `incident_zip` blanks are genuinely missing. The audit tells you which blanks to keep, which to fix, and which to flag, before you touch a single cell.

### 2. Removing duplicates (basic)
**What it is:** Real exports contain duplicated rows. Use `COUNTIF` on `unique_key` to flag them, then Data > Remove Duplicates.

**How it works:** Pick the column (or column combination) that should be unique — here, `unique_key`. `=COUNTIF($A:$A, A2)>1` flags every repeat. Remove Duplicates then drops the extras. Record how many rows were removed in the Cleaning Log: "removed 312 duplicates (0.3%)" is exactly the kind of line a recruiter reads.

### 3. Trimming whitespace (basic)
**What it is:** Stripping invisible leading, trailing, and double spaces from text fields.

**How it works:** `=TRIM(A2)` removes extra spaces; pair it with `=CLEAN(A2)` to also remove non-printing characters that sneak in from system exports. Do this on `city`, `street_name`, `complaint_type`, and other text columns. This is the fix that makes filters, PivotTables, and lookups behave — trailing spaces silently break all three.

### 4. Standardizing text case (basic)
**What it is:** Forcing consistent casing within each text column.

**How it works:** Decide one convention per column and document it — e.g., `city` in UPPER (`=UPPER(A2)`), names in Proper Case (`=PROPER(A2)`). Excel's Flash Fill (Ctrl+E) can also learn the pattern from two or three examples. After converting, paste as values so the formulas don't travel downstream.

### 5. Normalizing inconsistent categories (basic → intermediate)
**What it is:** Collapsing the many ways of saying the same thing into one standard value.

**How it works:** `status` and `open_data_channel_type` contain values like `"UNKNOWN"`, `"N/A"`, and `"Unspecified"`. Build a small mapping table on a separate sheet — raw value in one column, standard value in the next — and apply it with `=XLOOKUP(A2, Mapping[Raw], Mapping[Standard])`. The analyst habit this teaches: never hardcode fixes inside formulas; map them in a table you can audit and change.

### 6. Converting data types (intermediate)
**What it is:** Making sure dates are dates and numbers are numbers.

**How it works:** `created_date` arrives as text (`2026-10-06T02:05:43.000`). Convert with `=DATEVALUE(LEFT(A2,10)) + TIMEVALUE(MID(A2,12,8))`, or use Data > Text to Columns with a date format. The reverse problem: `incident_zip` as a number strips leading zeros — `"07030"` becomes `7030`. Format zip columns as Text *before* import (or use `=TEXT(A2,"00000")`). Wrong types silently corrupt sorts, date math, and joins.

### 7. Splitting compound fields (intermediate)
**What it is:** Breaking one overloaded column into its component parts.

**How it works:** The `location` column holds `"(40.71, -73.99)"` — two facts in one cell. Split it with Data > Text to Columns (delimited by comma), or with `=TEXTBEFORE(A2,",")` and `=TEXTAFTER(A2,",")`, then strip the parentheses with `SUBSTITUTE`. Each column should hold one atomic value — that's first normal form, and it's what makes the data analyzable.

### 8. Validating date logic (intermediate)
**What it is:** Checking that dates make sense relative to each other, not just that they're formatted correctly.

**How it works:** Add a helper column: `=closed_date - created_date` (resolution time in days). Any negative result is impossible — a complaint can't close before it's created. Flag these with Home > Conditional Formatting > Highlight Cell Rules, then decide the business rule: exclude the rows, or flag them for review. Document the decision in the Cleaning Log. This is how analysts catch errors no format check can see.

### 9. Detecting outliers in numerics (intermediate)
**What it is:** Finding numeric values that fall outside any plausible range.

**How it works:** Get `MIN` and `MAX` for `latitude` and `longitude`. Valid NYC coordinates sit roughly in 40.4–40.95 / -74.3–-73.65 — anything outside is a data entry error. Add a bounds-check helper column (`=OR(B2<40.4, B2>40.95)`) and filter on it. Same principle as a statistical IQR test, done with Excel basics. Decide: correct, flag, or remove — and log why.

### 10. Locking the clean data with validation (intermediate)
**What it is:** Preventing the mess from coming back.

**How it works:** On the `Clean` sheet, select `borough` and `status`, then Data > Data Validation > List, pointing at your standard value lists. Future entries are restricted to the dropdown — no more `"UNKNOWN"` vs `"Unknown"` vs `"N/A"`. Cleaning is half the job; preventing re-contamination is the other half, and this is the technique that shows you think like a data steward, not just a one-time fixer.

---

## 📝 Data Dictionary (key columns)

| Column | Description | Expected format |
|---|---|---|
| `unique_key` | Unique complaint identifier | Integer, no duplicates |
| `created_date` | When the complaint was filed | Date + time |
| `closed_date` | When it was resolved | Date + time (blank if still open) |
| `agency` / `agency_name` | Responsible city agency | Standardized text |
| `complaint_type` / `descriptor` | What the complaint is about | Standardized text |
| `incident_zip` / `incident_address` | Where it happened | 5-digit text zip |
| `city` / `borough` | Borough of the incident | Standardized text |
| `status` | Open, Closed, In Progress, etc. | Dropdown-validated |
| `latitude` / `longitude` | Coordinates | Numeric, within NYC bounds |
| `location` | Combined lat/long (source field) | Split during cleaning |

---

## 🚀 How to Use This Project

1. Open `NYC_311_Raw_100k.xlsx` — the `Raw` sheet is your starting point.
2. Work through techniques 1–10 in order, using `Working` for helper columns.
3. Fill in the `Cleaning Log` as you go — what you found, what you changed, how many records were affected.
4. Publish the finished workbook alongside this README.

## 📌 Notes

- The dataset is real and public domain (NYC Open Data). Some messiness is inherent to how the data is collected — that's the point.
- Techniques are ordered deliberately: profile → dedupe → text fixes → type fixes → structural fixes → logic checks → prevention.
- Excel-only: no Python, no Power Query. Every technique uses formulas, built-in tools, or Data Validation.

---

*Project by Taiyese Olamilekan · Dataset: NYC Open Data (public domain) · Built October 2026*
