# Internal Audit Observations Dashboard

A Streamlit dashboard that turns an Excel schedule of audit observations into an Audit Committee pack:
auto-detected internal-control areas, risk ratings, department comparisons, an action tracker, an
auto-generated executive summary and a Word committee report.

## What's in this repository

```
app.py                         the dashboard (single file)
requirements.txt               Python packages Streamlit Cloud installs
.streamlit/config.toml         colour theme and upload limit
.streamlit/secrets.toml.example  how to set an access password
data/observations.xlsx         your observations in the recommended template (optional, see Confidentiality)
```

## Deploy on GitHub + Streamlit Community Cloud (no coding needed)

1. **Create a private repository on GitHub.** github.com → New repository → name it (e.g. `audit-dashboard`)
   → choose **Private** → Create.
2. **Upload the files.** In the new repository click *Add file → Upload files* and drag in `app.py`,
   `requirements.txt`, `README.md` and the `data` folder. Then create the theme file: *Add file → Create new
   file*, type `.streamlit/config.toml` as the name, and paste in the contents of that file. Commit.
3. **Deploy.** Go to share.streamlit.io, sign in with GitHub, click *Create app*, pick the repository, branch
   `main`, main file `app.py`, and deploy. Allow Streamlit access to private repositories when prompted.
4. **Add a password (strongly recommended).** In Streamlit Cloud open the app → *Settings → Secrets* and paste:
   `APP_PASSWORD = "your-password"`. The app will ask for it before showing anything.
5. **Restrict viewers.** In the app's *Share* settings, make the app private and invite only the people who
   should see it (audit team, Committee members).

To update the app later, edit or replace files in GitHub; Streamlit redeploys automatically.

### Run on your own computer (optional)

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Using the dashboard

- **Upload** one or more `.xlsx`, `.xls` or `.csv` files in the sidebar, for example one file per department
  or per assignment. Each sheet is detected as either observations or audit queries. Header rows and title
  rows above them (like `ASSIGNMENT: ORACLE - ERP`) are picked up automatically.
- **Department for rows without one** fills blanks. Add a *Department* column to unlock real comparisons.
- **Position as at** drives overdue and ageing calculations.
- **Filters** apply to every page and to the downloaded reports.

| Page | What it shows |
|---|---|
| Overview | Risk strip, KPIs, executive summary, key matters, rating mix, control-area chart, heat map, COSO coverage |
| Control areas | Treemap of weaknesses, cross-cutting themes, design vs operating split, drill-down per area |
| Department comparison | Compare by department, assignment, entity, process or control area; exposure index, heat map, radar |
| Action tracker | Status, due-date profile, action timeline, owner workload, overdue list |
| Observation register | Every observation in criteria / condition / cause / effect format with data gaps flagged |
| Recommendations & actions | Write your recommendations, confirm ratings, record management action plans, owners, dates |
| Audit queries | Pending fieldwork queries, linked to control areas |
| Committee report | Word report, Excel register with analysis sheets, CSV |
| Data quality & template | Field completeness, what's missing per observation, blank input template |

**Saving your work.** Streamlit Cloud does not store edits permanently. After editing, click
*Download updated register (.xlsx)*. The *Observations* sheet in that file is re-uploadable, so it becomes
your master copy. To make it the default dataset, replace `data/observations.xlsx` in GitHub with it.

## Recommended columns

The app works with whatever columns you have and fills gaps with auto-suggestions, but for committee-grade
reporting add: Department, Observation Title, Criteria, Cause, Risk / Impact, Risk Rating, Recommendation,
Management Response, Management Action Plan, Action Owner, Timeline, Status. Useful extras: Entity,
Process, Likelihood and Impact (1-5), Finding Type, Repeat Finding, Financial Impact (PKR), Revised Timeline,
Date Raised, Closure Date, Evidence Ref, Auditor. The *Data quality & template* page explains each one and
downloads a template with dropdowns.

## How the automation works

- **Control areas**: weighted keyword matching across 14 internal-control areas, each mapped to a COSO 2013
  component. A value you put in *Control Area* always wins.
- **Themes**: patterns for recurring root causes such as missing policies, delays, approval gaps, breaches,
  records not maintained and resourcing gaps.
- **Risk rating**: 5 × 5 likelihood × impact. Impact starts from the area's inherent impact (+1 for
  regulatory, breach, loss or fraud language). Likelihood starts at 3 (+1 operating failure, +1 repeat,
  +1 pervasive, −1 when the finding reads as a note or query). Score ≥ 20 Critical, 12–19 High, 6–11 Medium,
  ≤ 5 Low. Anything in *Risk Rating*, *Likelihood* or *Impact* overrides it, and auto-suggested ratings are
  labelled as such everywhere, including the Word report.

Everything is rule-based and runs inside the app: no data leaves Streamlit, and no AI service or API key is
used. To tune it, edit the `AREAS` and `THEMES` lists near the top of `app.py`.

## Confidentiality

Audit findings are sensitive. Keep the GitHub repository **private**, set `APP_PASSWORD`, and restrict app
viewers. If you prefer not to store findings on GitHub at all, delete `data/observations.xlsx` and upload
the file each time you open the app.
