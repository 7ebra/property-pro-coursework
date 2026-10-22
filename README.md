# Property Pro²

A university Python desktop prototype for browsing commercial property listings,
filtering them against client requirements and exporting the results to Excel.
Developed for a Principles of Programming module, with a CustomTkinter interface
and openpyxl reporting. This is a coursework demonstration using sample data.

## What it demonstrates

- Separate user and administrator login screens and dashboards.
- Filters for property type, budget, size, parking, floors, availability and rating.
- Administrator forms for adding properties and creating administrator accounts.
- Property details, favourites and simple ranking by rating and view count.
- Formatted `.xlsx` reports with search criteria, result counts and property data.
- A written design and manual test plan in the original university submission.

The recommendation function is deterministic sorting, not machine learning.
User and administrator screens demonstrate different workflows; they do not
constitute a hardened authorization boundary.

## Run locally

Use Python 3.10+ with Tk support. Windows is the platform described in the
original coursework. From this directory:

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python app.py
```

On Linux or macOS, use `.venv/bin/python` instead. Some Linux distributions
require their Python Tk package to be installed separately. The dependencies
are pinned for this portfolio package; the submission did not record versions.

Demo user: `demo@example.invalid` / `demo-password`.
Demo administrator: `Boss` / `demo-admin-password`.
These are deliberately public, local demonstration credentials. Do not enter
real credentials or personal information into this app.

## Check the core behavior

```powershell
python -m pip install openpyxl==3.1.5
python -m unittest discover -s tests -v
```

The tests load the data and core functions from `app.py` using Python's AST,
without opening a desktop window. They execute the actual filter and report
functions; only the save-file dialog is replaced with a temporary destination.
They also compile the whole source. They do not verify GUI rendering,
CustomTkinter compatibility, login interactions or complete user journeys.

## Known limitations

- Accounts, favourites and properties are in memory and reset on restart.
- Passwords are stored and compared as plain text; demo logins are prefilled.
- Administrative operations have no centralized authorization checks.
- Several settings and contact actions are placeholders. Their dialogs now
  explicitly say that the operation is unimplemented.
- Availability matching is exact: searching `Lease` does not also include `Both`.
- Numeric conversion catches some malformed input, but domain validation is
  incomplete; negative values and non-finite numbers are not consistently rejected.
- Recommendations rank the top five before applying the optional budget filter.
- Property names are written directly into spreadsheets. Formula-leading input
  is not escaped, so export is limited to trusted demonstration data.
- The code is a single large script. This is not a production service or a
  deployed customer system.

The original report documents 16 manual cases, including an analytics failure
and a later repair. Those historical results are author-reported. The current
core checks are separate and do not revalidate every manual case.

## Portfolio context

This is supporting evidence of Python application development, GUI work,
requirements analysis and spreadsheet automation. For cybersecurity applications,
my [Causal Research Workbench](https://github.com/7ebra/causal-research-workbench)
and [Eventify](https://github.com/7ebra/eventify-code) provide stronger examples
of validation, credential handling and web application controls.

The next useful improvements would be persistent storage, password hashing,
centralized authorization, stricter validation and a separation of GUI and logic.
These improvements are not implemented here.

## Source and preparation

Author: Zain Ul Abideen. See [SOURCE.md](SOURCE.md) for the relationship to the
original submission and the limited publication edits. The source already
acknowledges AI assistance with interface taglines. Portfolio preparation and
the new core checks were completed with Codex assistance; this repository does
not claim entirely unaided authorship.

The university report, student identifier, third-party images, bundled virtual
environment, editor settings and generated exports are excluded from this repo.
