# TALLY ML — Streamlit

A Streamlit version of the TALLY ML Excel → Tally XML converter.

It **reuses the original Django app's processing logic** (`../TALLY ML/GST/GSTapp/views.py`)
unchanged, so the generated Tally XML/Excel is identical — only the UI is new.

## Setup

```bash
cd streamlit_tally
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

Opens at http://localhost:8501. To share on your local network:

```bash
streamlit run app.py --server.address 0.0.0.0
```

## How it works

The app imports the Django view functions and calls them with a tiny
`FakeRequest` (just `.FILES` + `.session`). Each converter:

1. **Step 1** — download the blank Excel template.
2. **Step 2** — upload the filled template; download the resulting
   `.xml`, `.xlsx`, or a `.zip` of both.

### Project location

By default the app expects the original Django project as a sibling:

```
TALLY_ML/
├── TALLY ML/GST/        <- original Django project (manage.py, GSTapp/)
└── streamlit_tally/     <- this app
```

If you move things, point the app at the project folder (the one with
`manage.py`):

```bash
TALLY_DJANGO_DIR="/path/to/TALLY ML/GST" streamlit run app.py
```
