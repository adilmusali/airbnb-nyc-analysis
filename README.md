# NYC Airbnb Market Analysis

**The purpose** of this analysis is to clean and analyze New York City Airbnb listings from three messy source files (CSV, Excel, and TSV). The project downloads **the latest** public [Inside Airbnb](http://insideairbnb.com/) data, **reshapes** it into practice-ready files, then walks through **ingesting**, **cleaning** and **analyzing** in a Jupyter notebook.

---

## What you get


| Step         | File                        | Role                                                            |
| ------------ | --------------------------- | --------------------------------------------------------------- |
| Prepare data | `prepare_data.py`           | Downloads NYC listings and writes three raw files under `data/` |
| Analyze      | `airbnb_nyc_analysis.ipynb` | Merges, cleans, and explores the market                         |


**Raw inputs produced by the prep script:**


| File                          | Format | Columns                                  |
| ----------------------------- | ------ | ---------------------------------------- |
| `data/airbnb_price.csv`       | CSV    | `listing_id`, `price`, `nbhood_full`     |
| `data/airbnb_room_type.xlsx`  | Excel  | `listing_id`, `description`, `room_type` |
| `data/airbnb_last_review.tsv` | TSV    | `listing_id`, `host_name`, `last_review` |


Prices look like `"225 dollars"`, room types have mixed casing, and review dates look like `"August 3 2026"` — so the notebook has real cleaning work to do.

---

## Requirements

- Python 3.10+ recommended
- Internet access (only needed when running `prepare_data.py`)

---

## Quick start

### 1. Clone and enter the project

```bash
cd airbnb
```

### 2. Create a virtual environment

**Windows (PowerShell):**

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

**macOS / Linux:**

```bash
python -m venv .venv
source .venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Download and prepare the data

```bash
python prepare_data.py
```

This looks up the **latest** Inside Airbnb NYC snapshot, downloads it, and writes the three raw files under `data/`. If you already have that snapshot locally, the script skips the download unless you force a rebuild.

**Refresh when new data is published:**

```bash
python prepare_data.py
```

Useful options:

```bash
python prepare_data.py --force                 # rebuild even if already current
python prepare_data.py --snapshot 2026-08-10   # pin a specific scrape date
```

### 5. Open the notebook

```bash
jupyter notebook airbnb_nyc_analysis.ipynb
```

Run all cells from top to bottom after refreshing data.

> **Note:** `data/` is gitignored. Always run `prepare_data.py` before opening the notebook if the folder is empty.

---

## Project layout

```
airbnb/
├── prepare_data.py              # Download + split into messy raw files
├── airbnb_nyc_analysis.ipynb    # Cleaning and market analysis
├── requirements.txt             # Python dependencies
├── data/                        # Generated locally (not committed)
│   ├── airbnb_price.csv
│   ├── airbnb_room_type.xlsx
│   ├── airbnb_last_review.tsv
│   └── snapshot_meta.json       # Last downloaded snapshot date/URL
└── README.md
```

---

## Notebook walkthrough

The notebook is organized into seven sections:

1. **Ingest** — Load CSV, Excel, and TSV; inspect shapes, types, and missing values.
2. **Combine** — Inner-merge all three sources on `listing_id`.
3. **Clean strings** — Strip `" dollars"` from price, normalize `room_type` casing, split `nbhood_full` into `borough` and `neighbourhood`.
4. **Format dates** — Parse `last_review` text into datetime.
5. **Handle gaps** — Drop duplicate IDs; exclude missing prices/reviews from related metrics.
6. **Analysis** — Average price overall and by borough/room type, room-type mix, review recency, and price charts.
7. **Summary table** — One-row snapshot of key market metrics.

---

## Sample findings

From a recent run on the prepared snapshot (~30,234 listings):


| Metric                          | Value               |
| ------------------------------- | ------------------- |
| Listings with a price           | 20,331              |
| Average price                   | ~$267               |
| Median price                    | ~$175               |
| Most common room type           | Entire home/apt     |
| Highest average price (borough) | Manhattan           |
| Review date range               | May 2011 → Aug 2026 |
| Reviews in last 90 days         | ~6,700              |


Numbers will shift if Inside Airbnb publishes a newer snapshot and you re-run `prepare_data.py`.

---

## Dependencies


| Package      | Used for                         |
| ------------ | -------------------------------- |
| `pandas`     | Data loading, cleaning, analysis |
| `openpyxl`   | Read/write Excel (`.xlsx`)       |
| `jupyter`    | Run the analysis notebook        |
| `matplotlib` | Charts                           |
| `seaborn`    | Chart styling                    |


---

## Data source

Listings come from [Inside Airbnb](http://insideairbnb.com/) New York City visualisations data. `prepare_data.py` reads the public [Get the Data](http://insideairbnb.com/get-the-data.html) index, picks the newest `new-york-city/YYYY-MM-DD` folder, and downloads:

```
https://data.insideairbnb.com/united-states/ny/new-york-city/{snapshot}/visualisations/listings.csv
```

The active snapshot is recorded in `data/snapshot_meta.json` after each successful run.