"""Download Inside Airbnb NYC listings and split into three messy raw files."""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook

DATA_URL = (
    "https://data.insideairbnb.com/united-states/ny/new-york-city/"
    "2026-08-10/visualisations/listings.csv"
)
DATA_DIR = Path(__file__).resolve().parent / "data"
RNG = np.random.default_rng(42)


def download_listings() -> pd.DataFrame:
    print(f"Downloading listings from:\n  {DATA_URL}")
    df = pd.read_csv(DATA_URL)
    print(f"Downloaded {len(df):,} rows, {len(df.columns)} columns")
    return df


def mess_up_room_type(series: pd.Series) -> pd.Series:
    """Apply inconsistent casing so the notebook has real cleaning work."""
    styles = ["title", "upper", "lower", "as_is"]
    choices = RNG.choice(styles, size=len(series))
    out = []
    for value, style in zip(series, choices):
        if pd.isna(value):
            out.append(value)
            continue
        text = str(value)
        if style == "title":
            out.append(text.title())
        elif style == "upper":
            out.append(text.upper())
        elif style == "lower":
            out.append(text.lower())
        else:
            out.append(text)
    return pd.Series(out, index=series.index, name=series.name)


def format_review_date(value) -> str | float:
    if pd.isna(value):
        return np.nan
    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        return np.nan
    # e.g. "August 3 2026" (no comma, day without leading zero)
    return f"{dt.strftime('%B')} {dt.day} {dt.year}"


def listing_ids_as_text(df: pd.DataFrame) -> pd.Series:
    """Keep IDs as strings so Excel does not round 18+ digit values."""
    return df["id"].astype("int64").astype(str)


def build_price_csv(df: pd.DataFrame) -> pd.DataFrame:
    price = df["price"].copy()
    # Keep missing prices as missing; otherwise append " dollars"
    price_text = price.apply(
        lambda x: f"{int(x)} dollars" if pd.notna(x) else np.nan
    )
    nbhood_full = (
        df["neighbourhood_group"].astype(str)
        + ", "
        + df["neighbourhood"].astype(str)
    )
    # If either part was NaN, pandas turns it into the string "nan" — restore missing
    missing_nbhood = df["neighbourhood_group"].isna() | df["neighbourhood"].isna()
    nbhood_full = nbhood_full.mask(missing_nbhood, np.nan)

    return pd.DataFrame(
        {
            "listing_id": listing_ids_as_text(df),
            "price": price_text,
            "nbhood_full": nbhood_full,
        }
    )


def build_room_type_xlsx(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "listing_id": listing_ids_as_text(df),
            "description": df["name"],
            "room_type": mess_up_room_type(df["room_type"]),
        }
    )


def build_last_review_tsv(df: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "listing_id": listing_ids_as_text(df),
            "host_name": df["host_name"],
            "last_review": df["last_review"].map(format_review_date),
        }
    )


def main() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    listings = download_listings()

    price_df = build_price_csv(listings)
    room_df = build_room_type_xlsx(listings)
    review_df = build_last_review_tsv(listings)

    price_path = DATA_DIR / "airbnb_price.csv"
    room_path = DATA_DIR / "airbnb_room_type.xlsx"
    review_path = DATA_DIR / "airbnb_last_review.tsv"

    price_df.to_csv(price_path, index=False)
    room_df.to_excel(room_path, index=False)
    # Force listing_id to Excel text so large IDs are not rounded
    wb = load_workbook(room_path)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    id_col = headers.index("listing_id") + 1
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=id_col)
        cell.number_format = "@"
        cell.value = str(cell.value)
    wb.save(room_path)
    review_df.to_csv(review_path, sep="\t", index=False)

    print(f"Wrote {price_path}  ({len(price_df):,} rows)")
    print(f"Wrote {room_path}  ({len(room_df):,} rows)")
    print(f"Wrote {review_path}  ({len(review_df):,} rows)")
    print("\nSample price values:", price_df["price"].dropna().head(3).tolist())
    print("Sample room_type values:", room_df["room_type"].dropna().unique()[:6].tolist())
    print("Sample last_review values:", review_df["last_review"].dropna().head(3).tolist())
    print("Sample nbhood_full:", price_df["nbhood_full"].dropna().head(2).tolist())


if __name__ == "__main__":
    main()
