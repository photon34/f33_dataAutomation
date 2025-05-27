from __future__ import annotations

import os
import re
import argparse
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

# --------------------------------------------------------------------------- #
# Funding helpers
# --------------------------------------------------------------------------- #

_SUFFIX = {"K": 1e3, "M": 1e6, "B": 1e9}


def _parse_funding(raw) -> float | np.nan:
    if pd.isna(raw):
        return np.nan
    text = (
        str(raw)
        .replace(",", "")
        .replace("$", "")
        .replace("€", "")
        .replace("£", "")
        .strip()
    )
    m = re.fullmatch(r"([\d.]+)\s*([KMB])?", text, flags=re.I)
    if not m:
        return np.nan
    number = float(m.group(1))
    multiplier = _SUFFIX.get(m.group(2).upper() if m.group(2) else "", 1)
    return number * multiplier


def _human_funding(usd: float | np.nan) -> str:
    if pd.isna(usd):
        return ""
    if usd >= 1e9:
        return f"${usd/1e9:.1f}B"
    if usd >= 1e6:
        return f"${usd/1e6:.1f}M"
    if usd >= 1e3:
        return f"${usd/1e3:.0f}K"
    return f"${usd:,.0f}"


# --------------------------------------------------------------------------- #
# Core routine
# --------------------------------------------------------------------------- #
def top_funded_companies(
    path: str | Path,
    top_n: int = 5,
    output_path: Optional[str | Path] = None,
) -> pd.DataFrame:
    path = Path(path)
    ext = path.suffix.lower()

    if ext == ".csv":
        df = pd.read_csv(path)
    elif ext in (".xls", ".xlsx"):
        df = pd.read_excel(path)
    else:
        raise ValueError(f"Unsupported input format: {ext}")

    df = df.rename(columns={c: c.strip() for c in df.columns})

    required = {
        "Company",
        "Recent Funding Amount",
        "Using cloud marketplaces?",
    }
    if not required.issubset(df.columns):
        missing = ", ".join(sorted(required - set(df.columns)))
        raise KeyError(f"Missing required column(s): {missing}")

    df["Company"] = df["Company"].astype(str).str.strip().str.title()
    df["FundingUSD"] = df["Recent Funding Amount"].apply(_parse_funding)
    df = df.dropna(subset=["Company", "FundingUSD"])

    grouped = (
        df.groupby("Company", as_index=False)
        .agg(
            FundingUSD=("FundingUSD", "max"),
            UsingCloud=(
                "Using cloud marketplaces?",
                lambda col: "Yes"
                if (col.astype(str).str.lower() == "yes").any()
                else "No",
            ),
        )
        .sort_values("FundingUSD", ascending=False)
        .head(top_n)
    )

    grouped.insert(
        1,
        "Recent Funding Amount",
        grouped.pop("FundingUSD").apply(_human_funding),
    )
    grouped = grouped.rename(columns={"UsingCloud": "Using cloud marketplace"})

    if output_path:
        output_path = Path(output_path)
        out_ext = output_path.suffix.lower()
        if out_ext == ".csv":
            grouped.to_csv(output_path, index=False)
        elif out_ext in (".xls", ".xlsx"):
            grouped.to_excel(output_path, index=False)
        else:
            raise ValueError(f"Unsupported output format: {out_ext}")

    return grouped


# --------------------------------------------------------------------------- #
# CLI helpers
# --------------------------------------------------------------------------- #
def _resolve_relative_to_script(p: Path) -> Path:
    return p if p.is_absolute() else Path(__file__).resolve().parent / p


def _auto_find_input() -> Path:
    """Return the single CSV/XLS(X) file in the script folder, or raise."""
    here = Path(__file__).resolve().parent
    candidates = sorted(
        here.glob("*.csv") + list(here.glob("*.xls")) + list(here.glob("*.xlsx"))
    )
    if not candidates:
        raise FileNotFoundError(
            "No .csv or .xlsx file found next to tester.py. "
            "Pass the filename explicitly:  python tester.py myfile.csv"
        )
    if len(candidates) > 1:
        names = "\n  • ".join(str(p.name) for p in candidates)
        raise FileExistsError(
            "Multiple data files found next to tester.py.\n"
            "Specify which one to use, e.g.:\n"
            "   python tester.py \"Some File.csv\"\n\n"
            f"Files detected:\n  • {names}"
        )
    return candidates[0]


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Show the most-recently funded partner companies and whether "
            "they sell on cloud marketplaces."
        )
    )
    parser.add_argument(
        "input_file",
        nargs="?",
        help="CSV or Excel file to analyse "
        "(if omitted, the script looks for exactly one data file beside itself)",
    )
    parser.add_argument("-o", "--output", help="Write the result to this file")
    parser.add_argument(
        "-n",
        "--top",
        type=int,
        default=5,
        help="How many companies to list (default: 5)",
    )
    args = parser.parse_args()

    if args.input_file:
        input_path = _resolve_relative_to_script(Path(args.input_file))
    else:
        input_path = _auto_find_input()  # may raise helpful error

    output_path = (
        _resolve_relative_to_script(Path(args.output)) if args.output else None
    )

    top_df = top_funded_companies(
        input_path,
        top_n=args.top,
        output_path=output_path,
    )

    print(f"\nUsing data file: {input_path.name}")
    print(f"Top {args.top} companies by recent funding:\n")
    print(top_df.to_string(index=False))
    print("")


if __name__ == "__main__":
    main()