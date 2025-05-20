import os
import re
import argparse
import numpy as np
import pandas as pd

_SUFFIX = {"K": 1e3, "M": 1e6, "B": 1e9}

def parse_funding(raw) -> float | np.nan:
    """
    Convert a funding string like '$2.5M' or '750K' to a numeric USD amount.
    Returns NaN for anything that cannot be parsed.
    """
    if pd.isna(raw):
        return np.nan

    s = (str(raw)
         .replace(",", "")            # strip thousands separators
         .replace("$", "")
         .replace("€", "")
         .replace("£", "")
         .strip())

    m = re.fullmatch(r"([\d.]+)\s*([KMB])?", s, flags=re.I)
    if not m:
        return np.nan

    number = float(m.group(1))
    multiplier = _SUFFIX.get(m.group(2).upper() if m.group(2) else "", 1)
    return number * multiplier

def human_funding(usd: float | np.nan) -> str:
    """Inverse of parse_funding – format a float back to $xxM / $x.xB, etc."""
    if pd.isna(usd):
        return ""
    if usd >= 1e9:
        return f"${usd/1e9:.1f}B"
    if usd >= 1e6:
        return f"${usd/1e6:.0f}M" if usd % 1e6 == 0 else f"${usd/1e6:.1f}M"
    if usd >= 1e3:
        return f"${usd/1e3:.0f}K"
    return f"${usd:,.0f}"

# --------------------------------------------------------------------------- #
# Core routine
# --------------------------------------------------------------------------- #
def top_funded_companies(
    path: str,
    top_n: int = 5,
    output_path: str | None = None,
) -> pd.DataFrame:
    """
    Return the `top_n` companies by RECENT funding amount.

    Parameters
    ----------
    path : str
        CSV or Excel file containing, at minimum,
        'Company', 'Recent Funding Amount', 'Using cloud marketplaces?'.
    top_n : int, default 5
    output_path : str, optional
        If given, write the result to this location (csv / xlsx matched by ext).

    Returns
    -------
    pd.DataFrame
        Columns: Company | Recent Funding Amount | Using cloud marketplace
    """
    # 1 | Load
    ext = os.path.splitext(path.lower())[1]
    if ext == ".csv":
        df = pd.read_csv(path)
    elif ext in (".xls", ".xlsx"):
        df = pd.read_excel(path)
    else:
        raise ValueError(f"Unsupported input format: {ext}")

    # 2 | Check required columns (ignore capitalisation / stray spaces)
    rename = {c: c.strip() for c in df.columns}
    df = df.rename(columns=rename)       # normalise headings

    required = {
        "Company",
        "Recent Funding Amount",
        "Using cloud marketplaces?",
    }
    if not required.issubset(df.columns):
        missing = required - set(df.columns)
        raise KeyError(f"Missing column(s): {', '.join(missing)}")

    # 3 | Clean & normalise
    df["Company"] = df["Company"].astype(str).str.strip().str.title()
    df["FundingUSD"] = df["Recent Funding Amount"].apply(parse_funding)
    df = df.dropna(subset=["Company", "FundingUSD"])

    # 4 | Collapse duplicates – max single round + any-yes marketplace flag
    agg = (
        df.groupby("Company", as_index=False)
          .agg(
              FundingUSD=("FundingUSD", "max"),
              UsingCloud=(
                  "Using cloud marketplaces?",
                  lambda col: "Yes" if (col.astype(str).str.lower() == "yes").any() else "No",
              ),
          )
          .sort_values("FundingUSD", ascending=False)
          .head(top_n)
    )

    # 5 | Prettify for output
    agg.insert(1, "Recent Funding Amount", agg.pop("FundingUSD").apply(human_funding))
    agg = agg.rename(columns={"UsingCloud": "Using cloud marketplace"})

    # 6 | Optional write-out
    if output_path:
        out_ext = os.path.splitext(output_path.lower())[1]
        if out_ext == ".csv":
            agg.to_csv(output_path, index=False)
        elif out_ext in (".xls", ".xlsx"):
            agg.to_excel(output_path, index=False)
        else:
            raise ValueError(f"Unsupported output format: {out_ext}")

    return agg

# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Show the top-funded partner companies "
                    "and whether they sell on cloud marketplaces."
    )
    parser.add_argument("input_file", help="CSV or Excel file to analyse")
    parser.add_argument("-o", "--output", help="Write the result to this file")
    parser.add_argument(
        "-n", "--top", type=int, default=5,
        help="How many companies to return (default 5)"
    )
    args = parser.parse_args()

    top_df = top_funded_companies(args.input_file, args.top, args.output)
    print(f"Top {args.top} companies by recent funding:\n")
    print(top_df.to_string(index=False))