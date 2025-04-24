import os
import pandas as pd

def process_partners(input_path: str, output_path: str = None) -> pd.DataFrame:
    """
    Reads a CSV or Excel file of partner companies, de-duplicates entries,
    and returns a DataFrame sorted by capital (descending).
    
    Parameters:
        input_path (str): Path to input .csv or .xlsx file.
        output_path (str, optional): If provided, writes the result to this file
                                     (.csv or .xlsx inferred by extension).
    
    Returns:
        pd.DataFrame: Cleaned, unique companies with capitals sorted descending.
    """
    # 1. Load data
    _, ext = os.path.splitext(input_path.lower())
    if ext == '.csv':
        df = pd.read_csv(input_path)
    elif ext in ('.xls', '.xlsx'):
        df = pd.read_excel(input_path)
    else:
        raise ValueError(f"Unsupported input format: {ext}")

    required = {'Company', 'Capital'}
    if not required.issubset(df.columns):
        missing = required - set(df.columns)
        raise KeyError(f"Missing required column(s): {', '.join(missing)}")

    #normalize output
    df['Company'] = df['Company'].astype(str).str.strip().str.title()
    df['Capital'] = pd.to_numeric(df['Capital'], errors='coerce')
    df = df.dropna(subset=['Company', 'Capital'])

    #delete repetitions
    df_unique = (
        df
        .groupby('Company', as_index=False)
        .agg({'Capital': 'max'})
    )

    df_sorted = df_unique.sort_values(by='Capital', ascending=False)

    if output_path:
        _, out_ext = os.path.splitext(output_path.lower())
        if out_ext == '.csv':
            df_sorted.to_csv(output_path, index=False)
        elif out_ext in ('.xls', '.xlsx'):
            df_sorted.to_excel(output_path, index=False)
        else:
            raise ValueError(f"Unsupported output format: {out_ext}")

    return df_sorted

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Generate a ranked list of unique partner companies by capital."
    )
    parser.add_argument("input_file", help="Path to input CSV or Excel file")
    parser.add_argument(
        "-o", "--output",
        help="(Optional) Path to write the output ranking (CSV or Excel)."
    )
    args = parser.parse_args()

    result_df = process_partners(args.input_file, args.output)
    print("Top 10 companies by capital:")
    print(result_df.head(10).to_string(index=False))
