import pandas as pd
import ftfy
from tqdm import tqdm


def fix_encoding(input_file, output_file):
    df = pd.read_excel(input_file)

    # Clean column names
    df.columns = df.columns.str.strip()

    # Fix encoding in selected columns
    columns_to_fix = ["Title", "Authors"]

    available_columns = [col for col in columns_to_fix if col in df.columns]

    for col in columns_to_fix:
        if col in df.columns:
            df[col] = [
                ftfy.fix_text(str(value)) if pd.notna(value) else value
                for value in tqdm(
                    df[col],
                    desc=f"Fixing {col}",
                    unit="cell",
                    dynamic_ncols=True,
                    colour="cyan",
                    leave=len(available_columns) == 1,
                )
            ]
        else:
            print(f"Column not found: {col}")

    if not available_columns:
        print("No supported text columns found to clean.")
    else:
        print(
            "Encoding fixed for: "
            + ", ".join(available_columns)
        )

    df.to_excel(output_file, index=False)

    print("Saved cleaned workbook to:", output_file)


if __name__ == "__main__":
    INPUT_FILE = ""
    OUTPUT_FILE = ""

    fix_encoding(INPUT_FILE, OUTPUT_FILE)
