import pandas as pd
import ftfy

INPUT_FILE = "" # Insert Name here
OUTPUT_FILE = "" # Insert Name here

# Read file
df = pd.read_excel(INPUT_FILE)

# Clean column names (removes hidden spaces)
df.columns = df.columns.str.strip()

# Fix encoding in Title and Authors columns
columns_to_fix = ["Title", "Authors"]

for col in columns_to_fix:
    if col in df.columns:
        df[col] = df[col].apply(
            lambda x: ftfy.fix_text(str(x)) if pd.notna(x) else x
        )
    else:
        print(f"Column not found: {col}")

# Save corrected file
df.to_excel(OUTPUT_FILE, index=False)

print("Encoding fixed for Title and Authors.")
print("Saved as:", OUTPUT_FILE)