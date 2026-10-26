
from pathlib import Path
import pandas as pd
import re
import math

# ============================================================
# 1. FILE PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = BASE_DIR / "input" / "assets_table1.xlsx"
OUTPUT_FILE = BASE_DIR / "input" / "assets_table1_cleaned.xlsx"

EXPECTED_COLUMNS = [
    "folder_path",
    "file_name",
    "format",
    "duration_second",
    "file_size_mb",
    "aspect_ratio"
]


# ============================================================
# 2. NORMALIZE COLUMN NAMES
# ============================================================

def normalize_header(value):
    if pd.isna(value):
        return ""

    value = str(value).strip().lower()
    value = re.sub(r"\s+", "_", value)

    return value


# ============================================================
# 3. READ EXCEL AND FIND ACTUAL HEADER
# ============================================================

# Read all rows as data to avoid Unnamed columns
raw_df = pd.read_excel(
    INPUT_FILE,
    header=None,
    dtype=object
)

header_row = None
header_mapping = None

for row_index, row in raw_df.iterrows():

    values = [
        normalize_header(value)
        for value in row.tolist()
    ]

    # Find the actual column names, regardless of position
    if all(col in values for col in EXPECTED_COLUMNS):

        header_row = row_index

        header_mapping = {
            col: values.index(col)
            for col in EXPECTED_COLUMNS
        }

        break

if header_row is None:
    raise ValueError(
        "Cannot find the expected column headers. "
        "Please check the Excel file."
    )

print(f"Header detected at Excel row: {header_row + 1}")

# Extract only rows below the correct header
df = raw_df.iloc[header_row + 1:].copy()

# Select the correct columns by their positions
df = df.iloc[
    :,
    [header_mapping[col] for col in EXPECTED_COLUMNS]
]

# Rename columns explicitly
df.columns = EXPECTED_COLUMNS

df = df.reset_index(drop=True)

print(f"Rows before cleaning: {len(df)}")


# ============================================================
# 4. REMOVE WHITESPACE
# ============================================================

def clean_whitespace(value):

    if pd.isna(value):
        return None

    if isinstance(value, str):

        # Remove spaces at beginning and end
        value = value.strip()

        # Replace multiple whitespace characters
        value = re.sub(r"\s+", " ", value)

        if not value:
            return None

    return value


df = df.map(clean_whitespace)


# ============================================================
# 5. REMOVE REPEATED HEADER ROWS
# ============================================================

def is_header_row(row):

    values = [
        normalize_header(value)
        for value in row.tolist()
    ]

    return all(
        values[i] == EXPECTED_COLUMNS[i]
        for i in range(len(EXPECTED_COLUMNS))
    )


df = df[
    ~df.apply(is_header_row, axis=1)
].copy()


# ============================================================
# 6. CONVERT FILE SIZE TO MB
# ============================================================

def convert_file_size(value):

    if pd.isna(value):
        return None

    # If already numeric, assume MB
    if isinstance(value, (int, float)):
        return float(value)

    value = str(value).strip().upper()

    # Remove whitespace and commas
    value = re.sub(r"\s+", "", value)
    value = value.replace(",", "")

    # Match numbers with optional KB or MB
    match = re.fullmatch(
        r"([0-9]+(?:\.[0-9]+)?)(KB|MB)?",
        value
    )

    if not match:
        print(f"WARNING: Invalid file size: {value}")
        return None

    number = float(match.group(1))
    unit = match.group(2)

    if unit == "KB":
        return number / 1000

    # MB or no unit
    return number


df["file_size_mb"] = df["file_size_mb"].apply(
    convert_file_size
)

df["file_size_mb"] = pd.to_numeric(
    df["file_size_mb"],
    errors="coerce"
)


# ============================================================
# 7. CLEAN DURATION COLUMN
# ============================================================

df["duration_second"] = pd.to_numeric(
    df["duration_second"],
    errors="coerce"
)


# ============================================================
# 8. PRESERVE ASPECT RATIO AS TEXT
# ============================================================

def clean_aspect_ratio(value):

    if pd.isna(value):
        return None

    value = str(value).strip()

    # Remove whitespace around colon
    value = re.sub(r"\s*:\s*", ":", value)

    return value


df["aspect_ratio"] = df["aspect_ratio"].apply(
    clean_aspect_ratio
)


# ============================================================
# 9. REMOVE EMPTY ROWS
# ============================================================

df = df.replace(r"^\s*$", pd.NA, regex=True)

# Remove rows where every column is empty
df = df.dropna(how="all")

df = df.reset_index(drop=True)


# ============================================================
# 10. SAVE CLEANED EXCEL
# ============================================================

OUTPUT_FILE.parent.mkdir(
    parents=True,
    exist_ok=True
)

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    df.to_excel(
        writer,
        sheet_name="Assets",
        index=False
    )

    worksheet = writer.sheets["Assets"]

    # Format file size as numeric values
    for cell in worksheet["E"][1:]:
        cell.number_format = "0.###"

    # Preserve aspect ratio as text
    for cell in worksheet["F"][1:]:
        cell.number_format = "@"

    # Adjust column widths
    column_widths = {
        "A": 60,
        "B": 45,
        "C": 15,
        "D": 20,
        "E": 18,
        "F": 18
    }

    for col, width in column_widths.items():
        worksheet.column_dimensions[col].width = width


# ============================================================
# 11. FINAL SUMMARY
# ============================================================

print("\n" + "=" * 60)
print("EXCEL CLEANING COMPLETED")
print("=" * 60)

print(f"Input file: {INPUT_FILE}")
print(f"Output file: {OUTPUT_FILE}")

print(f"Final rows: {len(df)}")
print(f"Columns: {df.columns.tolist()}")

print("\nFile size examples:")
print(df["file_size_mb"].head(10).tolist())

print("\nPreview:")
print(df.head())

print("\nSUCCESS: Cleaned Excel file saved.")
