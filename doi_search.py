import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
import os

# Caribbean Universities ONLY (no country flag anymore)
UNIVERSITIES = {
    "University of Guyana",
    "University of the Netherlands Antilles",
    "Universidad de Puerto Rico",
    "University of Belize",
    "University of Trinidad and Tobago",
    "Caribbean Maritime University",
    "Anton de Kom University of Suriname",
    "University of Technology Jamaica",
    "Université d'État d'Haïti",
    "Universidad Autónoma de Santo Domingo",
    "University of the Bahamas",
    "University of the West Indies",
    "Universidad de la Habana",
    "University of Havana",
    "State University of Haiti",
    "University of Suriname",
    "Autonomous University of Santo Domingo"
}

COUNTRIES = {
    
}


def is_caribbean_institution(inst_name):
    return any(u.lower() in inst_name.lower() for u in UNIVERSITIES)


def is_caribbean_country(value):
    if not value:
        return False

    value_l = str(value).strip().lower()
    if not value_l:
        return False

    for country in COUNTRIES:
        country_l = str(country).strip().lower()
        if country_l and (value_l == country_l or country_l in value_l):
            return True

    return False


def unique_pipe_join(values):
    return " | ".join(sorted(set(values)))


def resolve_doi_column(df):
    cleaned_to_original = {
        str(col).strip().lower(): col
        for col in df.columns
    }

    for candidate in ["doi", "doi id", "doi_id", "document doi"]:
        if candidate in cleaned_to_original:
            return cleaned_to_original[candidate]

    for cleaned_name, original_name in cleaned_to_original.items():
        if "doi" in cleaned_name:
            return original_name

    raise KeyError(
        "No DOI-like column found. Available columns: "
        + ", ".join(map(str, df.columns))
    )


def resolve_optional_column(df, candidates):
    cleaned_to_original = {
        str(col).strip().lower(): col
        for col in df.columns
    }

    for candidate in candidates:
        key = str(candidate).strip().lower()
        if key in cleaned_to_original:
            return cleaned_to_original[key]

    return None


# -----------------------------
# OpenAlex Lookup
# -----------------------------
def fetch_openalex_by_doi(doi):
    try:
        url = f"https://api.openalex.org/works/https://doi.org/{doi}"
        r = requests.get(url, timeout=10)

        if r.status_code == 200:
            return r.json()
    except requests.RequestException:
        return None

    return None


# -----------------------------
# Crossref Lookup (Fallback)
# -----------------------------
def fetch_crossref_by_doi(doi):
    try:
        url = f"https://api.crossref.org/works/{doi}"
        r = requests.get(url, timeout=10)

        if r.status_code == 200:
            return r.json().get("message")
    except requests.RequestException:
        return None

    return None


# -----------------------------
# Extract from OpenAlex
# -----------------------------
def extract_openalex(work, doi=""):
    universities_list = []
    affiliated_flag = False

    for author in work.get("authorships", [])[:10]:
        for inst in author.get("institutions", []):
            inst_name = inst.get("display_name", "")
            country_code = inst.get("country_code", "")
            country_name = inst.get("country", "")
            country_geo = inst.get("geo", {}).get("country", "") if isinstance(inst.get("geo", {}), dict) else ""

            if inst_name:
                universities_list.append(inst_name)

            if (
                is_caribbean_institution(inst_name)
                or is_caribbean_country(country_name)
                or is_caribbean_country(country_geo)
                or is_caribbean_country(country_code)
                or is_caribbean_country(inst_name)
            ):
                affiliated_flag = True

    return (
        unique_pipe_join(universities_list),
        "TRUE" if affiliated_flag else "FALSE"
    )


# -----------------------------
# Extract from Crossref
# -----------------------------
def extract_crossref(work, doi=""):
    universities_list = []
    affiliated_flag = False

    for author in work.get("author", [])[:10]:
        author_country = author.get("country", "")
        if is_caribbean_country(author_country):
            affiliated_flag = True

        for aff in author.get("affiliation", []):
            inst_name = aff.get("name", "")
            if inst_name:
                universities_list.append(inst_name)

            if is_caribbean_institution(inst_name) or is_caribbean_country(inst_name):
                affiliated_flag = True

    return (
        unique_pipe_join(universities_list),
        "TRUE" if affiliated_flag else "FALSE"
    )


# -----------------------------
# Process Row
# -----------------------------
def process_row(doi):

    if pd.isna(doi):
        return "", "Needs Manual Verification"

    doi = str(doi).strip().rstrip(".,);")

    if not doi:
        return "", "Needs Manual Verification"

    # Try OpenAlex first
    work = fetch_openalex_by_doi(doi)
    if work:
        return extract_openalex(work, doi)

    # Fallback Crossref
    work = fetch_crossref_by_doi(doi)
    if work:
        return extract_crossref(work, doi)

    return "", "Needs Manual Verification"


# -----------------------------
# MAIN
# -----------------------------
def main():
    INPUT_FILE = "test_set.xlsx"

    df = pd.read_excel(INPUT_FILE)
    df = df.dropna(how="all")

    # Standardize headers to avoid issues with leading/trailing spaces.
    df.columns = [str(col).strip() for col in df.columns]
    df = df.loc[:, [
        not str(col).lower().startswith("unnamed")
        and not str(col).lower().endswith("_extracted")
        for col in df.columns
    ]]
    doi_column = resolve_doi_column(df)

    df = df[df[doi_column].astype(str).str.strip() != ""]
    df = df.reset_index(drop=True)

    source_universities_col = resolve_optional_column(
        df,
        ["universities", "university", "affiliation", "affiliations", "institution", "institutions"]
    )

    n = len(df)

    universities_col = [""] * n
    caribbean_col = ["Needs Manual Verification"] * n

    manual_review_indices = []

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {executor.submit(process_row, df.iloc[i].get(doi_column, "")): i for i in range(n)}

        for future in tqdm(as_completed(futures), total=n):
            i = futures[future]

            universities, caribbean = future.result()

            universities_value = universities
            if caribbean == "FALSE" and not str(universities_value).strip() and source_universities_col is not None:
                universities_value = str(df.iloc[i].get(source_universities_col, "")).strip()

            if caribbean == "FALSE" and not str(universities_value).strip():
                caribbean = "Manual Review"

            universities_col[i] = universities_value
            caribbean_col[i] = caribbean

            if caribbean in {"Needs Manual Verification", "Manual Review"}:
                manual_review_indices.append(i)

    df["Universities"] = universities_col
    df["Is_Caribbean_Affiliated"] = caribbean_col

    manual_df = df.iloc[manual_review_indices].copy()

    output_file = os.path.splitext(INPUT_FILE)[0] + "_results.xlsx"

    with pd.ExcelWriter(output_file) as writer:
        df.to_excel(writer, sheet_name="Results", index=False)
        manual_df.to_excel(writer, sheet_name="Manual Review", index=False)

    print(f"\nSaved to: {output_file}")
    print("Done.")


if __name__ == "__main__":
    main()