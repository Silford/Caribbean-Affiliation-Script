import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
import os


UNIVERSITIES = [
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
    "Autonomous University of Santo Domingo",
]

COUNTRIES = [
    "Cuba",
    "Dominican Republic",
    "Puerto Rico",
    "Antigua and Barbuda",
    "Antigua & Barbuda",
    "Anguilla",
    "British Virgin Islands",
    "Bahamas",
    "Barbados",
    "Belize",
    "Cayman Islands",
    "Dominica",
    "Grenada",
    "Guyana",
    "Guadeloupe",
    "Jamaica",
    "Montserrat",
    "Saint Kitts and Nevis",
    "St. Kitts and Nevis",
    "St. Kitts & Nevis",
    "Saint Lucia",
    "St. Lucia",
    "Saint Vincent and the Grenadines",
    "St. Vincent and the Grenadines",
    "St. Vincent & the Grenadines",
    "Trinidad and Tobago",
    "Trinidad & Tobago",
    "Turks and Caicos Islands",
    "U.S. Virgin Islands",
    "Haiti",
    "Martinique",
    "Aruba",
    "Curaçao",
    "Saint-Martin",
    "Saint-Barthélemy",
    "Sint Maarten",
    "St. Maarten",
    "Bonaire",
    "Saba",
    "Sint Eustatius",
    "St. Eustatius",
]


def is_caribbean_institution(inst_name):
    return any(u.lower() in str(inst_name).lower() for u in UNIVERSITIES)


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


def get_crossref_author_name(author):
    full_name = str(author.get("name", "")).strip()

    if full_name:
        return full_name

    given = str(author.get("given", "")).strip()
    family = str(author.get("family", "")).strip()

    return " ".join(part for part in [given, family] if part).strip()


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


def fetch_openalex_by_doi(doi):
    try:
        url = f"https://api.openalex.org/works/https://doi.org/{doi}"
        response = requests.get(url, timeout=10)

        if response.status_code == 200:
            return response.json()

    except requests.RequestException:
        return None

    return None


def fetch_crossref_by_doi(doi):
    try:
        url = f"https://api.crossref.org/works/{doi}"
        response = requests.get(url, timeout=10)

        if response.status_code == 200:
            return response.json().get("message")

    except requests.RequestException:
        return None

    return None


def extract_openalex(work, doi=""):
    authors_list = []
    universities_list = []
    affiliated_flag = False

    for author in work.get("authorships", [])[:10]:
        author_name = author.get("author", {}).get("display_name", "")

        if author_name:
            authors_list.append(author_name)

        for inst in author.get("institutions", []):
            inst_name = inst.get("display_name", "")
            country_code = inst.get("country_code", "")
            country_name = inst.get("country", "")
            geo = inst.get("geo", {})
            country_geo = geo.get("country", "") if isinstance(geo, dict) else ""

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
        unique_pipe_join(authors_list),
        unique_pipe_join(universities_list),
        "TRUE" if affiliated_flag else "FALSE",
    )


def extract_crossref(work, doi=""):
    authors_list = []
    universities_list = []
    affiliated_flag = False

    for author in work.get("author", [])[:10]:
        author_name = get_crossref_author_name(author)

        if author_name:
            authors_list.append(author_name)

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
        unique_pipe_join(authors_list),
        unique_pipe_join(universities_list),
        "TRUE" if affiliated_flag else "FALSE",
    )


def process_row(doi):
    if pd.isna(doi):
        return "", "", "Needs Manual Verification"

    doi = str(doi).strip().rstrip(".,);")

    if not doi:
        return "", "", "Needs Manual Verification"

    work = fetch_openalex_by_doi(doi)

    if work:
        return extract_openalex(work, doi)

    work = fetch_crossref_by_doi(doi)

    if work:
        return extract_crossref(work, doi)

    return "", "", "Needs Manual Verification"


def process_doi_file(input_file, output_file=None, max_workers=10):
    df = pd.read_excel(input_file)
    df = df.dropna(how="all")

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
        [
            "universities",
            "university",
            "affiliation",
            "affiliations",
            "institution",
            "institutions",
        ],
    )

    row_count = len(df)

    authors_col = [""] * row_count
    universities_col = [""] * row_count
    caribbean_col = ["Needs Manual Verification"] * row_count

    manual_review_indices = []

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(process_row, df.iloc[i].get(doi_column, "")): i
            for i in range(row_count)
        }

        for future in tqdm(
            as_completed(futures),
            total=row_count,
            desc="Processing DOI rows",
            unit="row",
            dynamic_ncols=True,
            colour="cyan",
        ):
            i = futures[future]

            authors, universities, caribbean = future.result()

            universities_value = universities

            if (
                caribbean == "FALSE"
                and not str(universities_value).strip()
                and source_universities_col is not None
            ):
                universities_value = str(df.iloc[i].get(source_universities_col, "")).strip()

            if caribbean == "FALSE" and not str(universities_value).strip():
                caribbean = "Manual Review"

            authors_col[i] = authors
            universities_col[i] = universities_value
            caribbean_col[i] = caribbean

            if caribbean in {"Needs Manual Verification", "Manual Review"}:
                manual_review_indices.append(i)

    df["Authors_Extracted"] = authors_col
    df["Universities"] = universities_col
    df["Is_Caribbean_Affiliated"] = caribbean_col

    manual_df = df.iloc[manual_review_indices].copy()

    if output_file is None:
        output_file = os.path.splitext(input_file)[0] + "_results.xlsx"

    with pd.ExcelWriter(output_file) as writer:
        df.to_excel(writer, sheet_name="Results", index=False)
        manual_df.to_excel(writer, sheet_name="Manual Review", index=False)

    print(f"\nSaved DOI results to: {output_file}")

    return output_file


if __name__ == "__main__":
    INPUT_FILE = ""
    OUTPUT_FILE = None

    process_doi_file(INPUT_FILE, OUTPUT_FILE)
