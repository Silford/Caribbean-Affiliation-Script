import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
import re

INPUT_FILE = ""
OUTPUT_FILE = ""
MAX_WORKERS = 10

UNIVERSITIES = {
    "University of Guyana",
    "University of the Netherlands Antilles",
    "Universidad de Puerto Rico",
    "University of Belize",
    "University of Trinidad and Tobago",
    "Caribbean Maritime University",
    "Anton de Kom University of Suriname",
    "University of Technology Jamaica",
    "Universit\u00e9 d'\u00c9tat d'Ha\u00efti",
    "Universidad Aut\u00f3noma de Santo Domingo",
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


# -----------------------------
# Extract DOI from URL
# -----------------------------
def extract_doi_from_url(url):
    if not url or pd.isna(url):
        return None

    url = str(url).strip()

    if not url:
        return None

    doi_pattern = r"(10\.\d{4,9}/[-._;()/:A-Z0-9]+)"
    match = re.search(doi_pattern, url, re.I)

    if match:
        return match.group(1)

    return None


# -----------------------------
# Try Opening URL
# -----------------------------
def check_url_access(url):
    if not url:
        return None, "NO"

    try:
        response = requests.get(url, timeout=10, allow_redirects=True)
        is_success = 200 <= response.status_code < 300
        return response.status_code, "YES" if is_success else "NO"
    except requests.RequestException:
        return None, "NO"


# -----------------------------
# Fetch from OpenAlex by DOI
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
# Extract Affiliation Info
# -----------------------------
def extract_affiliation_info(work):
    if not work:
        return None, False

    matched_unis = []
    affiliated_flag = False

    for authorship in work.get("authorships", []):
        for inst in authorship.get("institutions", []):
            inst_name = inst.get("display_name", "")
            country_code = inst.get("country_code", "")
            country_name = inst.get("country", "")
            geo = inst.get("geo", {})
            geo_country = geo.get("country", "") if isinstance(geo, dict) else ""

            if inst_name and any(u.lower() in inst_name.lower() for u in UNIVERSITIES):
                matched_unis.append(inst_name)
                affiliated_flag = True

            if (
                is_caribbean_country(country_code)
                or is_caribbean_country(country_name)
                or is_caribbean_country(geo_country)
                or is_caribbean_country(inst_name)
            ):
                affiliated_flag = True

    return (
        "; ".join(set(matched_unis)) if matched_unis else None,
        affiliated_flag
    )


# -----------------------------
# Process Each Row
# -----------------------------
def process_row(row):
    article_url = row.get("Article_URL", "")

    if pd.isna(article_url):
        article_url = ""
    else:
        article_url = str(article_url).strip()

    # Extract DOI
    doi = extract_doi_from_url(article_url)

    work = None
    if doi:
        work = fetch_openalex_by_doi(doi)

    matched_unis, affiliated = extract_affiliation_info(work)

    is_caribbean_affiliated = "TRUE" if affiliated else "FALSE"
    manual_review = "YES" if (not work or (is_caribbean_affiliated == "FALSE" and not matched_unis)) else "NO"

    return {
        "Matched_Universities": matched_unis,
        "Is_Caribbean_Affiliated": is_caribbean_affiliated,
        "Manual_Review": manual_review
    }


# -----------------------------
# Main
# -----------------------------
def main():
    df = pd.read_excel(INPUT_FILE)
    df.columns = df.columns.str.strip()
    df = df.loc[:, [
        not str(col).lower().startswith("unnamed")
        and not str(col).lower().endswith("_extracted")
        for col in df.columns
    ]]

    results = [None] * len(df)

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {
                executor.submit(process_row, row): i
                for i, (_, row) in enumerate(df.iterrows())
            }

            for future in tqdm(as_completed(futures), total=len(futures)):
                i = futures[future]
                results[i] = future.result()

    results_df = pd.DataFrame(results)

    final_df = pd.concat([df.reset_index(drop=True), results_df], axis=1)
    final_df.to_excel(OUTPUT_FILE, index=False)

    print("\nURL check + affiliation check complete.")
    print("Saved as:", OUTPUT_FILE)


if __name__ == "__main__":
    main()