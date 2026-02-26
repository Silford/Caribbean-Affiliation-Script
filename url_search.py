import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
import re

INPUT_FILE = "fixed_titles.xlsx"
OUTPUT_FILE = "results.xlsx"
MAX_WORKERS = 5

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
        return response.status_code, "YES" if response.status_code == 200 else "NO"
    except:
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
    except:
        return None

    return None


# -----------------------------
# Extract Affiliation Info
# -----------------------------
def extract_affiliation_info(work):
    if not work:
        return None, None, False

    author_names = []
    matched_unis = []
    affiliated_flag = False

    for authorship in work.get("authorships", []):
        author = authorship.get("author", {})
        name = author.get("display_name")

        if name:
            author_names.append(name)

        for inst in authorship.get("institutions", []):
            inst_name = inst.get("display_name")
            if inst_name and any(u.lower() in inst_name.lower() for u in UNIVERSITIES):
                matched_unis.append(inst_name)
                affiliated_flag = True

    return (
        "; ".join(set(author_names)) if author_names else None,
        "; ".join(set(matched_unis)) if matched_unis else None,
        affiliated_flag
    )


# -----------------------------
# Process Each Row
# -----------------------------
def process_row(row):
    try: 
        article_url = row.get("Article_URL", "")

        if pd.isna(article_url):
            article_url = ""
        else:
            article_url = str(article_url).strip()
    except Exception as e:
        return {
            "URL_Status_Code": None,
            "URL_Reachable": "NO",
            "Authors": None,
            "Matched_Universities": None,
            "Caribbean_Affiliated": False,
            "Manual_Review": "YES"
        }

    # 1. Check URL itself
    status_code, reachable = check_url_access(article_url)

    # 2. Extract DOI
    doi = extract_doi_from_url(article_url)

    work = None
    if doi:
        work = fetch_openalex_by_doi(doi)

    authors, matched_unis, affiliated = extract_affiliation_info(work)

    return {
        "URL_Status_Code": status_code,
        "URL_Reachable": reachable,
        "Authors": authors,
        "Matched_Universities": matched_unis,
        "Caribbean_Affiliated": affiliated,
        "Manual_Review": "YES" if not work else "NO"
    }


# -----------------------------
# Main
# -----------------------------
def main():
    df = pd.read_excel(INPUT_FILE)
    df.columns = df.columns.str.strip()

    results = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [executor.submit(process_row, row) for _, row in df.iterrows()]

        for future in tqdm(as_completed(futures), total=len(futures)):
            results.append(future.result())

    results_df = pd.DataFrame(results)

    final_df = pd.concat([df.reset_index(drop=True), results_df], axis=1)
    final_df.to_excel(OUTPUT_FILE, index=False)

    print("\nURL check + affiliation check complete.")
    print("Saved as:", OUTPUT_FILE)


if __name__ == "__main__":
    main()