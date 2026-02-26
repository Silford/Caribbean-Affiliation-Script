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


# -----------------------------
# Try Opening URL
# -----------------------------
def check_url_access(url):
    if not url:
        return None, "No"

    try:
        r = requests.get(url, timeout=10, allow_redirects=True)
        return r.status_code, "Yes" if r.status_code == 200 else "No"
    except:
        return None, "No"


# -----------------------------
# OpenAlex Lookup
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
# Crossref Lookup (Fallback)
# -----------------------------
def fetch_crossref_by_doi(doi):
    try:
        url = f"https://api.crossref.org/works/{doi}"
        r = requests.get(url, timeout=10)

        if r.status_code == 200:
            return r.json().get("message")
    except:
        return None

    return None


# -----------------------------
# Extract from OpenAlex
# -----------------------------
def extract_openalex(work):
    resolved_title = work.get("display_name", "") or ""

    # Get landing page URL
    url = None
    primary_location = work.get("primary_location", {})
    if primary_location:
        url = primary_location.get("landing_page_url")

    authors_list = []
    universities_list = []
    affiliated_flag = False

    for author in work.get("authorships", [])[:10]:
        name = author.get("author", {}).get("display_name")
        if name:
            authors_list.append(name)

        for inst in author.get("institutions", []):
            inst_name = inst.get("display_name", "")
            if inst_name:
                universities_list.append(inst_name)

                if any(u.lower() in inst_name.lower() for u in UNIVERSITIES):
                    affiliated_flag = True

    status_code, reachable = check_url_access(url)

    return (
        resolved_title,
        " | ".join(sorted(set(authors_list))),
        " | ".join(sorted(set(universities_list))),
        "Yes" if affiliated_flag else "No",
        url,
        status_code,
        reachable
    )


# -----------------------------
# Extract from Crossref
# -----------------------------
def extract_crossref(work, doi=""):
    titles = work.get("title", [])
    resolved_title = titles[0] if titles else ""

    url = f"https://doi.org/{doi}" if doi else ""

    authors_list = []
    universities_list = []
    affiliated_flag = False

    for author in work.get("author", [])[:10]:
        name = f"{author.get('given','')} {author.get('family','')}".strip()
        if name:
            authors_list.append(name)

        for aff in author.get("affiliation", []):
            inst_name = aff.get("name", "")
            if inst_name:
                universities_list.append(inst_name)

                if any(u.lower() in inst_name.lower() for u in UNIVERSITIES):
                    affiliated_flag = True

    status_code, reachable = check_url_access(url)

    return (
        resolved_title,
        " | ".join(sorted(set(authors_list))),
        " | ".join(sorted(set(universities_list))),
        "Yes" if affiliated_flag else "No",
        url,
        status_code,
        reachable
    )


# -----------------------------
# Process Row
# -----------------------------
def process_row(row):
    doi = row.get("DOI", "")

    if pd.isna(doi):
        return "", "", "", "Needs Manual Verification", "", None, "No"

    doi = str(doi).strip().rstrip(".,);")

    if not doi:
        return "", "", "", "Needs Manual Verification", "", None, "No"

    # Try OpenAlex first
    work = fetch_openalex_by_doi(doi)
    if work:
        return extract_openalex(work)

    # Fallback Crossref
    work = fetch_crossref_by_doi(doi)
    if work:
        return extract_crossref(work, doi)

    return "", "", "", "Needs Manual Verification", "", None, "No"


# -----------------------------
# MAIN
# -----------------------------
def main():
    INPUT_FILE = "DOIs_Only.xlsx"

    df = pd.read_excel(INPUT_FILE)
    df = df.dropna(how="all")
    df = df[df["DOI"].astype(str).str.strip() != ""]
    df = df.reset_index(drop=True)

    n = len(df)

    resolved_titles = [""] * n
    authors_col = [""] * n
    universities_col = [""] * n
    caribbean_col = ["Needs Manual Verification"] * n
    url_col = [""] * n
    status_col = [None] * n
    reachable_col = ["No"] * n

    manual_review = []

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {executor.submit(process_row, df.iloc[i]): i for i in range(n)}

        for future in tqdm(as_completed(futures), total=n):
            i = futures[future]

            resolved_title, authors, universities, caribbean, url, status, reachable = future.result()

            resolved_titles[i] = resolved_title
            authors_col[i] = authors
            universities_col[i] = universities
            caribbean_col[i] = caribbean
            url_col[i] = url
            status_col[i] = status
            reachable_col[i] = reachable

            if caribbean == "Needs Manual Verification":
                manual_review.append({
                    "Row_Number": i + 1,
                    "DOI": df.iloc[i].get("DOI", ""),
                    "Title": df.iloc[i].get("Title", ""),
                    "Reason": "Not found in OpenAlex or Crossref"
                })

    df["Resolved_Title"] = resolved_titles
    df["Authors"] = authors_col
    df["Universities"] = universities_col
    df["Caribbean_Affiliated"] = caribbean_col
    df["URL"] = url_col
    df["URL_Status_Code"] = status_col
    df["URL_Reachable"] = reachable_col

    manual_df = pd.DataFrame(manual_review)

    output_file = os.path.splitext(INPUT_FILE)[0] + "_results.xlsx"

    with pd.ExcelWriter(output_file) as writer:
        df.to_excel(writer, sheet_name="Results", index=False)
        manual_df.to_excel(writer, sheet_name="Manual Review", index=False)

    print(f"\nSaved to: {output_file}")
    print("Done.")


if __name__ == "__main__":
    main()