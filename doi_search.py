import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
import os

# Caribbean Country List
CARIBBEAN_COUNTRIES = {
    "Jamaica", "Trinidad and Tobago", "Barbados", "Bahamas",
    "Antigua and Barbuda", "Saint Lucia", "Grenada", "Dominica",
    "Saint Vincent and the Grenadines", "Guyana", "Suriname",
    "Haiti", "Dominican Republic", "Belize",
    "Montserrat", "Saint Kitts and Nevis",
    "St. Lucia", "St. Vincent and the Grenadines",
    "St. Kitts and Nevis", "Cuba"
}

# Caribbean Universities
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

# OpenAlex Lookup
def fetch_openalex_by_doi(doi):
    try:
        url = f"https://api.openalex.org/works/https://doi.org/{doi}"
        r = requests.get(url, timeout=10)

        if r.status_code != 200:
            return None

        return r.json()
    except:
        return None

# Crossref Lookup (Fallback)
def fetch_crossref_by_doi(doi):
    try:
        url = f"https://api.crossref.org/works/{doi}"
        r = requests.get(url, timeout=10)

        if r.status_code != 200:
            return None

        return r.json().get("message")
    except:
        return None

# Extract + Classify (OpenAlex)
def extract_openalex(work):
    resolved_title = work.get("display_name", "") or ""
    url = work.get("canonical_url", "") or ""

    authors_list = []
    universities_list = []
    countries_list = []
    is_caribbean = False

    for author in work.get("authorships", [])[:10]:
        authors_list.append(author["author"]["display_name"])

        for inst in author.get("institutions", []):
            inst_name = inst.get("display_name", "") or ""
            country = inst.get("country", "") or ""

            if inst_name:
                universities_list.append(inst_name)
                for uni in UNIVERSITIES:
                    if uni.lower() in inst_name.lower():
                        is_caribbean = True

            if country:
                countries_list.append(country)
                if country in CARIBBEAN_COUNTRIES:
                    is_caribbean = True

    return (
        resolved_title,
        " | ".join(sorted(set(authors_list))),
        " | ".join(sorted(set(universities_list))),
        " | ".join(sorted(set(countries_list))),
        "Yes" if is_caribbean else "No",
        url
    )

# Extract + Classify (Crossref)
def extract_crossref(work, doi=""):
    titles = work.get("title", [])
    resolved_title = titles[0] if titles else ""
    url = f"https://doi.org/{doi}" if doi else ""

    authors_list = []
    universities_list = []
    countries_list = []
    is_caribbean = False

    for author in work.get("author", [])[:10]:
        name = f"{author.get('given','')} {author.get('family','')}".strip()
        if name:
            authors_list.append(name)

        for aff in author.get("affiliation", []):
            inst_name = aff.get("name", "")
            if inst_name:
                universities_list.append(inst_name)

                for uni in UNIVERSITIES:
                    if uni.lower() in inst_name.lower():
                        is_caribbean = True

                for c in CARIBBEAN_COUNTRIES:
                    if c.lower() in inst_name.lower():
                        countries_list.append(c)
                        is_caribbean = True

    return (
        resolved_title,
        " | ".join(sorted(set(authors_list))),
        " | ".join(sorted(set(universities_list))),
        " | ".join(sorted(set(countries_list))),
        "Yes" if is_caribbean else "No",
        url
    )

# Process each row in the Excel file
def process_row(row):
    doi = row.get("DOI", "")
    
    # Convert to string to handle numeric DOIs from Excel
    if pd.isna(doi):
        doi = ""
    else:
        doi = str(doi).strip()

    work = None

    if not doi:
        return "", "", "", "", "Needs Manual Verification", ""

    doi = doi.rstrip(".,);")

    # Try OpenAlex
    work = fetch_openalex_by_doi(doi)
    if work:
        resolved_title, authors, universities, countries, caribbean, url = extract_openalex(work)
        return resolved_title, authors, universities, countries, caribbean, url

    # Fallback to Crossref
    work = fetch_crossref_by_doi(doi)
    if work:
        resolved_title, authors, universities, countries, caribbean, url = extract_crossref(work, doi)
        return resolved_title, authors, universities, countries, caribbean, url

    # Manual
    return "", "", "", "", "Needs Manual Verification", ""

# RUN
def main():
    INPUT_FILE = "" # Put name of file here with extension. E.g: input.xlsx

    df = pd.read_excel(INPUT_FILE)

    df = df.dropna(how="all")
    df = df[df["DOI"].astype(str).str.strip() != ""]
    df = df.reset_index(drop=True)

    n = len(df)

    resolved_titles = [""] * n
    authors_col = [""] * n
    universities_col = [""] * n
    countries_col = [""] * n
    caribbean_col = ["Needs Manual Verification"] * n
    url_col = [""] * n

    manual_review = []

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {executor.submit(process_row, df.iloc[i]): i for i in range(n)}

        for future in tqdm(as_completed(futures), total=n):
            i = futures[future]
            resolved_title, authors, universities, countries, caribbean, url = future.result()

            resolved_titles[i] = resolved_title
            authors_col[i] = authors
            universities_col[i] = universities
            countries_col[i] = countries
            caribbean_col[i] = caribbean
            url_col[i] = url

            if caribbean == "Needs Manual Verification":
                manual_review.append({
                    "Row_Number": i + 1,
                    "DOI": df.iloc[i].get("DOI", ""),
                    "Title": df.iloc[i].get("Title", ""),
                    "URL": url,
                    "Reason": "Not found in OpenAlex or Crossref"
                })

    df["Resolved_Title"] = resolved_titles
    df["Authors"] = authors_col
    df["Universities"] = universities_col
    df["Countries"] = countries_col
    df["Caribbean"] = caribbean_col
    df["URL"] = url_col

    manual_df = pd.DataFrame(manual_review)

    output_file = os.path.splitext(INPUT_FILE)[0] + "Autumn_results.xlsx"

    with pd.ExcelWriter(output_file) as writer:
        df.to_excel(writer, sheet_name="Results", index=False)
        manual_df.to_excel(writer, sheet_name="Manual Review", index=False)

    print(f"\nSaved to: {output_file}")
    print("Done.")

if __name__ == "__main__":
    main()