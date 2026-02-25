import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm


INPUT_FILE = "fixed_titles.xlsx"
OUTPUT_FILE = "articles_caribbean_checked.xlsx"
MAX_WORKERS = 5

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

# OpenAlex Search Functions

def search_by_title(title):
    try:
        url = "https://api.openalex.org/works"
        params = {"search": title, "per-page": 1}
        r = requests.get(url, params=params, timeout=10)
        if r.status_code == 200:
            results = r.json().get("results", [])
            if results:
                return results[0]
    except:
        return None
    return None


# Affiliation Extraction
def extract_affiliation_info(work):
    if not work:
        return None, None, False

    institutions = []
    countries = []
    caribbean_flag = False

    for authorship in work.get("authorships", []):
        for inst in authorship.get("institutions", []):

            name = inst.get("display_name")
            country_name = inst.get("country")

            if name:
                institutions.append(name)

                # Flexible university matching
                if any(u.lower() in name.lower() for u in UNIVERSITIES):
                    caribbean_flag = True

            if country_name:
                countries.append(country_name)

                if country_name in CARIBBEAN_COUNTRIES:
                    caribbean_flag = True

    return (
        "; ".join(set(institutions)) if institutions else None,
        "; ".join(set(countries)) if countries else None,
        caribbean_flag
    )


# Processing each row in the Excel file 
def process_row(row):
    title = row.get("Title", "")

    work = None

    if not work and title:
        work = search_by_title(title)

    institutions, countries, caribbean = extract_affiliation_info(work)

    return {
        "Institutions": institutions,
        "Countries": countries,
        "Caribbean_Affiliated": caribbean,
        "Manual_Review": "YES" if not work else "NO"
    }


# Main function
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

    print("\nCaribbean affiliation check complete.")
    print("Saved as:", OUTPUT_FILE)


if __name__ == "__main__":
    main()