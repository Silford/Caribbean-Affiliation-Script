import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
import os

CARIBBEAN_COUNTRIES = {
    "Jamaica", "Trinidad and Tobago", "Barbados", "Bahamas", "Antigua and Barbuda",
    "Saint Lucia", "Grenada", "Dominica", "Saint Vincent and the Grenadines",
    "Guyana", "Suriname", "Haiti", "Dominican Republic", "Belize", "Montserrat", 
    "Saint Kitts and Nevis", "St. Lucia", "St. Vincent and the Grenadines", "St. Kitts and Nevis"
}

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
    "Universidad de la Habana"
}

def fix_encoding(text):
    if isinstance(text, str):
        try:
            return text.encode("latin1").decode("utf-8")
        except:
            return text
    return text

def fetch_from_openalex(doi=None, title=None):
    try:
        if doi:
            url = f"https://api.openalex.org/works?filter=doi:{doi}"
        elif title:
            url = f"https://api.openalex.org/works?search={title}"
        else:
            return None

        r = requests.get(url, timeout=10)
        if r.status_code != 200:
            return None

        data = r.json()
        if not data.get("results"):
            return None
        return data["results"][0]
    except:
        return None

def fetch_from_crossref(doi=None, title=None):
    try:
        if doi:
            url = f"https://api.crossref.org/works/{doi}"
        elif title:
            url = f"https://api.crossref.org/works?query.title={title}&rows=1"
        else:
            return None

        r = requests.get(url, timeout=10)
        if r.status_code != 200:
            return None

        data = r.json()

        if doi:
            return data.get("message")
        else:
            items = data.get("message", {}).get("items", [])
            if not items:
                return None
            return items[0]
    except:
        return None

def extract_openalex(work):
    resolved_title = work.get("display_name", "") or ""
    authors_list, universities_list, countries_list = [], [], []
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
        "Yes" if is_caribbean else "No"
    )

def extract_crossref(work):
    titles = work.get("title", [])
    resolved_title = titles[0] if isinstance(titles, list) and titles else ""

    authors_list, universities_list, countries_list = [], [], []
    is_caribbean = False

    for author in work.get("author", [])[:10]:
        name = f"{author.get('given', '')} {author.get('family', '')}".strip()
        if name:
            authors_list.append(name)

        for aff in author.get("affiliation", []):
            inst_name = aff.get("name", "") or ""
            if inst_name:
                universities_list.append(inst_name)

                for uni in UNIVERSITIES:
                    if uni.lower() in inst_name.lower():
                        is_caribbean = True

                # Crossref often doesn't provide a separate country field; detect in text
                for c in CARIBBEAN_COUNTRIES:
                    if c.lower() in inst_name.lower():
                        countries_list.append(c)
                        is_caribbean = True

    return (
        resolved_title,
        " | ".join(sorted(set(authors_list))),
        " | ".join(sorted(set(universities_list))),
        " | ".join(sorted(set(countries_list))),
        "Yes" if is_caribbean else "No"
    )

def fetch_metadata_for_row(row):
    doi = fix_encoding(str(row.get("DOI", "")).strip())
    title = fix_encoding(str(row.get("Title", "")).strip())

    # OpenAlex first
    work = fetch_from_openalex(doi=doi if doi else None,
                               title=title if (not doi and title) else None)
    if work:
        return extract_openalex(work)

    # Crossref fallback
    work = fetch_from_crossref(doi=doi if doi else None,
                               title=title if (not doi and title) else None)
    if work:
        return extract_crossref(work)

    return "", "", "", "", "Unknown"


# -----------------------------
# Run
# -----------------------------
INPUT_FILE = "test_input.xlsx"  # change this
df = pd.read_excel(INPUT_FILE)

# fix encoding across all text cells (helps titles)
df = df.applymap(fix_encoding)

# safe duplicate removal
existing_columns = [c for c in ["DOI", "Title"] if c in df.columns]
if existing_columns:
    df = df.drop_duplicates(subset=existing_columns)

# Pre-allocate result arrays by row count (prevents misalignment)
n = len(df)
resolved_titles = [""] * n
authors_col = [""] * n
universities_col = [""] * n
countries_col = [""] * n
caribbean_col = ["Unknown"] * n

manual_review = []

with ThreadPoolExecutor(max_workers=10) as executor:
    futures = {executor.submit(fetch_metadata_for_row, df.iloc[i]): i for i in range(n)}

    for future in tqdm(as_completed(futures), total=n):
        i = futures[future]
        resolved_title, authors, universities, countries, caribbean = future.result()

        resolved_titles[i] = resolved_title
        authors_col[i] = authors
        universities_col[i] = universities
        countries_col[i] = countries
        caribbean_col[i] = caribbean

        if caribbean == "Unknown":
            manual_review.append({"RowIndex": i, "DOI": df.iloc[i].get("DOI", ""), "Title": df.iloc[i].get("Title", "")})

# Attach results to SAME ROWS (this is what you wanted)
df["Resolved_Title"] = resolved_titles
df["Authors"] = authors_col
df["Universities"] = universities_col
df["Countries"] = countries_col
df["Caribbean"] = caribbean_col

manual_df = pd.DataFrame(manual_review)

output_file = os.path.splitext(INPUT_FILE)[0] + "_classified.xlsx"
with pd.ExcelWriter(output_file) as writer:
    df.to_excel(writer, sheet_name="Classified", index=False)
    manual_df.to_excel(writer, sheet_name="Manual_Review", index=False)

print(f"\nSaved to: {output_file}")
print("Done.")
