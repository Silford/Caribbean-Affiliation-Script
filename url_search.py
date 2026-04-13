import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
import re
import unicodedata
import time
import json
import html

INPUT_FILE = "7000-8155_batch_manualreview.xlsx"
OUTPUT_FILE = "7000-8155_batch_manualreview_results.xlsx"
MAX_WORKERS = 30
REQUEST_TIMEOUT = 8
REQUEST_RETRIES = 2
RETRY_BACKOFF_SECONDS = 0.5
CHECK_URL_ACCESS = False
HTTP_HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; URLMetadataBot/1.0; +https://example.org/bot)"
}

SESS = requests.Session()
SESS.headers.update(HTTP_HEADERS)
SESS.mount("http://", requests.adapters.HTTPAdapter(pool_connections=50, pool_maxsize=50))
SESS.mount("https://", requests.adapters.HTTPAdapter(pool_connections=50, pool_maxsize=50))

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
    "Autonomous University of Santo Domingo"
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
    "St. Vincent and the Grenandines",
    "St. Vincent & the Grenadines",
    "Trinidad and Tobago",
    "Trinidad & Tobago",
    "Turks and Caicos Islands",
    "U.S. Virgin Islands",
    "Haiti",
    "Martinique",
    "Guadeloupe",
    "Aruba",
    "Curaçao",
    "Saint-Martin",
    "Saint-Barthélemy",
    "Sint Maarteen",
    "St. Maarteen",
    "Bonaire",
    "Saba",
    "Sint Eustatius",
    "St. Eustatius"
]

def is_caribbean_country(value):
    if not value:
        return False

    value_l = normalize_text(value)
    if not value_l:
        return False

    for country in COUNTRIES:
        country_l = normalize_text(country)
        if country_l and re.search(rf"(?<!\w){re.escape(country_l)}(?!\w)", value_l):
            return True

    return False


def normalize_text(value):
    if value is None or pd.isna(value):
        return ""

    text = unicodedata.normalize("NFKD", str(value))
    text = text.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", text).strip().lower()


def unique_pipe_join(values):
    cleaned = sorted({str(value).strip() for value in values if str(value).strip()})
    return " | ".join(cleaned)


def request_json(url, expected_json_key=None):
    last_error = None

    for attempt in range(REQUEST_RETRIES):
        try:
            response = SESS.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
            if response.status_code == 200:
                payload = response.json()
                if expected_json_key is None or expected_json_key in payload:
                    return payload
                return payload
            last_error = f"HTTP {response.status_code}"
        except requests.RequestException as exc:
            last_error = str(exc)

        if attempt < REQUEST_RETRIES - 1:
            time.sleep(RETRY_BACKOFF_SECONDS * (2 ** attempt))

    return None


def request_html(url):
    last_error = None

    for attempt in range(REQUEST_RETRIES):
        try:
            response = SESS.get(
                url,
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True
            )
            if response.status_code == 200:
                return response.text
            last_error = f"HTTP {response.status_code}"
        except requests.RequestException as exc:
            last_error = str(exc)

        if attempt < REQUEST_RETRIES - 1:
            time.sleep(RETRY_BACKOFF_SECONDS * (2 ** attempt))

    return None


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
        doi = match.group(1).strip().rstrip(".,);\"]\'")
        return doi

    return None


def extract_doi_from_text(text):
    if not text:
        return None

    doi_pattern = r"(10\.\d{4,9}/[-._;()/:A-Z0-9]+)"
    match = re.search(doi_pattern, str(text), re.I)
    if not match:
        return None

    return match.group(1).strip().rstrip(".,);\"]\'")


def parse_meta_tags(html_content):
    meta_map = {}
    meta_pattern = re.compile(r"<meta\s+[^>]*>", re.I)
    attr_pattern = re.compile(r"([\w:-]+)\s*=\s*([\"'])(.*?)\2", re.I | re.S)

    for tag in meta_pattern.findall(html_content or ""):
        attrs = {}
        for key, _, value in attr_pattern.findall(tag):
            attrs[key.lower()] = html.unescape(value).strip()

        key = attrs.get("name") or attrs.get("property") or attrs.get("itemprop")
        content = attrs.get("content", "").strip()

        if key and content:
            key_l = key.lower()
            meta_map.setdefault(key_l, []).append(content)

    return meta_map


def extract_title(html_content, meta_map):
    for key in ["citation_title", "dc.title", "og:title", "twitter:title", "title"]:
        values = meta_map.get(key, [])
        if values:
            return values[0].strip()

    title_match = re.search(r"<title[^>]*>(.*?)</title>", html_content or "", re.I | re.S)
    if title_match:
        return re.sub(r"\s+", " ", html.unescape(title_match.group(1))).strip()

    return None


def extract_jsonld_blocks(html_content):
    blocks = []
    pattern = re.compile(
        r"<script[^>]*type=[\"']application/ld\+json[\"'][^>]*>(.*?)</script>",
        re.I | re.S
    )

    for raw_block in pattern.findall(html_content or ""):
        block = raw_block.strip()
        if not block:
            continue
        try:
            parsed = json.loads(block)
            blocks.append(parsed)
        except json.JSONDecodeError:
            continue

    return blocks


def collect_jsonld_fields(node, dois, affiliations):
    if isinstance(node, dict):
        for key, value in node.items():
            key_l = str(key).lower()

            if key_l in {"doi", "identifier"} and isinstance(value, str):
                candidate_doi = extract_doi_from_text(value)
                if candidate_doi:
                    dois.append(candidate_doi)

            if key_l == "affiliation":
                if isinstance(value, str):
                    affiliations.append(value)
                elif isinstance(value, dict):
                    name = value.get("name")
                    if isinstance(name, str) and name.strip():
                        affiliations.append(name.strip())
                elif isinstance(value, list):
                    for item in value:
                        if isinstance(item, str) and item.strip():
                            affiliations.append(item.strip())
                        elif isinstance(item, dict):
                            name = item.get("name")
                            if isinstance(name, str) and name.strip():
                                affiliations.append(name.strip())

            collect_jsonld_fields(value, dois, affiliations)

    elif isinstance(node, list):
        for item in node:
            collect_jsonld_fields(item, dois, affiliations)


def fetch_webpage_metadata(url):
    html_content = request_html(url)
    if not html_content:
        return None

    meta_map = parse_meta_tags(html_content)
    title = extract_title(html_content, meta_map)

    doi_candidates = []
    affiliation_candidates = []

    for key in ["citation_doi", "dc.identifier", "dc.identifier.doi", "prism.doi", "doi"]:
        for value in meta_map.get(key, []):
            candidate_doi = extract_doi_from_text(value)
            if candidate_doi:
                doi_candidates.append(candidate_doi)

    for key, values in meta_map.items():
        if any(token in key for token in ["affiliation", "institution", "university", "school"]):
            affiliation_candidates.extend([value for value in values if value])

    for block in extract_jsonld_blocks(html_content):
        collect_jsonld_fields(block, doi_candidates, affiliation_candidates)

    if not doi_candidates:
        body_doi = extract_doi_from_text(html_content)
        if body_doi:
            doi_candidates.append(body_doi)

    doi = doi_candidates[0] if doi_candidates else None
    affiliations = sorted({aff.strip() for aff in affiliation_candidates if aff and aff.strip()})

    return {
        "title": title,
        "doi": doi,
        "affiliations": affiliations
    }


def extract_webpage_affiliation_info(metadata):
    if not metadata:
        return None, False

    matched_unis = []
    affiliated_flag = False

    for inst_name in metadata.get("affiliations", []):
        normalized_inst_name = normalize_text(inst_name)

        if normalized_inst_name and any(normalize_text(u) in normalized_inst_name for u in UNIVERSITIES):
            matched_unis.append(inst_name)
            affiliated_flag = True

        if is_caribbean_country(inst_name):
            affiliated_flag = True

    return (
        unique_pipe_join(matched_unis) if matched_unis else None,
        affiliated_flag
    )


# -----------------------------
# Try Opening URL
# -----------------------------
def check_url_access(url):
    if not url:
        return None, "NO", "NO URL"

    try:
        response = SESS.get(
            url,
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True
        )
        is_success = 200 <= response.status_code < 300
        return response.status_code, "YES" if is_success else "NO", f"HTTP {response.status_code}"
    except requests.RequestException as exc:
        return None, "NO", f"EXCEPTION: {exc}"


# -----------------------------
# Fetch from OpenAlex by DOI
# -----------------------------
def fetch_openalex_by_doi(doi):
    url = f"https://api.openalex.org/works/https://doi.org/{doi}"
    return request_json(url)


# -----------------------------
# Fetch from Crossref by DOI
# -----------------------------
def fetch_crossref_by_doi(doi):
    url = f"https://api.crossref.org/works/{doi}"
    payload = request_json(url)
    if not payload:
        return None
    return payload.get("message")


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

            normalized_inst_name = normalize_text(inst_name)
            if normalized_inst_name and any(normalize_text(u) in normalized_inst_name for u in UNIVERSITIES):
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
        unique_pipe_join(matched_unis) if matched_unis else None,
        affiliated_flag
    )


def extract_crossref_affiliation_info(work):
    if not work:
        return None, False

    matched_unis = []
    affiliated_flag = False

    for author in work.get("author", []):
        for aff in author.get("affiliation", []):
            inst_name = aff.get("name", "")
            normalized_inst_name = normalize_text(inst_name)

            if inst_name and any(normalize_text(u) in normalized_inst_name for u in UNIVERSITIES):
                matched_unis.append(inst_name)
                affiliated_flag = True

            if is_caribbean_country(inst_name):
                affiliated_flag = True

    return (
        unique_pipe_join(matched_unis) if matched_unis else None,
        affiliated_flag
    )


# -----------------------------
# Process Each Row
# -----------------------------
def process_row(row):
    article_url = row.get("ArticleURL", "") # ENSURE THIS IS THE SAME NAME AS THE ACTUAL ARTICLE URL COLUMN IN THE INPUT FILE

    if pd.isna(article_url):
        article_url = ""
    else:
        article_url = str(article_url).strip()

    url_status_code, url_accessible, url_check_detail = (None, "NOT CHECKED", "NOT CHECKED")
    if CHECK_URL_ACCESS:
        url_status_code, url_accessible, url_check_detail = check_url_access(article_url)

    # Extract DOI
    doi = extract_doi_from_url(article_url)

    work = None
    webpage_metadata = None

    if doi:
        work = fetch_openalex_by_doi(doi)
        if not work:
            work = fetch_crossref_by_doi(doi)

    if not work and article_url:
        webpage_metadata = fetch_webpage_metadata(article_url)
        metadata_doi = webpage_metadata.get("doi") if webpage_metadata else None
        if metadata_doi and metadata_doi != doi:
            doi = metadata_doi
            work = fetch_openalex_by_doi(doi)
            if not work:
                work = fetch_crossref_by_doi(doi)

    if work and "authorships" in work:
        matched_unis, affiliated = extract_affiliation_info(work)
    elif work:
        matched_unis, affiliated = extract_crossref_affiliation_info(work)
    else:
        matched_unis, affiliated = extract_webpage_affiliation_info(webpage_metadata)

    is_caribbean_affiliated = "TRUE" if affiliated else "FALSE"
    has_metadata = bool(webpage_metadata and (
        webpage_metadata.get("title")
        or webpage_metadata.get("doi")
        or webpage_metadata.get("affiliations")
    ))
    manual_review = "YES" if ((not work and not has_metadata) or (is_caribbean_affiliated == "FALSE" and not matched_unis)) else "NO"

    return {
        "URL_Status_Code": url_status_code,
        "URL_Accessible": url_accessible,
        "URL_Check_Detail": url_check_detail,
        "Matched_Universities": matched_unis,
        "Is_Caribbean_Affiliated": is_caribbean_affiliated,
        "Manual_Review": manual_review
    }


# -----------------------------
# Main
# -----------------------------
def main():
    if not INPUT_FILE:
        raise ValueError("INPUT_FILE is empty. Set INPUT_FILE before running the script.")

    if not OUTPUT_FILE:
        raise ValueError("OUTPUT_FILE is empty. Set OUTPUT_FILE before running the script.")

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
                try:
                    results[i] = future.result()
                except Exception as exc:
                    results[i] = {
                        "URL_Status_Code": None,
                        "URL_Accessible": "NO",
                        "URL_Check_Detail": f"ERROR: {exc}",
                        "Matched_Universities": None,
                        "Is_Caribbean_Affiliated": "FALSE",
                        "Manual_Review": f"ERROR: {exc}"
                    }

    results_df = pd.DataFrame(results)

    final_df = pd.concat([df.reset_index(drop=True), results_df], axis=1)
    final_df.to_excel(OUTPUT_FILE, index=False)

    print("\nURL check + affiliation check complete.")
    print("Saved as:", OUTPUT_FILE)


if __name__ == "__main__":
    main()