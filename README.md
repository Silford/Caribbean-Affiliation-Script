


# 📄 **Caribbean Affiliation Classifier (DOI-Based)**

## **Overview**

This script automates the classification of academic publications as Caribbean or Non-Caribbean based on author institutional affiliations. It is **ONLY** for use by individuals working on the SoSCAI project as of ***18/FEB/2026*** until its conclusion.

**It uses:**

* DOI-based lookup
* Automated extraction of author affiliations
* Caribbean country and university matching
* Parallel processing for scalability 

The script processes an Excel file containing DOIs and outputs: 
* Resolved article metadata
* Author names
* University affiliations
* Countries
* Caribbean classification 
* Manual review sheet for unresolved cases

🔍 Classification Rule
A publication is classed as Caribbean = Yes **IF**:
* At least one author is affiliated with a recognized Caribbean university  
                                         
     **OR**
                                             
* an institution located in a Caribbean country 

If none of the authors meet this condition, the publication is classified as: **Caribbean = No**

If metadata cannot be retried from OpenAlex nor CrossRef, the publication is flagged : **Needs Manual Verification**



## 🛠️**How it works**

1. Reads input Excel file (***input.xlsx***)
2. Removes empty rows
3. Skips rows without DOI
4. Queries OpenAlex & CrossRef (fallback) using direct DOI endpoint: 

    https://api.openalex.org/works/https://doi.org/{DOI}
    https://api.crossref.org/works/{doi}

5. Extracts: 
	* *Title*
	* *Authors*
	* *Universities*
	* *Countries*
6. Applies Caribbean classification logic
7. Exports results to: 
          ***input_results.xlsx***
 

## 📂 Input Format
Your Excel file must contain: 
|Column Name|Required  |
|-|--|
|DOI| ✅ Yes |

Example: 
|DOI|
|-|
|10.1002/cjas.1296|

Rows without DOI are skipped.

## 📤 Output Files
The script generates:

1️⃣ **Results Sheet**

 - DOI
 - Resolved_Title
 - Authors
 - Universities
 - Countries
 
 2️⃣**Manual Review Sheet**
 
 - Row number
 - DOI
 - Title
 - Reason

## 🚀 **Installation**

 1. Clone Repository
  `git clone https://github.com/Silford/Caribbean-Affiliation-Script.git
`

2. **Install Dependencies**
`pip install pandas requests tqdm openpyxl
`

▶️**Running the Script**
Place your Excel file in the project directory

Edit: 

    INPUT_FILE = "input.xlsx"

Then run: 

    python auto.py
   
   Output will be saved as:
   

    input_results.xlsx

## ⚡Performance
* Uses ThreadPoolExecutor for parallel API requests
* Suitable for processing 1000-5000 DOIs
* Typical speed: ~20-50 records per second (depending on network speed)
## ⚠️Limitations

* Only processes rows with DOI
* If OpenAlex or CrossRef does not index a DOI, it is marked for manual review
* Some regional journals may not be indexed in OpenAlex or CrossRef
