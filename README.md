# 🌟 Caribbean Affiliation Script 🌟
The Caribbean Affiliation Script is a comprehensive data analyzing and processing script focused on extracting author information to determine whether they are affiliated with a Caribbean institution/organisation and/or is a Caribbean-based author. The script's core features include URL searching and processing, DOI searching and processing, and text encoding fixing in Excel files.

## 🚀 Key Features
* URL searching and processing using concurrent threads
* DOI searching and processing using concurrent threads
* Encoding issue fixing in Excel files
* Data manipulation and storage using pandas
* Concurrent processing using concurrent.futures
* Progress tracking using tqdm
* Text encoding fixes using ftfy

## 🛠️ Tech Stack
* Python 3.10
* requests library for making HTTP requests
* pandas library for data manipulation and storage
* concurrent.futures library for concurrent processing
* tqdm library for progress tracking
* ftfy library for text encoding fixes
* html library for text processing and encoding
* Excel files for input and output data

## 📦 Getting Started / Setup Instructions
### Prerequisites
* Python 3.10 installed on your system
* requests, pandas, concurrent.futures, tqdm, and ftfy libraries installed
* Excel files for input and output data

### Installation
1. Clone the repository using `git clone https://github.com/Silford/Caribbean-Affiliation-Script.git`
2. Install the required libraries using `pip install -r requirements.txt`
3. Set up your input and output Excel files

### Running Locally
1. Run the `url_search.py` file using `python url_search.py`
2. Run the `doi_search.py` file using `python doi_search.py`
3. Run the `fix_encoding_issues.py` file using `python fix_encoding_issues.py`

    ***NB**: It is recommended to run `fix_encoding_issues.py` first to ensure that the other scripts do not falsing flag a source as TRUE (Caribbean associated) or FALSE (not Caribbean associated)*

## 📂 Project Structure
```markdown
Caribbean-Affiliation-Script/
│
├── url_search.py
├── doiSearch.py
├── fixEncodingIssues.py
├── requirements.txt
├── input.xlsx
├── output.xlsx
└── README.md
```

## 🤝 Contributing
Contributions are welcome! If you'd like to contribute to the project, please fork the repository and submit a pull request.

## 📝 License
The Caribbean Affiliation Script is licensed under the MIT License.

## 📬 Contact Me
For any questions or concerns, please contact us at [silfordmoore@gmail.com](mailto:silfordmoore@gmail.com)

## 💖 Thanks Message
A huge thank you to everyone who has contributed to the project! 🙏