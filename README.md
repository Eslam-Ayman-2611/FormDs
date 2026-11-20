# FormDs Scraper

A Python-based web scraping and data management project that collects daily SEC Form D filing information from [FormDs](https://www.formds.com/) and exports the extracted data to Excel for further processing and storage in MySQL.

## Project Overview

The project consists of two independent phases:

- **Phase One – Data Scraping:** Collects Form D filing information for the current date and exports the results to an Excel file.
- **Phase Two – Database Upload:** Reads an Excel file and uploads its records into a MySQL database.

Each phase runs independently and can be executed separately.

## Key Features

- Automated collection of daily Form D filing records.
- Pagination handling to retrieve all records for the target date.
- Extraction of company names, fields, reported funding, incremental cash, filing dates, filing types, and links.
- Excel export for collected data.
- Automated MySQL database and table creation.
- Database schema validation.
- Full table refresh on every database upload.

## Technologies

- Python 3.10+
- Requests
- BeautifulSoup4
- lxml
- Pandas
- OpenPyXL
- MySQL Connector/Python
- MySQL

## Project Structure

```text
FormDs/
├── phase_one/
│   └── FormDs_Scrape.py
├── phase_two/
│   └── Formds_Upload.py
├── requirements.txt
├── .gitignore
└── README.md
```

## Installation

1. Clone the repository:

```bash
git clone https://github.com/Eslam-Ayman-2611/FormDs.git
cd FormDs
```

2. Create and activate a virtual environment:

```bash
python3 -m venv venv
source venv/bin/activate
```

3. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Phase One – Data Scraping

Run the scraper independently:

```bash
python phase_one/FormDs_Scrape.py
```

The script collects filing records matching the current date and exports them to an Excel file in the current working directory.

**Output format:**

```text
formds_YYYY-MM-DD.xlsx
```

## Phase Two – MySQL Database Upload

Before running Phase Two, configure the following variables in `phase_two/Formds_Upload.py`:

```python
EXCEL_FILE = "formds_YYYY-MM-DD.xlsx"
HOST = "localhost"
USER = "your_mysql_username"
PASSWORD = "your_mysql_password"
DB_NAME = "your_database_name"
TABLE_NAME = "your_table_name"
```

Make sure the MySQL server is running and the configured user has the necessary database privileges.

Run the uploader independently:

```bash
python phase_two/Formds_Upload.py
```

## Database Schema

The uploader creates the following table structure:

| Column              | Data Type                        |
| ------------------- | -------------------------------- |
| Order_num           | INT, AUTO_INCREMENT, PRIMARY KEY |
| Company             | VARCHAR(500)                     |
| Reported_funds      | VARCHAR(100)                     |
| Incremental_cash    | VARCHAR(100)                     |
| Filing_date         | DATE                             |
| New_or_amended      | VARCHAR(20)                      |
| Company_address     | VARCHAR(500)                     |
| Company_all_contact | VARCHAR(2000)                    |

The company name and field are combined into the `Company` column. Company address and contact information are currently stored as NULL.

## Full Refresh Behavior

**Important:** Phase Two performs a full table refresh every time it runs.

- If the table exists with the expected schema, its existing records are removed using `TRUNCATE TABLE`.
- If the existing table has a different schema, it is dropped and recreated.
- The records from the selected Excel file are then inserted into the table.

This process replaces the existing table data rather than appending new records.

**Warning:** Back up important database records before running Phase Two.

## Notes

- Phase One and Phase Two are separate scripts; running one does not automatically execute the other.
- Phase Two requires a valid Excel file and MySQL connection settings.
- Excel output files and environment-specific files are excluded from Git tracking.

## Author

**Eslam Ayman**

GitHub: [Eslam-Ayman-2611](https://github.com/Eslam-Ayman-2611)
