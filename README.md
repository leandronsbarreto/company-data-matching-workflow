# Sample Company Matching Workflow

This repository demonstrates a data-cleaning and company-matching workflow that I developed during my work in a Marketing & Business Development team in Cologne, Germany.

The original business case involved cleaning and enriching CRM company data before a migration to a new CRM system. The full dataset contained more than 600,000 company records. Many records had outdated or incomplete information, such as company names, addresses, phone numbers, VAT identification numbers and turnover data. Another important goal was to identify possible duplicate company records before the migration.

All Excel files in this repository are fictitious sample files generated for demonstration purposes. They do not contain real company data, employer data, customer data or confidential information.

## Project idea

The workflow connects three sources:

1. a base file exported from an internal CRM system;
2. a search export downloaded from an external company-information database;
3. an extracted result file with detailed company information.

The main challenge is to build a reliable bridge between the original CRM data and the external database identifier, called `Datenbank-ID` in this sample project. The script then uses this bridge to enrich the base file, move not-found records to a separate sheet, detect duplicate groups and document all actions in a log sheet.

## Why this workflow was useful

The original process would have required a large amount of manual checking, copying and pasting. I automated the extraction and matching steps with Python.

For the extraction script, I used Selenium to work with a browser session and read the result table from the website. To do this reliably, I inspected and mapped the page DOM and used JavaScript snippets through Selenium to extract the table contents.

The `extract_batch.py` script is included as a technical demonstration. It is not required to run this repository's sample workflow, because the sample extracted file is already provided.

The main script for testing the workflow is:

```bash
python apply.py
```

## Repository structure

The expected repository structure is:

```text
company-matching-workflow/
│
├── apply.py
├── extract_batch.py
├── excel_to_sqlite.py
├── README.md
│
└── sample_data/
    ├── Match_Data.xlsx
    ├── Search_Export_001.xlsx
    └── Extracted_Batch.xlsx
```

The folder name must be exactly:

```text
sample_data
```

The scripts expect the sample Excel files to be located inside this folder.

## Sample dataset

The sample dataset contains 5,000 fictitious company records.

In the sample workflow:

```text
5,000 records are present in Search_Export_001.xlsx
4,560 records are matched and remain in Source_Data after processing
440 records are moved to Not_Found
4,551 unique records are present in Extracted_Batch.xlsx
6 duplicate groups are created in Source_Data
```

The extracted batch file contains fewer rows than the number of matched companies because duplicate CRM records can point to the same external company entity. In such cases, the external result list may contain only one detailed company record, while the CRM base file still contains several internal records.

## Important terminology

The same tax identification value appears under different column names in the three sample files:

```text
Search_Export_001.xlsx:
C = Nationale ID

Match_Data.xlsx / Source_Data:
D = Nationale ID

Extracted_Batch.xlsx:
O = UID-Nummer
```

In this workflow, these fields represent the same type of information: the German VAT identification number, known in German as:

```text
Umsatzsteuer-Identifikationsnummer
```

In the sample files, the values may contain formatting noise, such as spaces:

```text
DE 101 000 001
```

The script removes whitespace before comparing these values, so the following values are treated as equivalent:

```text
DE 101 000 001
DE101000001
```

This normalization is important because the same identifier can appear with slightly different formatting in different exports.

## Sample file 1: Match_Data.xlsx

`Match_Data.xlsx` is the base file. It simulates an export from an internal CRM system.

It contains the sheet:

```text
Source_Data
```

The original columns are:

```text
A = Interne ID
B = Unternehmensname
C = Ort
D = Nationale ID
```

The `Interne ID` is essential. The external company-information database does not use this ID to search for companies, but the ID is kept in the search export. This makes it possible to connect the original CRM row back to the external database result later.

The sample file deliberately contains data-quality issues, such as extra spaces and special characters in company names, as well as spaces in tax identification numbers. This shows how the script normalizes values for matching without changing the meaning of the data.

The file also contains the sheet:

```text
Not_Found
```

This sheet receives companies that could not be matched or safely enriched.

The script also creates or updates:

```text
Apply_Log
```

This sheet documents what happened to each processed row.

## Sample file 2: Search_Export_001.xlsx

`Search_Export_001.xlsx` simulates the file downloaded after submitting the company records to an external company-information database.

It contains both matched and not-matched companies.

In this sample:

```text
5,000 companies were submitted
4,560 companies were matched
440 companies were not found
```

Important columns in this file are:

```text
A = Interne ID
B = Unternehmensname
C = Nationale ID
D = Ort
G = Gematchte Datenbank-ID
H = Gematchter Name
```

The bridge between the CRM data and the external `Datenbank-ID` is created in this file.

The original CRM values, such as `Interne ID`, `Unternehmensname` and `Nationale ID`, remain available in the same row. The external database adds the matched `Datenbank-ID` to that row. The script then uses this `Datenbank-ID` to find the full details in the extracted result file.

Column H can also contain previous company names, for example:

```text
Previous name: "OLD COMPANY NAME GMBH"
```

or:

```text
Ehemaliger Name: "OLD COMPANY NAME GMBH"
```

The script extracts these previous names and writes them to column V in `Source_Data`.

## Sample file 3: Extracted_Batch.xlsx

`Extracted_Batch.xlsx` simulates the detailed result file extracted from the external company-information database.

It contains only matched companies. Not-found companies are not present in this file.

The external database can automatically exclude duplicate results from the detailed result list. This is why the extracted result file may contain fewer rows than the number of matched companies in the search export.

This behavior is important: duplicates may still exist in the CRM data even if the detailed result list contains only one row for the external database record. The script therefore detects duplicate groups in the base file after enrichment.

Important columns in this file include:

```text
A = Seite
B = SeqNr
C = Unternehmensname
D = Datenbank-ID
E = Handelsregisternummer
F = Straße und Hausnummer
G = Postleitzahl
H = Ort
I = Telefon
J = E-Mail-Adresse
K = Handelsregisterstatus
L = VVC-Status
M = Umsatzjahr
N = Umsatz
O = UID-Nummer
P = NACE Rev. 2 - Haupttätigkeit - Code
Q = NACE Rev. 2 - Haupttätigkeit - Beschreibung
R = Bank-Postleitzahl
S = Name der Bank
```

The script copies columns C to S from this file into columns E to U of `Source_Data`.

## Matching logic

The script does not rely on only one criterion.

First, it identifies the original row by comparing:

```text
Search_Export_001.xlsx:
A = Interne ID
B = Unternehmensname
C = Nationale ID
```

against:

```text
Match_Data.xlsx / Source_Data:
A = Interne ID
B = Unternehmensname
D = Nationale ID
```

The comparison allows controlled normalization.

For `Interne ID`, the script removes only leading and trailing quotation marks and surrounding spaces.

For company names, the script normalizes the values for matching by:

```text
removing leading and trailing quotation marks
ignoring upper/lower case
removing whitespace
removing special characters
keeping letters and numbers
```

This means that the following values can be treated as equivalent during matching:

```text
COMPANY RGTZP23 MBH
COMPANY   RGTZ-P23 MBH
Company RgtZp23 mbH
```

For `Nationale ID`, the script removes all whitespace before comparison. Therefore, values such as `DE 101 000 001` and `DE101000001` are treated as the same German VAT identification number.

In `Match_Data.xlsx`, the script also physically removes whitespace from column D, `Nationale ID`, in `Source_Data` before building the matching index. In the search export, the values are cleaned in memory during comparison, without modifying the export file itself.

Second, after the correct row has been identified, the script checks whether:

```text
Search_Export_001.xlsx column G
=
Extracted_Batch.xlsx column D
```

This confirms that the `Datenbank-ID` from the search export exists in the extracted result file.

Third, the script checks whether:

```text
Extracted_Batch.xlsx column O
=
Match_Data.xlsx / Source_Data column D
```

This confirms that the `UID-Nummer` in the extracted result matches the original `Nationale ID` from the CRM row.

Because `Nationale ID` and `UID-Nummer` represent the same German VAT identification number in this workflow, this step is a key safety check.

Only if these checks pass does the script copy the detailed company data into the base file.

## Not-found handling

Rows are moved from `Source_Data` to `Not_Found` only when the current search export shows that the company was not found or when a safety check fails.

A row is moved to `Not_Found` if:

1. the matched `Datenbank-ID` is empty;
2. the `Datenbank-ID` is not present in `Extracted_Batch.xlsx`;
3. the `Datenbank-ID` appears more than once in `Extracted_Batch.xlsx`;
4. the `UID-Nummer` in the extracted result does not match the original `Nationale ID`.

Rows that are not part of the current search export are not moved. This makes the workflow cumulative and safe for block-based processing.

## Duplicate detection

Duplicate groups are written to column W:

```text
Duplikatgruppen-ID
```

A duplicate group is defined as:

```text
same Datenbank-ID
same UID-Nummer
different Interne IDs
```

In `Source_Data`, this means:

```text
A = Interne ID
F = Datenbank-ID
Q = UID-Nummer
```

The company name, city, original `Nationale ID` and register number are not used to define duplicate groups.

This is intentional. The goal is to identify CRM records that point to the same external company entity but still have different internal IDs.

The original `Nationale ID` in column D is still used earlier in the workflow as part of the matching and safety checks. Before data is copied, the script verifies that the `UID-Nummer` from `Extracted_Batch.xlsx` matches the original `Nationale ID` from `Match_Data.xlsx`.

The script recalculates duplicate groups cumulatively across the full `Source_Data` sheet. This is necessary because a duplicate from one batch can appear only after another batch has been processed.

The script does not delete duplicate rows. It only marks them with group IDs such as:

```text
DUP_000001
DUP_000002
```

## Previous names

The search export can contain previous company names in column H.

When the script finds a previous name, it writes it to column V in `Source_Data`:

```text
Ehemaliger Name
```

The previous name is not used for matching. It is only additional information.

## Turnover handling

The extracted result file stores turnover values in thousands of euros.

Therefore, the script multiplies turnover values by:

```text
1000
```

Example:

```text
320
```

becomes:

```text
320,000.00 €
```

The script applies this conversion only to rows updated in the current run. It does not recalculate old rows from previous runs.

## Apply_Log

The `Apply_Log` sheet records the result of each processed search-export row.

It includes:

```text
Datum
Uhrzeit
Status
Quellblatt
Quellzeile
Interne ID
Unternehmensname
Nationale ID
Datenbank-ID
Grund
```

Typical levels are:

```text
OK
INFO
WARNING
ERROR
```

This provides detailed process logging and helps identify rows that need manual review.

## Running the Excel workflow

Install the required Python packages:

```bash
pip install openpyxl
```

Then run:

```bash
python apply.py
```

The script will:

```text
create a backup of Match_Data.xlsx
read Search_Export_001.xlsx
read Extracted_Batch.xlsx
match the search export against Source_Data
copy detailed company information into Source_Data
move not-found records to Not_Found
write previous company names to column V
mark duplicate groups in column W
create or update Apply_Log
save the updated Match_Data.xlsx
```

## Optional: creating a SQLite database

The repository also includes a simple SQLite step for demonstrating how the processed Excel data can be queried with SQL.

The script is:

```bash
excel_to_sqlite.py
```

It reads the sample Excel files and creates a local SQLite database:

```text
company_matching_demo.db
```

Install the required packages:

```bash
pip install pandas openpyxl
```

Then run:

```bash
python excel_to_sqlite.py
```

The script creates the following SQLite tables:

```text
source_data
not_found
search_export
extracted_batch
```

## Optional: simple SQL examples

After creating `company_matching_demo.db`, open it with DB Browser for SQLite or another SQLite client.

### Count records in Source_Data

```sql
SELECT COUNT(*) AS total_source_data
FROM source_data;
```

Expected result after running `apply.py` and then `excel_to_sqlite.py`:

```text
4560
```

### Confirm the total number of records

```sql
SELECT
    (SELECT COUNT(*) FROM source_data) AS total_source_data,
    (SELECT COUNT(*) FROM not_found) AS total_not_found,
    (SELECT COUNT(*) FROM source_data) + (SELECT COUNT(*) FROM not_found) AS total_records;
```

Expected result:

```text
total_source_data | total_not_found | total_records
4560              | 440             | 5000
```

### Inspect the first rows

```sql
SELECT *
FROM source_data
LIMIT 10;
```

### Inspect the central matching columns

```sql
SELECT
    interne_id,
    unternehmensname,
    ort,
    nationale_id,
    datenbank_id
FROM source_data
LIMIT 20;
```

### Check duplicate groups

```sql
SELECT
    duplikatgruppen_id,
    COUNT(*) AS anzahl
FROM source_data
WHERE duplikatgruppen_id IS NOT NULL
  AND duplikatgruppen_id <> ''
GROUP BY duplikatgruppen_id
ORDER BY duplikatgruppen_id;
```

Expected result: six duplicate groups.

### Inspect companies inside duplicate groups

```sql
SELECT
    duplikatgruppen_id,
    interne_id,
    unternehmensname,
    datenbank_id
FROM source_data
WHERE duplikatgruppen_id IS NOT NULL
  AND duplikatgruppen_id <> ''
ORDER BY duplikatgruppen_id, interne_id;
```

### Inspect previous company names

```sql
SELECT
    interne_id,
    unternehmensname,
    ehemaliger_name
FROM source_data
WHERE ehemaliger_name IS NOT NULL
  AND ehemaliger_name <> '';
```

### Inspect turnover values

```sql
SELECT
    interne_id,
    unternehmensname,
    umsatzjahr,
    umsatz
FROM source_data
WHERE umsatz IS NOT NULL
  AND umsatz <> ''
LIMIT 20;
```

## What the project demonstrates

This project demonstrates:

* Excel automation with `openpyxl`;
* browser-based extraction with Selenium;
* DOM inspection and JavaScript-based table extraction;
* safe multi-step matching logic;
* cumulative processing in batches;
* data enrichment;
* not-found handling;
* duplicate detection;
* detailed process logging;
* backup creation before modifying the base file;
* optional conversion of Excel data into SQLite;
* basic SQL queries for inspecting processed data.

## Data privacy

All sample files are fictitious and were generated for demonstration purposes. They do not contain real CRM data, real customer data, employer data or confidential company information.

The workflow is based on a real data-cleaning use case, but all names, IDs, VAT numbers, addresses and financial values in the sample files are artificial.