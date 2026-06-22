from pathlib import Path
import re
import time
from datetime import datetime
from copy import copy

from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
SAMPLE_DATA_DIR = BASE_DIR / "sample_data"

DEBUGGER_ADDRESS = "127.0.0.1:9222"

OUTPUT_FILE = SAMPLE_DATA_DIR / "Extracted_Batch.xlsx"

PAGE_WAIT_SECONDS = 30
POLL_SECONDS = 0.25
PAGE_CHANGE_RETRIES = 3
PAGE_RETRY_SLEEP_SECONDS = 2
PAGE_STABILITY_SECONDS = 1.0

PAGE_INPUT_ID = "ContentContainer1_ctl00_Content_ListHeader_ListNavigation_CurrentPage"
PAGES_LABEL_ID = "ContentContainer1_ctl00_Content_ListHeader_ListNavigation_PagesLabel"

LEFT_TABLE_ID = "ContentContainer1_ctl00_Content_ListCtrl1_LB1_FDTBL"
RIGHT_TABLE_ID = "ContentContainer1_ctl00_Content_ListCtrl1_LB1_VDTBL"

OUTPUT_HEADERS = [
    "Seite",
    "SeqNr",
    "Unternehmensname",
    "Datenbank-ID",
    "Handelsregisternummer",
    "Straße und Hausnummer",
    "Postleitzahl",
    "Ort",
    "Telefon",
    "E-Mail-Adresse",
    "Handelsregisterstatus",
    "VVC-Status",
    "Umsatzjahr",
    "Umsatz",
    "UID-Nummer",
    "NACE Rev. 2 - Haupttätigkeit - Code",
    "NACE Rev. 2 - Haupttätigkeit - Beschreibung",
    "Bank-Postleitzahl",
    "Name der Bank",
]


# ============================================================
# BASIC HELPERS
# ============================================================

def clean_text(value):
    if value is None:
        return ""

    text = str(value).replace("\xa0", " ")
    text = " ".join(text.split())
    return text.strip()


def value_or_not_found(value):
    value = clean_text(value)
    return value if value else "Not found"


def shorten(value, max_len=80):
    value = clean_text(value)

    if len(value) <= max_len:
        return value

    return value[:max_len - 3] + "..."


def looks_like_postal_code(value):
    return bool(re.fullmatch(r"\d{5}", clean_text(value)))


def looks_like_database_id(value):
    """
    Typical generic database ID example:
    DE7370127732

    This deliberately does not match register-number values like:
    60313 HRB 40807.
    """
    text = clean_text(value)
    return bool(re.fullmatch(r"[A-Z]{2}[A-Z0-9]{6,}", text))


# ============================================================
# CHROME CONNECTION
# ============================================================

def connect_to_open_chrome():
    options = Options()
    options.add_experimental_option("debuggerAddress", DEBUGGER_ADDRESS)
    return webdriver.Chrome(options=options)


def get_visible_text(driver):
    try:
        return clean_text(
            driver.execute_script("return document.body ? document.body.innerText : '';")
        )
    except Exception:
        return ""


def find_result_list_tab(driver):
    """
    Finds the open batch/search/list tab containing the result table.

    The function looks for:
    - result-list table elements;
    - pagination elements;
    - typical batch/list/result text.
    """

    target_handle = None
    debug_tabs = []

    for handle in driver.window_handles:
        driver.switch_to.window(handle)

        url = clean_text(driver.current_url)
        title = clean_text(driver.title)

        try:
            text = get_visible_text(driver)
        except Exception:
            text = ""

        try:
            has_left_table = driver.execute_script(
                "return Boolean(document.getElementById(arguments[0]));",
                LEFT_TABLE_ID
            )
            has_right_table = driver.execute_script(
                "return Boolean(document.getElementById(arguments[0]));",
                RIGHT_TABLE_ID
            )
            has_page_input = driver.execute_script(
                "return Boolean(document.getElementById(arguments[0]));",
                PAGE_INPUT_ID
            )
        except Exception:
            has_left_table = False
            has_right_table = False
            has_page_input = False

        haystack = f"{url} {title} {text}".casefold()

        debug_tabs.append(
            {
                "title": title,
                "url": url,
                "has_left_table": has_left_table,
                "has_right_table": has_right_table,
                "has_page_input": has_page_input,
            }
        )

        is_batch_or_list_url = (
            "search.batchsearchlaunch" in haystack
            or "batchsearch" in haystack
            or "list.serv" in haystack
            or "search.quicksearch.serv" in haystack
        )

        has_batch_or_list_text = (
            "saved company identifiers" in haystack
            or "saved unternehmen identifiers" in haystack
            or "total" in haystack
            or "gesamt" in haystack
            or "company name" in haystack
            or "name des unternehmens" in haystack
        )

        has_result_tables = has_left_table and has_right_table

        if (
            has_result_tables
            or (is_batch_or_list_url and has_batch_or_list_text)
            or (is_batch_or_list_url and has_page_input)
        ):
            target_handle = handle
            break

    if target_handle is None:
        print("")
        print("DEBUG: Open Chrome tabs seen by Selenium:")

        for index, tab in enumerate(debug_tabs, start=1):
            print(f"Tab {index}:")
            print(f"  Title: {tab['title']}")
            print(f"  URL: {tab['url']}")
            print(f"  Has left result table: {tab['has_left_table']}")
            print(f"  Has right result table: {tab['has_right_table']}")
            print(f"  Has page input: {tab['has_page_input']}")
            print("")

        raise RuntimeError(
            "No batch/search/list tab found. "
            "Open the result list with the company rows visible."
        )

    driver.switch_to.window(target_handle)
    return target_handle


# ============================================================
# PAGE STATUS AND PAGINATION
# ============================================================

def get_page_status(driver):
    driver.switch_to.default_content()

    status = driver.execute_script(
        """
        const pageInput = document.getElementById(arguments[0]);
        const pagesLabel = document.getElementById(arguments[1]);

        const current = pageInput ? pageInput.value : "1";
        const label = pagesLabel ? pagesLabel.textContent : "";

        return {current: current, label: label};
        """,
        PAGE_INPUT_ID,
        PAGES_LABEL_ID
    )

    current_value = clean_text(status.get("current", "1"))
    label_text = clean_text(status.get("label", ""))

    if re.fullmatch(r"\d+", current_value):
        current_page = int(current_value)
    else:
        current_page = 1

    match = re.search(r"(?:von|of)\s+(\d+)", label_text, flags=re.IGNORECASE)

    if match:
        total_pages = int(match.group(1))
    else:
        total_pages = 1

    return current_page, total_pages


def go_to_page(driver, target_page):
    """
    Important:
    Do not use Ctrl+A.
    Do not fire keyup/keypress/change/blur.

    Some result-list pages can jump back if several events trigger several
    AJAX postbacks. This function sets the value and calls ajaxPanelPostBack
    exactly once.
    """

    driver.switch_to.default_content()

    driver.execute_script(
        """
        const input = document.getElementById(arguments[0]);
        const value = String(arguments[1]);

        if (!input) {
            throw new Error("CurrentPage input not found");
        }

        input.focus();
        input.value = value;

        if (typeof ajaxPanelPostBack === 'function') {
            ajaxPanelPostBack(
                true,
                'GroupAjaxForSearchAndList',
                'ContentContainer1$ctl00$Content$ListHeader$ListNavigation$CurrentPage',
                ''
            );
        } else {
            throw new Error("ajaxPanelPostBack function not found");
        }
        """,
        PAGE_INPUT_ID,
        target_page
    )


def get_page_signature(driver):
    try:
        snapshot = get_batch_snapshot(driver)
        rows = snapshot.get("rows", [])
        parts = []

        for row in rows[:8]:
            result = row["result"]

            parts.append(
                "|".join(
                    [
                        str(row.get("seqnr", "")),
                        row.get("company_name", ""),
                        result.get("Datenbank-ID", ""),
                        result.get("Handelsregisternummer", ""),
                        result.get("Postleitzahl", ""),
                        result.get("Ort", ""),
                        result.get("UID-Nummer", ""),
                    ]
                )
            )

        return "\n".join(parts)

    except Exception:
        return get_visible_text(driver)[:1000]


def wait_for_page_change(driver, old_signature, expected_page):
    end_time = time.time() + PAGE_WAIT_SECONDS

    while time.time() < end_time:
        time.sleep(POLL_SECONDS)

        try:
            current_page, _ = get_page_status(driver)
            new_signature = get_page_signature(driver)

            if current_page != expected_page:
                continue

            if not new_signature or new_signature == old_signature:
                continue

            time.sleep(PAGE_STABILITY_SECONDS)

            stable_page, _ = get_page_status(driver)
            stable_signature = get_page_signature(driver)

            if (
                stable_page == expected_page
                and stable_signature
                and stable_signature == new_signature
                and stable_signature != old_signature
            ):
                return True

        except Exception:
            pass

    return False


def go_to_next_page(driver):
    current_page, total_pages = get_page_status(driver)

    if current_page >= total_pages:
        return False

    target_page = current_page + 1

    for attempt in range(1, PAGE_CHANGE_RETRIES + 1):
        print(
            f"Moving to page {target_page}/{total_pages} | "
            f"attempt {attempt}/{PAGE_CHANGE_RETRIES}"
        )

        old_signature = get_page_signature(driver)

        try:
            go_to_page(driver, target_page)
        except Exception as error:
            print(f"  Page change command failed: {error}")
            time.sleep(PAGE_RETRY_SLEEP_SECONDS)
            continue

        if wait_for_page_change(driver, old_signature, target_page):
            return True

        current_after, _ = get_page_status(driver)

        print(
            f"  Page {target_page} not confirmed. "
            f"Current page is still/again {current_after}. Retrying..."
        )

        time.sleep(PAGE_RETRY_SLEEP_SECONDS)

    return False


# ============================================================
# RESULT TABLE SNAPSHOT
# ============================================================

def wait_for_batch_table(driver):
    end_time = time.time() + PAGE_WAIT_SECONDS

    while time.time() < end_time:
        try:
            exists = driver.execute_script(
                """
                return Boolean(
                    document.getElementById(arguments[0]) &&
                    document.getElementById(arguments[1])
                );
                """,
                LEFT_TABLE_ID,
                RIGHT_TABLE_ID
            )

            if exists:
                return True

        except Exception:
            pass

        time.sleep(POLL_SECONDS)

    return False


def get_batch_snapshot(driver):
    """
    Reads the current result page in one JavaScript snapshot.

    The result page may place some fields in the fixed left table and others
    in the right table. Therefore this function:

    - reads company names and possible postal codes from the left table;
    - reads the main data cells from the right table;
    - detects whether a generic database ID is present as the first right-side value;
    - detects whether the postal code is present in the right-side row;
    - maps all later fields without blindly shifting columns.
    """

    driver.switch_to.default_content()

    if not wait_for_batch_table(driver):
        raise RuntimeError("Batch result table not found.")

    raw = driver.execute_script(
        """
        const leftTable = document.getElementById(arguments[0]);
        const rightTable = document.getElementById(arguments[1]);

        function clean(s) {
            if (s === null || s === undefined) return "";
            return String(s).replace(/\\u00a0/g, " ").replace(/\\s+/g, " ").trim();
        }

        function seqFromOnclick(onclick) {
            const m = String(onclick || "").match(/SeqNr=(\\d+)/);
            return m ? parseInt(m[1], 10) : null;
        }

        function looksLikeGermanPostalCode(s) {
            return /^\\d{5}$/.test(clean(s));
        }

        const linkItems = Array.from(
            leftTable.querySelectorAll("a[onclick*='LB1_listContentClicked'][onclick*='SeqNr=']")
        ).map(a => {
            const tr = a.closest("tr");

            let postalCode = "";

            if (tr) {
                const leftCells = Array.from(
                    tr.querySelectorAll("td.resultsItems")
                ).map(td => clean(td.textContent)).filter(Boolean);

                const postalCandidates = leftCells.filter(looksLikeGermanPostalCode);

                if (postalCandidates.length > 0) {
                    postalCode = postalCandidates[0];
                }
            }

            return {
                seqnr: seqFromOnclick(a.getAttribute("onclick")),
                company_name: clean(a.textContent),
                postal_code_from_left_table: postalCode
            };
        }).filter(x => x.seqnr !== null);

        linkItems.sort((a, b) => a.seqnr - b.seqnr);

        const rightRows = Array.from(
            rightTable.querySelectorAll("tr")
        ).filter(tr => {
            return tr.querySelectorAll("td.resultsItems").length > 0;
        }).map(tr => {
            const cells = Array.from(
                tr.querySelectorAll("td.resultsItems")
            ).map(td => clean(td.textContent));

            return cells;
        });

        return {
            names: linkItems,
            rightRows: rightRows
        };
        """,
        LEFT_TABLE_ID,
        RIGHT_TABLE_ID
    )

    names = raw.get("names", [])
    right_rows = raw.get("rightRows", [])

    max_len = min(len(names), len(right_rows))
    rows = []

    for index in range(max_len):
        name_item = names[index]
        values = [value_or_not_found(v) for v in right_rows[index]]

        postal_code_from_left_table = value_or_not_found(
            name_item.get("postal_code_from_left_table", "")
        )

        database_id_number = "Not found"

        if values and looks_like_database_id(values[0]):
            database_id_number = values[0]
            values = values[1:]

        # ------------------------------------------------------------
        # After removing the generic database ID, possible right-table orders:
        #
        # CASE A: Postal code is present in right table
        # 0  Handelsregisternummer
        # 1  Straße und Hausnummer
        # 2  Postleitzahl
        # 3  Ort
        # 4  Telefon
        # 5  E-Mail-Adresse
        # 6  Handelsregisterstatus
        # 7  VVC-Status
        # 8  Umsatz
        # 9  Umsatzjahr
        # 10 UID-Nummer
        # 11 NACE Code
        # 12 NACE Beschreibung
        # 13 Bank-Postleitzahl
        # 14 Name der Bank
        #
        # CASE B: Postal code is not present in right table
        # 0  Handelsregisternummer
        # 1  Straße und Hausnummer
        # 2  Ort
        # 3  Telefon
        # 4  E-Mail-Adresse
        # 5  Handelsregisterstatus
        # 6  VVC-Status
        # 7  Umsatz
        # 8  Umsatzjahr
        # 9  UID-Nummer
        # 10 NACE Code
        # 11 NACE Beschreibung
        # 12 Bank-Postleitzahl
        # 13 Name der Bank
        # ------------------------------------------------------------

        postal_code_in_right_table = (
            len(values) >= 15
            and looks_like_postal_code(values[2])
        )

        if postal_code_in_right_table:
            values = values[:15]

            result = {
                "Datenbank-ID": database_id_number,
                "Handelsregisternummer": values[0],
                "Straße und Hausnummer": values[1],
                "Postleitzahl": values[2],
                "Ort": values[3],
                "Telefon": values[4],
                "E-Mail-Adresse": values[5],
                "Handelsregisterstatus": values[6],
                "VVC-Status": values[7],
                "Umsatzjahr": values[9],
                "Umsatz": values[8],
                "UID-Nummer": values[10],
                "NACE Rev. 2 - Haupttätigkeit - Code": values[11],
                "NACE Rev. 2 - Haupttätigkeit - Beschreibung": values[12],
                "Bank-Postleitzahl": values[13],
                "Name der Bank": values[14],
            }

        else:
            while len(values) < 14:
                values.append("Not found")

            if len(values) > 14:
                values = values[:14]

            result = {
                "Datenbank-ID": database_id_number,
                "Handelsregisternummer": values[0],
                "Straße und Hausnummer": values[1],
                "Postleitzahl": postal_code_from_left_table,
                "Ort": values[2],
                "Telefon": values[3],
                "E-Mail-Adresse": values[4],
                "Handelsregisterstatus": values[5],
                "VVC-Status": values[6],
                "Umsatzjahr": values[8],
                "Umsatz": values[7],
                "UID-Nummer": values[9],
                "NACE Rev. 2 - Haupttätigkeit - Code": values[10],
                "NACE Rev. 2 - Haupttätigkeit - Beschreibung": values[11],
                "Bank-Postleitzahl": values[12],
                "Name der Bank": values[13],
            }

        rows.append(
            {
                "seqnr": name_item.get("seqnr"),
                "company_name": clean_text(name_item.get("company_name")),
                "result": result,
            }
        )

    return {
        "rows": rows,
        "names_count": len(names),
        "right_rows_count": len(right_rows),
    }


def extract_visible_batch_rows(driver):
    snapshot = get_batch_snapshot(driver)

    rows = snapshot["rows"]

    if snapshot["names_count"] != snapshot["right_rows_count"]:
        print(
            f"WARNING: name rows and data rows differ. "
            f"names={snapshot['names_count']} | "
            f"data={snapshot['right_rows_count']} | "
            f"using={len(rows)}"
        )

    return rows


# ============================================================
# EXCEL OUTPUT
# ============================================================

def save_extract_file(extracted_records):
    SAMPLE_DATA_DIR.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted_Batch"

    ws.append(OUTPUT_HEADERS)

    for record in extracted_records:
        result = record["result"]

        row = [
            record.get("page", ""),
            record.get("seqnr", ""),
            record.get("company_name", ""),
            result.get("Datenbank-ID", "Not found"),
            result.get("Handelsregisternummer", "Not found"),
            result.get("Straße und Hausnummer", "Not found"),
            result.get("Postleitzahl", "Not found"),
            result.get("Ort", "Not found"),
            result.get("Telefon", "Not found"),
            result.get("E-Mail-Adresse", "Not found"),
            result.get("Handelsregisterstatus", "Not found"),
            result.get("VVC-Status", "Not found"),
            result.get("Umsatzjahr", "Not found"),
            result.get("Umsatz", "Not found"),
            result.get("UID-Nummer", "Not found"),
            result.get("NACE Rev. 2 - Haupttätigkeit - Code", "Not found"),
            result.get("NACE Rev. 2 - Haupttätigkeit - Beschreibung", "Not found"),
            result.get("Bank-Postleitzahl", "Not found"),
            result.get("Name der Bank", "Not found"),
        ]

        ws.append(row)

    for col in range(1, len(OUTPUT_HEADERS) + 1):
        cell = ws.cell(row=1, column=col)
        font = copy(cell.font)
        font.bold = True
        cell.font = font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for column_cells in ws.columns:
        max_length = 0
        column_letter = column_cells[0].column_letter

        for cell in column_cells:
            value = clean_text(cell.value)

            if len(value) > max_length:
                max_length = len(value)

        adjusted_width = min(max(max_length + 2, 12), 55)
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(OUTPUT_FILE)


# ============================================================
# MAIN EXTRACTION
# ============================================================

def extract_all_batch_results(driver):
    extracted_records = []
    total_rows_read = 0

    while True:
        current_page, total_pages = get_page_status(driver)

        rows = extract_visible_batch_rows(driver)
        total_rows_read += len(rows)

        print(f"Page {current_page}/{total_pages} | visible rows: {len(rows)}")

        for row in rows:
            record = {
                "page": current_page,
                "seqnr": row.get("seqnr", ""),
                "company_name": row.get("company_name", ""),
                "result": row["result"],
            }

            extracted_records.append(record)

            print(
                f"  EXTRACT | Page {current_page} | "
                f"SeqNr={record['seqnr']} | "
                f"{shorten(record['company_name'])}"
            )

        if current_page >= total_pages:
            break

        moved = go_to_next_page(driver)

        if not moved:
            print("WARNING: Could not move to next page. Stopping pagination.")
            break

    return extracted_records, total_rows_read


def main():
    start_time = datetime.now()

    print(f"[START] {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Output file: {OUTPUT_FILE}")
    print("Connecting to open Chrome session...")

    try:
        driver = connect_to_open_chrome()
        find_result_list_tab(driver)

        print("Batch result tab found.")
        print("Starting extraction from batch result list...")
        print("")

        extracted_records, total_rows_read = extract_all_batch_results(driver)

        print("")
        print(f"Saving extracted data to: {OUTPUT_FILE}")
        save_extract_file(extracted_records)

        end_time = datetime.now()
        duration = end_time - start_time

        print("")
        print(f"[END] {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Duration: {duration}")
        print(f"Rows read: {total_rows_read}")
        print(f"Rows saved to extract file: {len(extracted_records)}")
        print(f"Output file: {OUTPUT_FILE}")
        print("Done. Batch extract file created.")

    except Exception as error:
        print("")
        print("ERROR:")
        print(error)
        raise


if __name__ == "__main__":
    main()