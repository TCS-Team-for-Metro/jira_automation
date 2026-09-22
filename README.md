# Data flow from browser open to final file

This automation does the following in order:

1. Opens a Chromium browser and logs in to Jira through Microsoft authentication.
2. Downloads the Jira CSV export.
3. Cleans, reorders, and filters the CSV columns.
4. Splits the data into two Excel sheets: `CZ&SK` and `Ukraine`.
5. Formats the workbook and saves the final `.xlsx` file.
6. Deletes the original CSV and closes the browser.

## 1) Browser launch and page setup

Entry point: `main()` in `src/web_automation/__init__.py`

Function flow:
- `main()`
- `sync_playwright()`
- `p.chromium.launch(headless=False)`
- `browser.new_context(accept_downloads=True)`
- `context.new_page()`

This is the start of the run. The browser is now ready for the Jira login sequence.

## 2) Microsoft/Jira login flow

Functions involved:
- `main()`
- `login_to_microsoft(page)`
- `is_jira_url(url)`
- `is_microsoft_login_url(url)`
- `fill_microsoft_email(page)`
- `click_primary_button(page, *texts)`

Flow:
- `main()` calls `login_to_microsoft(page)`.
- `login_to_microsoft(page)` calls `page.goto(JIRA_ISSUES_URL, wait_until="domcontentloaded", timeout=120000)`.
- The loop checks `is_jira_url(current_url)` and `is_microsoft_login_url(current_url)`.
- When the URL is Microsoft login, `fill_microsoft_email(page)` fills the email box.
- `click_primary_button(page, "Next", "Continue")` clicks the next step button.
- The loop repeats until the page reaches Jira.

If the flow does not reach Jira within 120 seconds, `login_to_microsoft(page)` raises `RuntimeError`.

## 3) CSV download from Jira

Functions involved:
- `main()`
- `build_download_filename()`
- `download_csv(page)`

Flow:
- `main()` calls `download_csv(page)` after login.
- `build_download_filename()` creates the download CSV name.
- `download_csv(page)` creates a Downloads folder and temp file path.
- `page.expect_download()` is used before the CSV export URL is opened.
- `page.evaluate("url => window.location.href = url", CSV_DOWNLOAD_URL)` triggers the download.
- `download.save_as(path=temp_file)` saves it to the temp file.
- `temp_file.replace(final_path)` renames it to the final CSV filename.
- `download_csv(page)` returns the final `Path` to `main()`.

At this point, the browser is still open, and the file on disk is the raw Jira CSV export.

## 4) Converting CSV to workbook

Function: `convert_csv_to_xlsx(csv_path)`

This is the core transformation stage.

### 4.1 Workbook creation

Functions involved:
- `convert_csv_to_xlsx(csv_path)`
- `create_status_counts()`

Flow:
- `Workbook()` creates the workbook.
- `workbook.active` becomes the `CZ&SK` sheet.
- `workbook.create_sheet("Ukraine")` creates the second sheet.
- `create_status_counts()` initializes counters for each status bucket.

### 4.2 Header processing and column removal

Functions involved:
- `build_columns_to_remove_indexes(header_row)`
- `remove_columns(row, columns_to_remove_indexes)`
- `filter_row(row, columns_to_remove_indexes)`
- `rename_columns(row)`
- `rename_headers(row)`

Flow:
- The first row is read from the CSV by the loop in `convert_csv_to_xlsx(csv_path)`.
- `build_columns_to_remove_indexes(row)` computes indexes to remove.
- `COLUMNS_TO_REMOVE` contains:
  - `Issue id`
  - `Project key`
  - `Project type`
  - `Project lead`
  - `Project description`
  - `Project url`
- `remove_columns(row, columns_to_remove_indexes)` removes those columns.
- `filter_row(...)` is a thin wrapper that calls `remove_columns(...)`.
- `rename_columns(row)` changes `Custom field (Product)` to `Product` and `Custom field (Solution)` to `Solution`.
- `rename_headers(row)` is a compatibility wrapper calling `rename_columns(row)`.

This is the first major cleanup. The raw Jira headers are reduced to the columns the script actually needs.

### 4.3 Column re-positioning

Functions involved:
- `move_project_name_to_start(row, project_name_index)`
- `swap_product_and_solution(row, product_index, solution_index)`
- `reposition_columns(row, project_name_index, product_index, solution_index)`

There are two explicit reorder steps.

#### A. Move Project name to the start

Function: `move_project_name_to_start(row, project_name_index)`

Flow:
- `convert_csv_to_xlsx(csv_path)` finds the index of `Project name`.
- `move_project_name_to_start(...)` removes the column from its current position.
- It inserts the same column at index `0`.

This ensures `Project name` is always the first column.

#### B. Swap Product and Solution columns

Function: `swap_product_and_solution(row, product_index, solution_index)`

Flow:
- `convert_csv_to_xlsx(csv_path)` finds the `Product` and `Solution` indexes.
- `swap_product_and_solution(...)` swaps the values in place.

This is done for the header row and again for every transformed data row.

#### C. Combined reposition step

Function: `reposition_columns(row, project_name_index, product_index, solution_index)`

Flow:
- `reposition_columns(...)` first calls `move_project_name_to_start(...)`.
- Then it calls `swap_product_and_solution(...)` on the resulting row.
- This provides one reusable function for the full column reordering operation.

This is the main column re-positioning helper used during CSV-to-XLSX transformation.

#### D. Final preferred output order

Function: `reorder_row_to_preferred_columns(row, preferred_column_order)`

Flow:
- The final transformation applies the exact column sequence requested for the workbook output:
  - `Project name`
  - `Solution`
  - `Product`
  - `Issue key`
  - `Summary`
  - `Assignee`
  - `Reporter`
  - `Priority`
  - `Status`
  - `Created`
  - `Updated`
  - `Time to resolution`
  - `Labels` (kept in all three label columns)
- The function normalizes the CSV names, removes extra columns, and reconstructs each row in that exact order.
- This ensures both the header row and every data row follow the same arrangement before the workbook is saved.

### 4.4 Finding important columns

Functions involved:
- `convert_csv_to_xlsx(csv_path)`
- `next(...)` index scanning logic

Flow:
- `convert_csv_to_xlsx(csv_path)` locates indexes for:
  - `Issue key`
  - `Assignee`
  - `Status`
- These are stored as variables like:
  - `issue_key_index`
  - `assignee_index`
  - `status_index`

These indexes drive later row routing and status counting.

### 4.5 Header row is written to both sheets

Functions involved:
- `convert_csv_to_xlsx(csv_path)`
- `czsk_sheet.append(filtered_row)`
- `ukraine_sheet.append(filtered_row)`

Flow:
- After filtering and reordering, the header is appended to both sheets.
- This creates the column structure for both country sections.

## 5) Processing each issue row

Functions involved:
- `convert_csv_to_xlsx(csv_path)`
- `filter_row(row, columns_to_remove_indexes)`
- `move_project_name_to_start(...)`
- `swap_product_and_solution(...)`
- `set_default_assignee(row, assignee_index)`
- `has_assigned_owner(row, assignee_index)`
- `update_status_counts(row, status_index, status_counts)`

Flow for each data row:
1. `filter_row(row, columns_to_remove_indexes)` removes unwanted columns.
2. `move_project_name_to_start(...)` reorders the row so `Project name` is first.
3. `swap_product_and_solution(...)` places `Product` and `Solution` in the expected order.
4. `set_default_assignee(filtered_row, assignee_index)` fills blank assignees with `Unassigned`.
5. `issue_key = filtered_row[issue_key_index].strip()` extracts the issue key.
6. If issue key starts with:
   - `SDMCCCS` -> `czsk_sheet.append(filtered_row)`
   - `SDMCCUA` -> `ukraine_sheet.append(filtered_row)`
7. If the owner is valid, `has_assigned_owner(...)` returns `True` and `update_status_counts(...)` increments the matching status count.

### 5.1 Assignee normalization

Function: `set_default_assignee(row, assignee_index)`

Flow:
- If the assignee cell is empty, it replaces it with `Unassigned`.
- The row is then used consistently in later checks.

### 5.2 Assigned owner check

Function: `has_assigned_owner(row, assignee_index)`

Flow:
- Reads the assignee value from the row.
- Returns `False` if empty or equal to `Unassigned`.
- Returns `True` otherwise.

### 5.3 Status counting

Function: `update_status_counts(row, status_index, status_counts)`

Flow:
- Reads the status value from the current row.
- Lowercases it.
- Maps it into:
  - `open`, `delegated`, `feedback received` -> `work in progress`
  - `pending` -> `pending`
  - `waiting for user` -> `waiting for user`
  - `waiting for acceptance` -> `waiting for acceptance`
  - `closed` -> `closed`

This happens only for rows with valid owners and only for the relevant sheet.

## 6) Worksheet formatting

Functions involved:
- `format_worksheet(worksheet)`
- `create_status_counts()`
- `print_status_counts(sheet_name, status_counts)`

Flow:
- After all rows are processed, `convert_csv_to_xlsx(csv_path)` calls `format_worksheet(czsk_sheet)` and `format_worksheet(ukraine_sheet)`.
- `format_worksheet(worksheet)` applies:
  - header fill color
  - bold font
  - centered alignment
  - width tuning for columns
- `print_status_counts(...)` prints the final counts for each sheet.

## 7) Final file creation

Functions involved:
- `build_output_filename()`
- `convert_csv_to_xlsx(csv_path)`

Flow:
- `build_output_filename()` creates the date-based XLSX filename.
- `convert_csv_to_xlsx(csv_path)` calls `workbook.save(xlsx_path)`.
- `csv_path.unlink()` removes the original CSV.
- The final file path is returned to `main()`.

This is the end of the transformation stage.

## 8) End of run

Functions involved:
- `main()`
- `browser.close()`

Flow:
- `main()` prints the final XLSX location.
- It returns `0` on success.
- In the `finally` block, `browser.close()` closes the browser.

This is the final shutdown step for the automation run.

## Function-by-function data path

- `main()`
  - starts browser
  - calls `login_to_microsoft(page)`
  - calls `download_csv(page)`
  - calls `convert_csv_to_xlsx(csv_path)`
- `login_to_microsoft(page)`
  - checks `is_jira_url(url)` / `is_microsoft_login_url(url)`
  - calls `fill_microsoft_email(page)`
  - calls `click_primary_button(page, *texts)`
- `download_csv(page)`
  - calls `build_download_filename()`
  - downloads the Jira CSV
- `convert_csv_to_xlsx(csv_path)`
  - calls `build_columns_to_remove_indexes(header_row)`
  - calls `filter_row(...)`
  - calls `rename_headers(...)`
  - calls `move_project_name_to_start(...)`
  - calls `swap_product_and_solution(...)`
  - calls `set_default_assignee(...)`
  - calls `has_assigned_owner(...)`
  - calls `update_status_counts(...)`
  - calls `format_worksheet(...)`
  - calls `print_status_counts(...)`
  - saves workbook and deletes CSV

## Summary of the data path

Raw Jira export -> `download_csv(page)` -> CSV file on disk -> `convert_csv_to_xlsx(csv_path)` -> `build_columns_to_remove_indexes(...)` -> `filter_row(...)` -> `rename_headers(...)` -> `move_project_name_to_start(...)` -> `swap_product_and_solution(...)` -> `set_default_assignee(...)` -> `has_assigned_owner(...)` -> `update_status_counts(...)` -> `format_worksheet(...)` -> final `.xlsx` -> `csv_path.unlink()` -> `browser.close()`

This is the complete flow of the data from the browser open through the final generated Excel file.
