# Excel Automation Script - ReadMe

## Overview
This script automates the processing of an Excel workbook using **ExcelScript** in TypeScript. It cleans up data, processes specific columns, filters data, and dynamically generates new sheets based on filtered values.

## Features
- Deletes the first 4 rows **only if the sheet is a new workplan file**.
- Removes all **images** from the worksheet.
- **Detects and creates a table** from the used range if none exist.
- **Finds the 'Assignee' column** and processes its data.
- **Splits multi-value cells** containing `;` into multiple rows.
- **Cleans 'Assignee' column values** by removing prefixes before `:` and filtering values containing 'deloitte'.
- **Filters tasks by 'Assignee' and 'Not Started' status** and creates individual sheets per assignee.
- Splits the **'Work Area' column** into separate columns.
- Updates headers for newly added columns.
- Resets filters at the end to maintain data integrity.

## Script Workflow
1. **Initial Cleanup**
   - Checks if `A1` is empty to determine if it is a new file.
   - If yes, deletes the first **four rows**.
   - Removes all **images** from the worksheet.

2. **Creating/Detecting a Table**
   - Retrieves the used range.
   - If no table exists, creates one with headers.
   - Extracts table headers and **finds the 'Assignee' column' index**.

3. **Processing 'Assignee' Column**
   - Extracts **'Assignee'** column values.
   - Cleans the values by:
     - Splitting values containing `;` into separate rows.
     - Removing text before `:` (including `:` itself).
     - Filtering only values containing 'deloitte'.
   
4. **Filtering & Sheet Generation**
   - **Filters 'Assignee' column** for each unique value.
   - Further filters tasks where **'Status' = 'Not Started'**.
   - Creates **a new worksheet per assignee** and populates it with their filtered tasks.

5. **Processing 'Work Area' Column**
   - Splits values containing `|` into separate columns.
   - Updates the original row with **newly split values**.
   - Adds **new column headers** ('Region' and 'Statement Type').

6. **Final Cleanup**
   - Clears filters to **reset the table state**.

## Key Functions
### `removeDuplicates(arr: string[])`
- Removes duplicate values from an array.

### `cleanData(inputArray: string[])`
- Cleans and filters 'Assignee' column values by:
  - Splitting values with `;`.
  - Removing prefixes before `:`.
  - Keeping only values containing 'deloitte'.

## Requirements
- Must be executed within an **Excel environment supporting ExcelScript**.
- Ensure the workbook has a valid structure with an **'Assignee' column** and task data.

## Expected Output
- Cleaned and formatted table with organized data.
- Individual sheets generated for each assignee with their tasks.
- Properly structured columns with new headers ('Region', 'Statement Type').

