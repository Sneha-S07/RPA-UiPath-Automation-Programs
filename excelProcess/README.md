# Excel Sheet and Column Operations – UiPath

## Description
This UiPath automation demonstrates Excel workbook and worksheet management using Excel activities. It performs operations such as iterating through worksheets, displaying sheet names, inserting columns, splitting data into columns, and deleting a specified column.

## Workflow Overview

1. **Excel Process Scope:** Initializes the Excel automation process.
2. **Use Excel File:** Opens the specified workbook using the reference `Excel`. The options to create the file if it does not exist and save changes are enabled.
3. **Delay:** Adds a 5-second delay before performing Excel operations.
4. **For Each Excel Sheet:** Iterates through all worksheets in the workbook using the `CurrentSheet` variable.
5. **Message Box:** Displays the name of each worksheet using `CurrentSheet.Name`.
6. **Insert Column:** Inserts a new column into the worksheet.
7. **Text to Columns:** Splits the data starting from cell A1 in the `New Sheet1` worksheet into separate columns.
8. **Delete Column:** Deletes the column named `usn` from the `CSD 7th sem` worksheet. The worksheet is configured to contain headers.

## Additional Worksheet Operations

The workflow also contains a **Comment Out** section with worksheet management activities. These activities are currently disabled and are not executed.

The commented-out operations include:

- Inserting a new worksheet named `New Sheet1`.
- Renaming `Sheet1` to `CSD 7th sem`.
- Duplicating `Sheet1` and renaming the duplicate to `Batch 2`.
- Deleting the `New Sheet` worksheet.

## UiPath Activities Used

- Excel Process Scope
- Use Excel File
- Delay
- For Each Excel Sheet
- Message Box
- Insert Sheet
- Rename Sheet
- Duplicate Sheet
- Delete Sheet
- Insert Column
- Text to Columns
- Delete Column
- Comment Out

## Requirements

- UiPath Studio with Excel activities support.
- Microsoft Excel or a compatible Excel workbook.
- A valid workbook path configured in the **Use Excel File** activity.
- The referenced worksheets and column headers should exist in the workbook before the corresponding operations are executed.

## Execution

1. Open the project in UiPath Studio.
2. Update the workbook path in the **Use Excel File** activity according to your system.
3. Verify that the referenced worksheet names and column headers match the workbook.
4. Run the workflow.
5. Check the workbook to verify the Excel operations and saved changes.

## Expected Outcome

The automation iterates through the worksheets and displays their names. It also performs the configured column insertion, text-to-columns conversion, and deletion of the specified `usn` column. The workbook is saved after execution.
