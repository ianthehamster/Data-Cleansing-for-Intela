function main(workbook: ExcelScript.Workbook) {
    // Get the active sheet and table
    let sheet = workbook.getActiveWorksheet();
    let table = sheet.getTables()[0];

    // Add 2 new columns after the 15th column
    table.addColumn(16);
    table.addColumn(17);

    // Get the range of the table
    let range = table.getRange();

    // Get the values in the range
    let values1 = range.getValues();
    let values2 = range.getValues();

    // Process the "Assignee" column values
    let columnValues: (string | null)[][] = table.getColumnByName("Assignee").getRangeBetweenHeaderAndTotal().getValues() as (string | null)[][];
    let columnValuesCleansed: string[] = [];

    for (let i = 0; i < columnValues.length; i++) {
        let value = columnValues[i][0];
        if (value !== "") {
            columnValuesCleansed.push(value);
        }
    }

    // Remove duplicates from the Assignee column
    function removeDuplicates(arr: string[]): string[] {
        let unique: string[] = [];
        arr.forEach(assignee => {
            if (!unique.includes(assignee)) {
                unique.push(assignee);
            }
        });
        return unique;
    }

    let columnValuesCleansedAndUnique = removeDuplicates(columnValuesCleansed);

    // Process each assignee and filter the table based on their email
    columnValuesCleansedAndUnique.forEach(assigneeEmail => {
        table.getAutoFilter().apply(table.getAutoFilter().getRange(), 6, { filterOn: ExcelScript.FilterOn.values, values: [assigneeEmail] });
        table.getAutoFilter().apply(table.getAutoFilter().getRange(), 5, { filterOn: ExcelScript.FilterOn.values, values: ["Overdue"] });

        let filteredRange = table.getRange().getVisibleView().getValues();

        if (filteredRange.length == 1) return;

        let newSheetName: string = assigneeEmail;
        let newSheet = workbook.addWorksheet(newSheetName);
        const assigneesData = newSheet.getRangeByIndexes(0, 0, filteredRange.length, filteredRange[0].length);
        assigneesData.setValues(filteredRange);
        const newTable = newSheet.addTable(assigneesData, true);
        table.getAutoFilter().apply(table.getRange());
    });

    // Split the values in the "Assignee" column if they contain a semicolon and add new rows
    for (let i = 1; i < values1.length; i++) {
        let currentRow = values1[i];
        let assignee = values1[i][6];
        if (assignee !== "" && assignee.includes(";")) {
            const splitValues: string[] = assignee.split(";");

            let newRow1 = [...currentRow];
            newRow1[6] = splitValues[0];

            let newRow2 = [...currentRow];
            newRow2[6] = splitValues[1];

            values1[i] = newRow1;
            table.addRow(-1, newRow2);
        }
    }

    // After the loop, update the range with the modified values
    range.setValues(values1);

    // Split the values in the "Work Area" column and assign them to the new columns
    for (let i = 0; i < values2.length; i++) {
        let workArea = values2[i][15];
        let splitValuesOfWorkArea: string[] = workArea.split(' | ');

        // Update the original row with the split values in the adjacent column
        values1[i][15] = splitValuesOfWorkArea[0];
        values1[i][16] = splitValuesOfWorkArea[1];
        values1[i][17] = splitValuesOfWorkArea[2];
    }

    // Update the range with the new values
    range.setValues(values1);

    // Update the header row for the new columns
    const headers = table.getHeaderRowRange().getValues()[0];
    console.log(headers);

    headers[16] = "Region";
    headers[17] = "Statement Type";
    table.getHeaderRowRange().setValues([headers]);

  // Clear all filters at the end by resetting the auto-filter range
  table.getAutoFilter().apply(table.getAutoFilter().getRange(), 6, { filterOn: ExcelScript.FilterOn.values, values: [] });
  table.getAutoFilter().apply(table.getAutoFilter().getRange(), 5, { filterOn: ExcelScript.FilterOn.values, values: [] });
}
