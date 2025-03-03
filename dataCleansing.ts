function main(workbook: ExcelScript.Workbook) {

    // Get the active sheet and table
    let sheet = workbook.getActiveWorksheet();
  
    // Delete the first 4 rows only if it is a new workplan file
    let cellA1Value = sheet.getRange("A1").getValue()
    if (cellA1Value === "" || !cellA1Value) {
      sheet.getRange("1:4").delete(ExcelScript.DeleteShiftDirection.up)
    }
  
    // Remove all images in the worksheet
    let shapes = sheet.getShapes()
    for (let shape of shapes) {
      console.log(shape.getType())
      if (shape.getType() === ExcelScript.ShapeType.image) {
        shape.delete()
      }
    }
  
    // Getting used range of the sheet (all cells with data)
    let usedRange = sheet.getUsedRange()
    if (!usedRange) {
      console.warn("No data found in the workplan tasks")
      return
    }
  
    let rowCount = usedRange.getRowCount()
    let colCount = usedRange.getColumnCount()
  
    console.log(rowCount, colCount)
  
    let totalTables = sheet.getTables()
  
    let mainTable: ExcelScript.Table
  
    if (totalTables.length < 1) {
      mainTable = sheet.addTable(usedRange, true) // true means it has headers
    }
  
    let table = sheet.getTables()[0]
  
    // Get the table headers from table
    let headersOfCreatedTable: string[] = table.getHeaderRowRange().getValues()[0] as string[]
  
    // Find the column index for Assignee
    let assigneeColumnIndex = headersOfCreatedTable.findIndex(header=> header.trim().toLowerCase() === "assignee")
  
    console.log(assigneeColumnIndex, headersOfCreatedTable)
  
    // Get the range of the table
    let range = table.getRange();
  
    // Get the values in the range
    let values1 = range.getValues();
    let values2 = range.getValues();
  
    // Process the "Assignee" column values
    let columnValues: (string | null)[][] = table.getColumnByName("Assignee").getRangeBetweenHeaderAndTotal().getValues() as (string | null)[][];
    let columnValuesCleansed: string[] = [];
  
    let assigneeColumnValuesDeloiteOnly: string[] = []
  
    for (let i = 0; i < columnValues.length; i++) {
      let value = columnValues[i][0];
      if (value !== "") {
        columnValuesCleansed.push(value);
      }
    }
  
    console.log(columnValuesCleansed)
  
    
  
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
  
    console.log(columnValuesCleansedAndUnique)
  
    function cleanData(inputArray: string[]): string[] {
      let cleanedArray: string[] = [];
  
      // Step 1: Split strings containing ";"
      for (let str of inputArray) {
        if (str.includes(";")) {
          cleanedArray.push(...str.split(";").map(s => s.trim())); // Trim spaces after splitting
        } else {
          cleanedArray.push(str);
        }
      }
  
      // Step 2: Remove text before ":" including ":"
      cleanedArray = cleanedArray.map(str => {
        if (str.includes(":")) {
          return str.split(":").pop()?.trim() || ""; // Keep only the part after the last ":"
        }
        return str;
      });
  
      // Step 3: Keep only strings containing "deloitte"
      cleanedArray = cleanedArray.filter(str => str.toLowerCase().includes("deloitte"));
  
      return cleanedArray;
    }
  
    assigneeColumnValuesDeloiteOnly = cleanData(columnValuesCleansedAndUnique)
  
    console.log(assigneeColumnValuesDeloiteOnly)
  
    
  
    // Process each assignee and filter the table based on their email
    columnValuesCleansedAndUnique.forEach(assigneeEmail => {
      table.getAutoFilter().apply(table.getAutoFilter().getRange(), 12, { filterOn: ExcelScript.FilterOn.values, values: [assigneeEmail] });
      table.getAutoFilter().apply(table.getAutoFilter().getRange(), 8, { filterOn: ExcelScript.FilterOn.values, values: ["Not Started"] });
  
      let filteredRange = table.getRange().getVisibleView().getValues();
  
      if (filteredRange.length == 1) return;
  
      let newSheetName: string = assigneeEmail;
      let newSheet = workbook.addWorksheet(newSheetName);
      const assigneesData = newSheet.getRangeByIndexes(0, 0, filteredRange.length, filteredRange[0].length);
      assigneesData.setValues(filteredRange);
      const newTable = newSheet.addTable(assigneesData, true);
      table.getAutoFilter().apply(table.getRange());
    });
    return
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
  