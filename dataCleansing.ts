function main(workbook: ExcelScript.Workbook) {

    let sheet = workbook.getActiveWorksheet();

    let cellA1Value = sheet.getRange("A1").getValue();
    if (!cellA1Value) {
        sheet.getRange("1:4").delete(ExcelScript.DeleteShiftDirection.up);
    }

    sheet.getShapes().forEach(shape => {
        if (shape.getType() === ExcelScript.ShapeType.image) shape.delete();
    });

    let usedRange = sheet.getUsedRange();
    if (!usedRange) return;

    let table: ExcelScript.Table;

    if (sheet.getTables().length < 1) {
        table = sheet.addTable(usedRange, true);
    } else {
        table = sheet.getTables()[0];
    }

    let assigneeColumnIndex = table.getHeaderRowRange().getValues()[0]
        .findIndex(header => header.trim().toLowerCase() === "assignee");

    if (assigneeColumnIndex === -1) return;

    let rows = table.getRangeBetweenHeaderAndTotal().getValues();

    let filteredRows = rows.filter(row => {
        let assigneeCell = row[assigneeColumnIndex];
        return assigneeCell && assigneeCell.toString().toLowerCase().includes("deloitte");
    });

    if (filteredRows.length === 0) return;

    let newSheet = workbook.addWorksheet("Deloitte Only");
    let headers = table.getHeaderRowRange().getValues()[0];

    newSheet.getRange("A1").getResizedRange(0, headers.length - 1).setValues([headers]);

    newSheet.getRange("A2").getResizedRange(filteredRows.length - 1, headers.length - 1).setValues(filteredRows);

    newSheet.addTable(newSheet.getUsedRange(), true);

    let uniqueAssignees: string[] = [];
    filteredRows.forEach(row => {
        let assignee = row[assigneeColumnIndex].toString();
        if (!uniqueAssignees.includes(assignee)) {
            uniqueAssignees.push(assignee);
        }
    });
    
    uniqueAssignees.forEach(assignee => {
      if (assignee != undefined || assignee != "" || !assignee.includes(':')) {
        console.log(assignee)
        console.log(assignee.length)
        }
    })

    uniqueAssignees.forEach(assignee => {

        if(!assignee.includes(':') && assignee.length < 30){
          
            let assigneeRows = filteredRows.filter(row => row[assigneeColumnIndex] === assignee);

            let assigneeSheet = workbook.addWorksheet(assignee);
            assigneeSheet.getRange("A1").getResizedRange(0, headers.length - 1).setValues([headers]);

            assigneeSheet.getRange("A2").getResizedRange(assigneeRows.length - 1, headers.length - 1).setValues(assigneeRows);
            assigneeSheet.addTable(assigneeSheet.getUsedRange(), true);
        }
       
    });
}
