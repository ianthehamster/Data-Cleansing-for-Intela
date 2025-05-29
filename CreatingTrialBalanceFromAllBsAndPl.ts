
function main(workbook: ExcelScript.Workbook) {
    const summarySheet = workbook.getWorksheet("Summary")
    const summaryRange = summarySheet.getUsedRange()
    const summaryValues = summaryRange.getValues()
    let sheetCountExceptForSummaryAndHiddenSheets = 0

    // Build a mapping from sheet name (column D) to Entity Code (B) and Entity Name (E)
    const sheetMetadata = new Map<string, { code:string, name: string}>()
    for(let i = 4; i < summaryValues.length; i++){
        const sheetName = summaryValues[i][3] as string; // Column D
        const entityCode = summaryValues[i][1] as string; // Column B
        const entityName = summaryValues[i][4] as string // Column E

        if(sheetName){
            sheetMetadata.set(sheetName, {code: entityCode, name: entityName})
        }
    }

    // Prepare a master Trial Balance array
    let trialBalance: (string | number)[][] = []
    trialBalance.push(["Entity Code", "Entity Name", "Account Code", "Account Name", "Amount"])

    // Process each sheet except "Summary"
    for(const sheet of workbook.getWorksheets()){
        const sheetName = sheet.getName()
        if(sheetName === "Summary") continue;

        const meta = sheetMetadata.get(sheetName)
        console.log(meta)

        if(!meta){
            console.log(`No metadata found for sheet ${sheetName}`)
            continue
        } else {
            sheetCountExceptForSummaryAndHiddenSheets++
        }

        const dataRange = sheet.getRange(`B10:E${sheet.getUsedRange().getRowCount() + 10}`)
        const dataValues = dataRange.getValues()

        for(let row of dataValues){
            const accountCode = row[0] as string
            const accountName = row[1] as string
            const amount = row[3] as number

            if(!accountCode || amount === undefined || amount === null) continue;

            trialBalance.push([
                meta.code,
                meta.name,
                accountCode,
                accountName,
                amount
            ])
        }

    }
    console.log(sheetCountExceptForSummaryAndHiddenSheets)
    // Output to a new worksheet
    const outputSheet = workbook.addWorksheet("Trial Balance Output")
    outputSheet.getRangeByIndexes(0,0, trialBalance.length, 5).setValues(trialBalance)

    // Clean column D (Account Name)
    const accountNameRange = outputSheet.getRangeByIndexes(1, 3, trialBalance.length - 1, 1)
    const accountNameValues = accountNameRange.getValues()

    for(let i = 0; i < accountNameValues.length; i++){
        const rawName = accountNameValues[i][0] as string
        if(rawName){
            const cleaned = rawName.replace(/^\s*([0-9A-Za-z]+)/, ""); // Match alphanumeric characters starting chunk
            accountNameValues[i][0] = cleaned
        }
    }

    accountNameRange.setValues(accountNameValues)

}
