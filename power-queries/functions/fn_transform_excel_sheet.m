// Query   : fn_transform_excel_sheet
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-24T17:03:57+00:00

let
        Source = (ExcelWorkbook as binary, SheetName as text) =>
            let
                Source = Excel.Workbook(ExcelWorkbook, null, true), Sheet1 = Source{[Name = SheetName]}[Data]
            in
                Sheet1
    in
        Source
