// Query   : fn_transform_excel_sheet
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-03-03T17:39:11+00:00

let
        Source = (ExcelWorkbook as binary, SheetName as text) =>
            let
                Source = Excel.Workbook(ExcelWorkbook, null, true), Sheet1 = Source{[Name = SheetName]}[Data]
            in
                Sheet1
    in
        Source
