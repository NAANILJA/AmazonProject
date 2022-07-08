SystemUtil.CloseProcessByName"chrome.exe"
SystemUtil.Run"chrome.exe"
'Browser(browserObject).Navigate"https://www.amazon.in/"

On Error Resume Next
FilePath= "C:\Users\user255\Documents\AmazonProject\Test Data\Test Data.xlsx"
excelSheet="Test Data"
Sheetname="Amazon Data"

DataTable.AddSheet excelSheet
DataTable.ImportSheet FilePath, Sheetname , excelSheet

rowCount = DataTable.GetSheet(excelSheet).GetRowCount

For rows= 1 To rowCount

DataTable.SetCurrentRow rows

If DataTable.Value("Expected_Flag",excelSheet)= "y" Then

executeTest(DataTable.Value("TC_ID",excelSheet))

DataTable.Value("Result",excelSheet) = Environment.Value("Result")
End If

Next
DataTable.ExportSheet FilePath, excelSheet, Sheetname

