SystemUtil.CloseProcessByName"chrome.exe"
SystemUtil.Run"chrome.exe"
Browser(browserObject).Navigate"https://www.amazon.in/"

On Error Resume Next
DataTable.AddSheet"Test Data"
DataTable.ImportSheet "C:\Users\user255\Documents\AmazonProject\Test Data\Test Data.xlsx","Amazon Data","Test Data"

rowCount = DataTable.GetSheet("Test Data").GetRowCount

For rows= 1 To rowCount

DataTable.SetCurrentRow rows

If DataTable.Value("Expected_Flag","Test Data")= "y" Then

executeTest(DataTable.Value("TC_ID","Test Data"))
 @@ script infofile_;_ZIP::ssf142.xml_;_
 @@ script infofile_;_ZIP::ssf141.xml_;_
'Environment.Value("Result")="Pass"
DataTable.Value("Result","Test Data") = Environment.Value("Result")
End If
	
Next
DataTable.ExportSheet "C:\Users\user255\Documents\AmazonProject\Test Data\Test Data.xlsx","Test Data","Amazon Data"

