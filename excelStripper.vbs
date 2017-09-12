Function GetFileName( myDir, myFilter)
'opens a dialogbox and returns a string
	Dim objDialog
	
	Set objDialog = CreateObject( "UserAccounts.CommonDialog")
	
	If myDir = "" Then
		objDialog.InitialDir = CreateObject( "WScript.Shell" ).SpecialFolders( "MyDocuments" )
	Else
		objDialog.InitialDir = myDir
	End If
	
	If myFilter = "" Then
		objDialog.Filter = "All files|*.*"
	Else
		objDialog.Filter = myFilter
	End If
	
	If objDialog.ShowOpen Then 
		GetFileName = objDialog.FileNmae
	Else
		GetFileName = ""
	End If
End Function

Set fileName = GetFileName("C:\Users\kyle.goode\","")

Set objExcel = CreateObject("Excel.Application")

objExcel.Application.Visible = True

Replace keywordInput, ", ", ","

Set keyWords = Split(keywordInput, ",")

Set objWorkbook = objExcel.Workbooks.Open(locationInput)

For Each i In keyWords
	
	i = 1
	
	Do Until objExcel.Cells(i, 1).Value = ""
	
		If objExcel.Cells(i, 1).Value = x Then
		
			Set objRange = objExcel.Cells(i, 1).EntireRow
			
			objRange.Delete
			
			i = i - 1
			
		End If
		
		i = i + 1
		
	Loop

Next

objExcel.ActiveWorkbook.Save locationInput

objExcel.ActiveWorkbook.Close

objExcel.Application.Quit