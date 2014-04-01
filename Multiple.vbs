On Error Resume Next 
Const xpRangeAutoFormatList2 = 11 
Const HKEY_LOCAL_MACHINE = &H80000002
 
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True 
Set objWorkbook = objExcel.Workbooks.Add() 
Set objWorksheet = objWorkbook.Worksheets(1) 
objExcel.Cells(1, 1).Value = "Server" 
objExcel.Cells(1, 1).Font.Italic = True 
objExcel.Cells(1, 1).Font.Bold = True 
objExcel.Cells(1, 2).Value = "Instance" 
objExcel.Cells(1, 2).Font.Italic = True 
objExcel.Cells(1, 2).Font.Bold = True 
objExcel.Cells(1, 3).Value = "Edition" 
objExcel.Cells(1, 3).Font.Italic = True 
objExcel.Cells(1, 3).Font.Bold = True 
objExcel.Cells(1, 4).Value = "Build" 
objExcel.Cells(1, 4).Font.Italic = True 
objExcel.Cells(1, 4).Font.Bold = True 

x = 2 

Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objFile = objFSO.OpenTextFile("C:\Temp\Serverlist.txt", 1)
 
Do Until objFile.AtEndOfStream 
  strComputer = objFile.ReadLine
  Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
  strKeyPath = "SOFTWARE\Microsoft\Microsoft SQL Server"
  strValueName = "InstalledInstances"
  objReg.GetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,arrValues
	If IsNull(arrValues) = 0 Then
	  For Each strValue In arrValues	  
		strKeyPath = "SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL"
		strValueName = strValue
		strSQLInstance = strValue
		objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strVersion
		strKeyPath = "SOFTWARE\Microsoft\Microsoft SQL Server\" & strVersion & "\Setup"
		strValueName = "Edition"
		objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strSQLEdition	
		strValueName = "Version"
		objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strSQLVersion
		strValueName = "ProductCode"
		objReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strSQLProductCode	
  	  	WScript.echo strComputer & " - " & strSQLInstance & " - " & strSQLEdition & " - " & strSQLVersion
          	objExcel.Cells(x, 1) = strComputer  
          	objExcel.Cells(x, 2) = strSQLInstance
          	objExcel.Cells(x, 3) = strSQLEdition
          	objExcel.Cells(x, 4) = strSQLVersion
	  	x = x + 1 
	  Next
	End If
Set arrValues = Nothing
Set strSQLEdition = Nothing
Set strSQLVersion = Nothing
Set objReg = Nothing
Loop 
Set objRange = objWorksheet.UsedRange 
objRange.AutoFormat(xpRangeAutoFormatList2) 
objRange.EntireColumn.Autofit() 
Set objRange5 = objExcel.Range("E1") 
objRange.Sort objRange5,,,,,,,1 