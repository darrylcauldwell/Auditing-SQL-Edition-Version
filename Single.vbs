Const HKEY_LOCAL_MACHINE = &H80000002
strServer = "TESTSERVER"

Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strServer & "\root\default:StdRegProv")


  strKeyPath = "SOFTWARE\Microsoft\Microsoft SQL Server"
  strValueName = "InstalledInstances"
  objReg.GetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPath, strValueName,arrValues
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
  	  WScript.echo strServer & " - " & strSQLInstance & " - " & strSQLEdition & " - " & strSQLVersion
	  Next
	End If 

 Set objReg = Nothing