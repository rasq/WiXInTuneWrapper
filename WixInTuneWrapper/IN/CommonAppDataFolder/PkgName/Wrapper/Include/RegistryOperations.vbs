'****************************************************************************************************************************************************
' FUNCTIONS: 
'	fnSimpleKeyExist(strRegKey)
'	fnCheckIfKeyExist(strRegKey) 
'	fnRemoveKey(strRegKey) 
'	fnRenameKey(strRegKey, strRegKeyNewName) 
'	fnReadKey(strRegKey)
'	fnReadKeyType(strRegKey)
'	fnBackupKey(strRegKeyName, strBackupRegFile) 
'	fnImportReg(strRegFile, strVbsDirectory) 
'	fnChangeKey(strRegKey, strRegKeyNewValue) 
'	fnAddKey(strRegKey, strRegKeyValue, sType) ' - sType cause Type cannot be parameter
'   fnAddKeyWithPath(KeyHive, strRegKey, strRegValueName, strRegKeyValue, sType)
'	fnAddActiveSetup(strLocalPKGName, strASFriendlyName, strFunctionToRun, boolDateTime)
'	fnCheckActiveSetup(strLocalPKGName)
'	fnRemoveActiveSetup(strRegKeyName) 
'	fnARPSwap(strGUID)
'	fnARPHide(strGUID)
'	fnARPNoRep(strGUID)
'	fnARPNoMod(strGUID)
'	fnARPNoRemove(strGUID)
'	fnARPDisableOptions(strGUID)
'	fnARPDisablePointedOption(strRegKey, strOptionName)
'	fnARPClearCustoms(strRegKey)
'   fnRemoveKeyAndSubkeys(KeyHive, strRegKey)
' 	fnSubstituteARP(strGUID)
' 	fnAddDummyARP()
' 	fnRemoveDummyARP()
'	fnAppBitness()
'	
' FUNCTION EXAMPLE USAGE (IN and OUT):
' 
'fnReadKey("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{1D8E6291-B0D5-35EC-8441-6616F567A0F7}\DisplayVersion")
'	Return "10.0.40219"
'
'fnReadKeyType(strRegKey)
'	Return 1 - fail
'			REG_SZ, REG_EXPAND_SZ, REG_BINARY, REG_DWORD, REG_MULTI_SZ
'
'fnRemoveKeyAndSubkeys(KeyHive, strRegKey)
'	fnRemoveKeyAndSubkeys("HKLM", "SOFTWARE\Classes\PAACE")
'	Return: 0 - success, 1 - fail
'***************************************************************************************************************************************************
'! Will return 1 when specified registry strRegKey existing.
'!
'! @param  strRegKey            Registry strRegKey, we need to test.
'!
'! @return 0,1
Function fnSimpleKeyExist(strRegKey)	
    Dim strCurrentFNName                   : strCurrentFNName = "fnSimpleKeyExist"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKey - " & strRegKey & DotFN(0))
		
		If FnParamChecker(strRegKey, "strRegKey") = 1 Then
            fnAddLog(VarsFNErr(0) & "strRegKey" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnSimpleKeyExist = 1
		End If
		
    On Error Resume Next
	Err.Clear
	
	objShell.RegRead(strRegKey)
	
    If Err = 0 Then 
		fnAddLog("INFO: strRegKey found: " & strRegKey & ".")
		fnAddLog("FUNCTION INFO: ending fnSimpleKeyExist - success.")
		
		strLogTAB = strLogTABBak
		fnSimpleKeyExist = 0
	Else 
		fnAddLog("INFO: strRegKey not found: " & strRegKey & ".")
		fnAddLog("FUNCTION INFO: ending fnSimpleKeyExist - fail.") 
	
		strLogTAB = strLogTABBak
		fnSimpleKeyExist = 1
	End If
End Function
'-------------------------------------------------------------------------------------------
'! Will return 1 when specified registry strRegKey existing.
'!
'! @param  strRegKey            Registry strRegKey, we need to test.
'!
'! @return 0,1
Function fnCheckIfKeyExist(strRegKey)		' If strRegKey has QWORD value - function cannot find it
									' ! If existence of strRegKey (not value) is checking, on end  of 'strRegKey' must be '\'								
    Dim strCurrentFNName                   : strCurrentFNName = "fnCheckIfKeyExist"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKey - " & strRegKey & DotFN(0))
		
		If FnParamChecker(strRegKey, "strRegKey") = 1 Then
            fnAddLog(VarsFNErr(0) & "strRegKey" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnCheckIfKeyExist = 1
		End If
		
		On Error Resume Next
			objShell.RegRead(strRegKey)
		
		If Err = 0 Then
			intRVal = 0
			fnAddLog("INFO: strRegKey found: " & strRegKey & ".")
		Else
			fnAddLog("INFO: strRegKey not found: " & strRegKey & ".")
		End If
		
		On Error GoTo 0
		
		If intRVal <> 0 Then  ' FIXED by MIREK WDOWIAK
			fnAddLog("FUNCTION INFO: ending fnCheckIfKeyExist - fail.") 
			
			strLogTAB = strLogTABBak
			fnCheckIfKeyExist = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnCheckIfKeyExist - success.") 
			
			strLogTAB = strLogTABBak
			fnCheckIfKeyExist = intRVal
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will remove specified registry strRegKey or value.
'!
'! @param  strRegKey            Registry strRegKey, we need to remove.
'!
'! @return 0,1
Function fnRemoveKey(strRegKey)
    Dim strCurrentFNName                   : strCurrentFNName = "fnRemoveKey"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKey - " & strRegKey & DotFN(0))
		
		If FnParamChecker(strRegKey, "strRegKey") = 1 Then
            fnAddLog(VarsFNErr(0) & "strRegKey" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRemoveKey = 1
		End If
		
		On Error Resume Next
			wshShell.RegDelete(strRegKey)
		
		If Err = 0 Then
			fnAddLog("INFO: Deleting strRegKey: " & strRegKey & ".")
			intRVal = 0			' regdelete success
		End If
		
		
		Err.Clear
			wshShell.RegRead(strRegKey)
		If Err <> 0 Then		' strRegKey not exist
			If intRVal = 0 Then	' strRegKey not exist and regdelete success
				fnAddLog("INFO: strRegKey has deleted successfully.")
			Else				' strRegKey not exist and regdelete fail
				fnAddLog("INFO/!WARNING: Cannot delete already not existent strRegKey: " & strRegKey & ".")
				intRVal = 0
			End If
		Else					' strRegKey still exists
			fnAddLog("FUNCTION ERROR: Cannot delete strRegKey: " & strRegKey & ".")
			intRVal = 1
		End If
		On Error GoTo 0
		
		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnRemoveKey - fail.") 
			
			strLogTAB = strLogTABBak
			fnRemoveKey = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnRemoveKey - success.") 
			
			strLogTAB = strLogTABBak
			fnRemoveKey = intRVal
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will rename specified registry strRegKey.
'!
'! @param  strRegKey            Original registry strRegKey Name.
'! @param  strRegKeyNewName     New registry strRegKey Name.
'!
'! @return 0,1
Function fnRenameKey(strRegKey, strRegKeyNewName)		' fnCheckIfKeyExist, fnAddKey, fnRemoveKey functions are required	' strRegKey - full path with strRegKeyName, strRegKeyNewName - only Name of strRegKey
	Dim strValue, Root, ROOTx, KeyOldName, arrSubKeys, arrSubTypes
    Dim strCurrentFNName                   : strCurrentFNName = "fnRenameKey"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKey - " & strRegKey & ", strRegKeyNewName - " & strRegKeyNewName & DotFN(0))
		
		If FnParamChecker(strRegKey, "strRegKey") = 1 Then
            fnAddLog(VarsFNErr(0) & "strRegKey" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRenameKey = 1
		End If
    
		If FnParamChecker(strRegKeyNewName, "strRegKeyNewName") = 1 Then
            fnAddLog(VarsFNErr(0) & "strRegKeyNewName" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRenameKey = 1
		End If
		
		If fnCheckIfKeyExist(strRegKey) = 0 Then
			strValue = wshShell.RegRead(strRegKey)
		End If
		
		If Right(strRegKey,1) = "\" Then
			strRegKey = Left(strRegKey,Len(strRegKey)-1)
		End If
		
		KeyOldName = Mid(strRegKey,InStrRev(strRegKey,"\")+1)	'phrase after the intLast "\"
		strRegKey = Left(strRegKey,InStrRev(strRegKey,"\"))
		
		Root = Left(strRegKey,InStr(strRegKey,"\"))		' ex. HKEY_LOCAL_MACHINE\
		strRegKey = Mid(strRegKey,InStr(strRegKey,"\")+1)		' ex. SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\

		Select Case Root
			Case "HKEY_CLASSES_ROOT\"   : ROOTx = &H80000000
			Case "HKCR\"   				: ROOTx = &H80000000
			Case "HKEY_CURRENT_USER\"   : ROOTx = &H80000001
			Case "HKCU\"   				: ROOTx = &H80000001
			Case "HKEY_LOCAL_MACHINE\"  : ROOTx = &H80000002
			Case "HKLM\"  				: ROOTx = &H80000002
			Case "HKEY_USERS\"          : ROOTx = &H80000003
			Case "HKUS\"          		: ROOTx = &H80000003
			Case "HKEY_CURRENT_CONFIG\" : ROOTx = &H80000004
			Case "HKCC"				    : ROOTx = &H80000004
			Case "HKEY_DYN_DATA\"       : ROOTx = &H80000005
			
			Case Else: fnAddLog("FUNCTION ERROR: Invalid root of registry strRegKey! End function fnRenameKey. End code: " & intRVal & ".")
						fnFinalize(intRVal)
		End Select

		Dim i, j
		If objReg.EnumValues(ROOTx, strRegKey, arrSubKeys, arrSubTypes) = 0 AND IsArray(arrSubKeys) Then
			For i = 0 To UBound(arrSubKeys)
				If arrSubKeys(i) = KeyOldName Then
					j = i
					intRVal = 0
					Exit For
				End If
			Next
		Else
			fnAddLog("FUNCTION ERROR: Registry strRegKey path not exist or there is no value of strRegKey.")
		End If
		
		If intRVal = 0 Then
			Select Case arrSubTypes(j)
				Case 1		: arrSubTypes(j) = "REG_SZ"
				Case 2		: arrSubTypes(j) = "REG_EXPAND_SZ"
				Case 3		: arrSubTypes(j) = "REG_BINARY"
				Case 4		: arrSubTypes(j) = "REG_DWORD"
				Case 7		: arrSubTypes(j) = "REG_MULTI_SZ"
				Case Else: fnAddLog("FUNCTION INFO: Invalid reg type. End function fnRenameKey. End code 1.")
							fnFinalize(1)
			End Select
			intRVal = fnAddKey(Root & strRegKey & strRegKeyNewName, strValue, arrSubTypes(j))
			If intRVal = 0 Then
				intRVal = fnRemoveKey(Root & strRegKey & KeyOldName)
			End If
		End If
		
		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnRenameKey - fail.") 
			
			strLogTAB = strLogTABBak
			fnRenameKey = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnRenameKey - success.") 
			
			strLogTAB = strLogTABBak
			fnRenameKey = intRVal
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will add new registry strRegKey, which may contain path to the directory in the entry Name.
'!
'! @param  KeyHive			Registry strRegKey main hive (short Name of full Name, case insensitive, not working for "HKEY_USERS").
'! @param  strRegKey				Registry strRegKey Name.
'! @param  strRegValueName		Name of the registry entry (containing path to the directory).
'! @param  strRegKeyValue			strRegKeyValue for specified registry strRegKey (for "REG_DWORD" type only numerical values are allowed).
'! @param  sType			Data type for specified registry strRegKey (only "REG_SZ" and "REG_DWORD" types are allowed, case insensitive).
'!
'! @return 0,1
Function fnAddKeyWithPath(KeyHive, strRegKey, strRegValueName, strRegKeyValue, sType)
	Dim KeyHiveValue, KeyPath
	Dim strCurrentFNName					: strCurrentFNName = "fnAddKeyWithPath"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "KeyHive" & DotFN(3) & KeyHive & ", strRegKey" & DotFN(3) & strRegKey & ", strRegValueName" & DotFN(3) & strRegValueName & ", strRegKeyValue" & DotFN(3) & strRegKeyValue & ", sType" & DotFN(3) & sType & DotFN(0))
		
	'----- Check correctness of KeyHive parameter and if proper, then convert it into value:
		If FnParamChecker(KeyHive, "KeyHive") = 1 Then
			fnAddLog(VarsFNErr(0) & "KeyHive" & VarsFNErr(1))
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddKeyWithPath = 1
		Else
			KeyHive = UCase(KeyHive)	'Convert KeyHive value to uppercase
			
			If KeyHive = "HKCR" OR KeyHive = "HKEY_CLASSES_ROOT" Then
				KeyHiveValue = &H80000000
			ElseIf KeyHive = "HKCU" OR KeyHive = "HKEY_CURRENT_USER" Then
				KeyHiveValue = &H80000001
			ElseIf KeyHive = "HKLM" OR KeyHive = "HKEY_LOCAL_MACHINE" Then
				KeyHiveValue = &H80000002
			ElseIf KeyHive = "HKCC" OR KeyHive = "HKEY_CURRENT_CONFIG" Then
				KeyHiveValue = &H80000005
			Else
				fnAddLog(VarsFNErr(0) & "KeyHive value" & DotFN(3) & KeyHive & " is out of range.")
				fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
				
				strLogTAB = strLogTABBak
				fnFinalize(1)
				fnAddKeyWithPath = 1
			End If
		End If
		
		
		If FnParamChecker(strRegKey, "strRegKey") = 1 Then
			fnAddLog(VarsFNErr(0) & "strRegKey" & VarsFNErr(1))
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddKeyWithPath = 1
		End If
		
		If FnParamChecker(strRegValueName, "strRegValueName") = 1 Then
			fnAddLog(VarsFNErr(0) & "strRegValueName" & VarsFNErr(1))
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddKeyWithPath = 1
		End If
		
		If FnParamChecker(strRegKeyValue, "strRegKeyValue") = 1 Then
			fnAddLog(VarsFNErr(0) & "strRegKeyValue" & VarsFNErr(1))
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddKeyWithPath = 1
		End If
		
		
	'----- Check correctness of sType parameter:
		If FnParamChecker(sType, "sType") = 1 Then
			fnAddLog(VarsFNWar(0) & "sType" & VarsFNWar(1) & "REG_SZ" & DotFN(0))
			sType = "REG_SZ"
		Else
			sType = UCase(sType)		'Convert sType value to uppercase
			
			If sType <> "REG_SZ" AND sType <> "REG_DWORD" Then
				fnAddLog(VarsFNErr(0) & "sType value" & DotFN(3) & sType & " is out of range.")
				fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
				
				strLogTAB = strLogTABBak
				fnFinalize(1)
				fnAddKeyWithPath = 1
			End If
		End If
		
		
	'----- Create registry strRegKey (the strRegKey must exist for SetStringValue and SetDWORDValue methods):
		KeyPath = KeyHive & Chr(92) & strRegKey & Chr(92)
		
		If fnSimpleKeyExist(KeyPath) = 1 Then
			fnAddLog(PreTxtFN(0) & "Registry strRegKey" & DotFN(3) & Chr(34) & KeyPath & Chr(34) & " not found. Proceed with creation.")
			objReg.CreateKey KeyHiveValue, strRegKey
			
			If fnSimpleKeyExist(KeyPath) = 0 Then
				fnAddLog(PreTxtFN(0) & "Added registry strRegKey" & DotFN(3) & Chr(34) & KeyPath & Chr(34) & DotFN(0))
			Else
				fnAddLog(PreTxtFN(1) & "Cannot create registry strRegKey" & DotFN(3) & Chr(34) & KeyPath & Chr(34) & DotFN(0))
				fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
				
				strLogTAB = strLogTABBak
				fnFinalize(1)
				fnAddKeyWithPath = 1
			End If
		Else
			fnAddLog(PreTxtFN(0) & "Registry strRegKey" & DotFN(3) & Chr(34) & KeyPath & Chr(34) & CreationFN(2))
		End If
		
		
	'----- Create registry entry:
		On Error Resume Next
		Err.Clear
		
			If sType = "REG_SZ" Then
				objReg.SetStringValue	KeyHiveValue, strRegKey, strRegValueName, strRegKeyValue
			ElseIf sType = "REG_DWORD" Then
				objReg.SetDWORDValue	KeyHiveValue, strRegKey, strRegValueName, strRegKeyValue
			End If
			
			If Err = 0 Then
				fnAddLog(PreTxtFN(0) & "Added registry entry with the following data: Name" & DotFN(3) & strRegValueName & ", Type" & DotFN(3) & sType & ", strRegKeyValue" & DotFN(3) & strRegKeyValue & DotFN(0))
				fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0))
				
				strLogTAB = strLogTABBak
				fnAddKeyWithPath = 0
			Else
				fnAddLog(PreTxtFN(0) & "Cannot create registry entry with the following data: Name" & DotFN(3) & strRegValueName & ", Type" & DotFN(3) & sType & ", strRegKeyValue" & DotFN(3) & strRegKeyValue & DotFN(0))
				fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
				
				strLogTAB = strLogTABBak
				fnFinalize(1)
				fnAddKeyWithPath = 1
			End If
		
		On Error GoTo 0
End Function
'-------------------------------------------------------------------------------------------
'! Will return specified registry strRegKey value.
'!
'! @param  strRegKey            Registry strRegKey Name.
'!
'! @return strRegKey value
Function fnReadKey(strRegKey) 'Andrzej Ozga // POPRAWIONE M WDOWIAK
	Dim KeyValue
    Dim strCurrentFNName                   : strCurrentFNName = "fnReadKey"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKey - " & strRegKey & DotFN(0))
		
		If FnParamChecker(strRegKey, "strRegKey") = 1 Then
			fnAddLog("FUNCTION ERROR: variable strRegKey cannot be NULL.")
			fnAddLog("FUNCTION INFO: ending fnReadKey - fail.") 
			
			strLogTAB = strLogTABBak
			fnReadKey = 1
		End If
		
		
		If fnCheckIfKeyExist(strRegKey) = 0 Then 
			KeyValue = objShell.RegRead(strRegKey)
			fnAddLog("FUNCTION INFO: " & strRegKey & " value is - " & KeyValue & ".")
		Else
			fnAddLog("FUNCTION INFO: strRegKey: " & strRegKey & " don't exist. Nothing to read.") 'Simple condition to know that we don't have value, by G.G.
			KeyValue = NULL
		End If
		
		
		strLogTAB = strLogTABBak
		fnReadKey = KeyValue
End Function
'-------------------------------------------------------------------------------------------
'! Will return specified registry strRegKey data type.
'!
'! @param  strRegKey            Registry strRegKey Name.
'!
'! @return strRegKey data type
Function fnReadKeyType(strRegKey)
	Dim KeyValue, strRegValueName, ValueType, keyType, KeyP, x, i
    Dim strCurrentFNName                   : strCurrentFNName = "fnReadKeyType"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKey - " & strRegKey & DotFN(0))
		
		If FnParamChecker(strRegKey, "strRegKey") = 1 Then
			fnAddLog("FUNCTION ERROR: variable strRegKey cannot be NULL.")
			fnAddLog("FUNCTION INFO: ending fnReadKeyType - fail.") 
			
			strLogTAB = strLogTABBak
			fnReadKeyType = 1
		End If
			
	const HKEY_LOCAL_MACHINE = &H80000002
	const REG_SZ = 1
	const REG_EXPAND_SZ = 2
	const REG_BINARY = 3
	const REG_DWORD = 4
	const REG_MULTI_SZ = 7
	
		If (fnCheckIfKeyExist(strRegKey)) Then
			objReg.EnumValues HKEY_LOCAL_MACHINE, KeyP, strRegValueName, ValueType
			
			For i=0 To UBound(strRegValueName)
				fnAddLog("FUNCTION INFO: strRegKeyValue Name: " & strRegValueName(i) & ".")
				
				Select Case ValueType(i)
					Case REG_SZ
						keyType = "REG_SZ"
					Case REG_EXPAND_SZ
						keyType = "REG_EXPAND_SZ"
					Case REG_BINARY
						keyType = "REG_BINARY"
					Case REG_DWORD
						keyType = "REG_DWORD"
					Case REG_MULTI_SZ
						keyType = "REG_MULTI_SZ"
				End Select 
				
				fnAddLog("FUNCTION INFO: fnReadKeyType for " & strRegKey & " type is - " & keyType & ".")
			Next
		Else
			fnAddLog("FUNCTION ERROR: variable strRegKey doesn't exist.")
			fnAddLog("FUNCTION INFO: ending fnReadKeyType - fail.") 
			
			strLogTAB = strLogTABBak
			fnReadKeyType = 1
		End If
	
	fnAddLog("FUNCTION INFO: ending fnReadKeyType.") 
	
	fnReadKeyType = keyType
	strLogTAB = strLogTABBak
End Function
'-------------------------------------------------------------------------------------------
'! Will copy and change specified registry strRegKey Name.
'!
'! @param  strRegKey            Original registry strRegKey Name.
'! @param  strRegKeyNewName     Backup registry strRegKey Name.
'!
'! @return 0,1
Function fnBackupKey(strRegKeyName, strBackupRegFile)	 ' strRegKeyName = path and strRegKey/value, strBackupRegFile = path with Name of file with extension (recommended: .reg)
    Dim strCurrentFNName                   : strCurrentFNName = "fnBackupKey"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKeyName - " & strRegKeyName & ", strBackupRegFile - " & strBackupRegFile & DotFN(0))
		
		If FnParamChecker(strRegKeyName, "strRegKeyName") = 1 Then
            fnAddLog(VarsFNErr(0) & "strRegKeyName" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnBackupKey = 1
		End If
		
		If FnParamChecker(strBackupRegFile, "strBackupRegFile") = 1 Then
            fnAddLog(VarsFNErr(0) & "strBackupRegFile" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnBackupKey = 1
		End If
		
		intRVal = fnCheckIfKeyExist(strRegKeyName)
		
		If intRVal = 0 Then
			Dim strCommand
			strCommand = "reg copy " & strRegKeyName & " " & strBackupRegFile & " /s /f"  'regedit /e
			
			intRVal = objShell.Run(strCommand, 0, TRUE)
			If intRVal <> 0 Then
				fnAddLog("INFO: Error returned from exporting registry: " & intRVal)
			Else
				fnAddLog("INFO: No errors returned from exporting the registry file") 	' intRVal = 0
			End If
		End If
		
		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnBackupKey - fail.") 
			
			strLogTAB = strLogTABBak
			fnBackupKey = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnBackupKey - success.") 
			
			strLogTAB = strLogTABBak
			fnBackupKey = intRVal
			Exit Function
		End If
 End Function
'-------------------------------------------------------------------------------------------
'! Will import .reg file into registry
'!
'! @param  strRegFile        .reg file Name.
'! @param  strVbsDirectory      strVbsDirectory when .reg file can be found.
'!
'! @return 0,1
Function fnImportReg(strRegFile, strVbsDirectory)	'NO QA!!!!	' strVbsDirectory = path to file ex. C:\<MyFolder>\<MyFile>, strRegFile = regfile Name
    Dim strCurrentFNName                   : strCurrentFNName = "fnImportReg"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegFile - " & strRegFile & ", strVbsDirectory - " & strVbsDirectory & DotFN(0))
		
		If FnParamChecker(strRegFile, "strRegFile") = 1 Then
            fnAddLog(VarsFNErr(0) & "strRegFile" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnImportReg = 1
		End If
		
		If FnParamChecker(strVbsDirectory, "strVbsDirectory") = 1 Then
            fnAddLog(VarsFNErr(0) & "strVbsDirectory" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnImportReg = 1
		End If
		
		If right(strVbsDirectory, 1) <> "\" Then
			strVbsDirectory = strVbsDirectory & "\"
		End If
		
		On Error Resume Next
		fnAddLog("Run command: Regedit.exe /s " & Chr(34) & strVbsDirectory & strRegFile & Chr(34) & ", 0, True")
		intRVal = objShell.Run("Regedit.exe /s " & Chr(34) & strVbsDirectory & strRegFile & Chr(34), 0, True)
		If Err <> 0 OR intRVal <> 0 Then
			Err.Clear
			fnAddLog "Run command: Cmd.exe /K REG.arrEXE IMPORT " & Chr(34) & strVbsDirectory & strRegFile & Chr(34), 0, True
			intRVal = objShell.Run("Cmd.exe /K REG.arrEXE IMPORT " & Chr(34) & strVbsDirectory & strRegFile & Chr(34) & ", 0, True")
		End If
		If Err <> 0 OR intRVal <> 0 Then
			fnAddLog("FUNCTION ERROR: Import reg failed.")
		Else
			intRVal = 0
			fnAddLog("INFO: Regfile imported successfully.")
		End If
		
		On Error GoTo 0
		
		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnImportReg - fail.") 
			
			strLogTAB = strLogTABBak
			fnImportReg = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnImportReg - success.") 
			
			strLogTAB = strLogTABBak
			fnImportReg = intRVal
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will change registry strRegKey value
'!
'! @param  strRegKey            Original registry strRegKey Name.
'! @param  strRegKeyNewValue    New value for specified registry strRegKey..
'!
'! @return 0,1
Function fnChangeKey(strRegKey, strRegKeyNewValue) ' strRegKey - full path with strRegKeyName, strRegKeyNewName - only value of strRegKey ' The type of strRegKey remains the same 
										' strRegKey cannot be a QWORD value cause error fnCheckIfKeyExist function
    Dim strCurrentFNName                   : strCurrentFNName = "fnChangeKey"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKey - " & strRegKey & ", strRegKeyNewValue - " & strRegKeyNewValue & DotFN(0))
		
		If FnParamChecker(strRegKey, "strRegKey") = 1 Then
            fnAddLog(VarsFNErr(0) & "strRegKey" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnChangeKey = 1
		End If
		
		If FnParamChecker(strRegKeyNewValue, "strRegKeyNewValue") = 1 Then
            fnAddLog(VarsFNErr(0) & "strRegKeyNewValue" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnChangeKey = 1
		End If
		
		Dim strValue, Root, ROOTx, KeyOldName, arrSubKeys, arrSubTypes
		If fnCheckIfKeyExist(strRegKey) = 0 Then
			strValue = wshShell.RegRead(strRegKey)
			fnAddLog("INFO: strRegKey old value: " & strValue)
		End If
		If Right(strRegKey,1) = "\" Then
			strRegKey = Left(strRegKey,Len(strRegKey)-1)
		End If
		KeyOldName = Mid(strRegKey,InStrRev(strRegKey,"\")+1)	'phrase after the intLast "\"
		strRegKey = Left(strRegKey,InStrRev(strRegKey,"\"))
		
		Root = Left(strRegKey,InStr(strRegKey,"\"))		' ex. HKEY_LOCAL_MACHINE\
		strRegKey = Mid(strRegKey,InStr(strRegKey,"\")+1)		' ex. SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\

		Select Case Root
			Case "HKEY_CLASSES_ROOT\"   : ROOTx = &H80000000
			Case "HKCR\"   				: ROOTx = &H80000000
			Case "HKEY_CURRENT_USER\"   : ROOTx = &H80000001
			Case "HKCU\"   				: ROOTx = &H80000001
			Case "HKEY_LOCAL_MACHINE\"  : ROOTx = &H80000002
			Case "HKLM\"  				: ROOTx = &H80000002
			Case "HKEY_USERS\"          : ROOTx = &H80000003
			Case "HKUS\"          		: ROOTx = &H80000003
			Case "HKEY_CURRENT_CONFIG\" : ROOTx = &H80000004
			Case "HKCC"				    : ROOTx = &H80000004
			Case "HKEY_DYN_DATA\"       : ROOTx = &H80000005
			
			Case Else: fnAddLog("FUNCTION ERROR: Invalid root of registry strRegKey! End function fnChangeKey. End code: " & intRVal) 
						fnFinalize(intRVal)
		End Select

		Dim i, j
		If objReg.EnumValues(ROOTx, strRegKey, arrSubKeys, arrSubTypes) = 0 AND IsArray(arrSubKeys) Then
			For i = 0 To UBound(arrSubKeys)
				If arrSubKeys(i) = KeyOldName Then
					j = i
					intRVal = 0
					Exit For
				End If
			Next
		Else
			fnAddLog("FUNCTION ERROR: Registry strRegKey path not exist or there is no value of strRegKey!")
		End If
		If intRVal = 0 Then
			Select Case arrSubTypes(j)
				Case 1		: arrSubTypes(j) = "REG_SZ"
				Case 2		: arrSubTypes(j) = "REG_EXPAND_SZ"
				Case 3		: arrSubTypes(j) = "REG_BINARY"
				Case 4		: arrSubTypes(j) = "REG_DWORD"
				Case 7		: arrSubTypes(j) = "REG_MULTI_SZ"
				Case Else: fnAddLog("FUNCTION ERROR: Invalid reg type! End function fnChangeKey. End code: 1")
							fnFinalize(1)
			End Select
			intRVal = fnAddKey(Root & strRegKey & KeyOldName, strRegKeyNewValue, arrSubTypes(j))
		End If
		
		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnChangeKey - fail.") 
			
			strLogTAB = strLogTABBak
			fnChangeKey = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnChangeKey - success.") 
			
			strLogTAB = strLogTABBak
			fnChangeKey = intRVal
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will add new registry strRegKey.
'!
'! @param  strRegKey            Original registry strRegKey Name.
'! @param  strRegKeyValue    	  strRegKeyValue for specified registry strRegKey.
'! @param  sType    	  Data type for specified registry strRegKey.
'!
'! @return 0,1
Function fnAddKey(strRegKey, strRegKeyValue, sType) 	' sType [Optional]= String: "REG_SZ"(Default), "REG_EXPAND_SZ", Integer/String: "REG_DWORD", Array: "REG_BINARY", "REG_MULTI_SZ"
									' optional so if not need = NULL
									' if strRegKey already exists - changing strRegKey (value and type) and ending with success
									' comment: may function should check if strRegKey already exists?
    Dim strCurrentFNName                   : strCurrentFNName = "fnAddKey"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKey - " & strRegKey & ", strRegKeyValue - " & strRegKeyValue & ", sType - " & sType & DotFN(0))
		
		If FnParamChecker(strRegKey, "strRegKey") = 1 Then
            fnAddLog(VarsFNErr(0) & "strRegKey" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddKey = 1
		End If
		
		If FnParamChecker(strRegKeyValue, "strRegKeyValue") = 1 Then
            fnAddLog(VarsFNErr(0) & "strRegKeyValue" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddKey = 1
		End If
		
		' sType is optional!
		
		On Error Resume Next
		If sType <> "" Then			' optional parameter checking
			objShell.RegWrite strRegKey, strRegKeyValue, sType
		Else
			objShell.RegWrite strRegKey, strRegKeyValue
			sType = "REG_SZ"
		End If
		
		objShell.RegRead strRegKey
		If Err = 0 Then
			fnAddLog("INFO: Added registry strRegKey: " & chr(34) & strRegKey & chr(34) & ", strRegKeyValue: " & strRegKeyValue & ", Type: " & sType)
			intRVal = 0
			Else
			fnAddLog("FUNCTION ERROR: Cannot write registry: " & chr(34) & strRegKey & chr(34) & ", strRegKeyValue: " & strRegKeyValue & ", Type: " & sType)
		End If
		On Error GoTo 0

		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnAddKey - fail.") 
			
			strLogTAB = strLogTABBak
			fnAddKey = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnAddKey - success.") 
			
			strLogTAB = strLogTABBak
			fnAddKey = intRVal
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will add new active setup entry with custom function to run.
'!
'! @param  strLocalPKGName       Package Name accorging naming convention.
'! @param  strASFriendlyName   Friendly Name for custom action.
'! @param  strFunctionToRun  Function/file/script that will be run when active setup will be triggered.
'! @param  boolDateTime    If data stamp will be used in active setup Name.
'!
'! @return 0,1
Function fnAddActiveSetup(strLocalPKGName, strASFriendlyName, strFunctionToRun, boolDateTime)
    Dim pathAS, strDateTime
    Dim strCurrentFNName                   : strCurrentFNName = "fnAddActiveSetup"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strLocalPKGName - " & strLocalPKGName & ", strASFriendlyName - " & strASFriendlyName & ", strFunctionToRun - " & strFunctionToRun & DotFN(0))
		
		If FnParamChecker(strLocalPKGName, "strLocalPKGName") = 1 Then
            fnAddLog(VarsFNErr(0) & "strLocalPKGName" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddActiveSetup = 1
		End If
		
		If FnParamChecker(strASFriendlyName, "strASFriendlyName") = 1 Then
            fnAddLog(VarsFNErr(0) & "strASFriendlyName" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))

			strLogTAB = strLogTABBak			
			fnFinalize(1)
			fnAddActiveSetup = 1
		End If
		
		If FnParamChecker(strFunctionToRun, "strFunctionToRun") = 1 Then
            fnAddLog(VarsFNErr(0) & "strFunctionToRun" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddActiveSetup = 1
		End If
		
		If FnParamChecker(boolDateTime, "boolDateTime") = 1 Then
            fnAddLog(VarsFNErr(0) & "boolDateTime" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddActiveSetup = 1
		End If
		
        if boolDateTime = false Then
            pathAS = "HKEY_LOCAL_MACHINE\Software\Microsoft\Active Setup\Installed Components\" & strLocalPKGName & "\"
        Else
            strDateTime = fnGetDateAndTime()
            pathAS = "HKEY_LOCAL_MACHINE\Software\Microsoft\Active Setup\Installed Components\" & strLocalPKGName & "_" & strDateTime & "\"
        End If
        
		intRVal = fnAddKey(pathAS, strASFriendlyName, NULL)
		If intRVal = 0 Then
			intRVal = fnAddKey(pathAS & "StubPath", strFunctionToRun, NULL)
			fnAddLog("INFO: Adding AS")
		End If

		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnAddActiveSetup - fail.") 
			
			strLogTAB = strLogTABBak
			fnAddActiveSetup = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnAddActiveSetup - success.") 
			
			strLogTAB = strLogTABBak
			fnAddActiveSetup = intRVal
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will check if active setup exist.
'!
'! @param  strLocalPKGName       Package Name accorging naming convention.
'!
'! @return 0,1
Function fnCheckActiveSetup(strLocalPKGName)		'only 32 bit     'if vendor's AS strLocalPKGName is strGUID
	Dim pathAS
    Dim strCurrentFNName                   : strCurrentFNName = "fnCheckActiveSetup"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strLocalPKGName - " & strLocalPKGName & DotFN(0))
		
		If FnParamChecker(strLocalPKGName, "strLocalPKGName") = 1 Then
            fnAddLog(VarsFNErr(0) & "strLocalPKGName" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnCheckActiveSetup = 1
		End If
		
		pathAS = "HKEY_LOCAL_MACHINE\Software\Microsoft\Active Setup\Installed Components\" & strLocalPKGName & "\StubPath"
		intRVal = fnCheckIfKeyExist(pathAS)
		If intRVal = 0 Then
			fnAddLog("INFO: AS for " & strLocalPKGName & " finded. Run command (stubKey): " & wshShell.RegRead(pathAS))
		Else
			fnAddLog("FUNCTION INFO/!WARNING: AS for " & strLocalPKGName & " not finded!")
		End If
		
		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnCheckActiveSetup - fail.") 
			
			strLogTAB = strLogTABBak
			fnCheckActiveSetup = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnCheckActiveSetup - success.") 
			
			strLogTAB = strLogTABBak
			fnCheckActiveSetup = intRVal
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will remove specified active setup.
'!
'! @param  strRegKeyName       Name of active setup to remove.
'!
'! @return 0,1
Function fnRemoveActiveSetup(strRegKeyName)		'strRegKeyName can be strGUID or part strGUID/GUIDname (solution if app has different types endings/beginnings) 'Note: This function fnRemoveActiveSetup 32bit app and 64bit app
	Const HKEY_LOCAL_MACHINE =	&H80000002	
	Dim strSubKey, subKey64, arrSubKeys, arrSubKeys32, arrSubKeys64, subKey32
    Dim strCurrentFNName                   : strCurrentFNName = "fnRemoveActiveSetup"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKeyName - " & strRegKeyName & DotFN(0))
		
		If FnParamChecker(strRegKeyName, "strRegKeyName") = 1 Then
            fnAddLog(VarsFNErr(0) & "strRegKeyName" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRemoveActiveSetup = 1
		End If
		
		If fnIs64BitOS(True) = 0 Then
			objReg.EnumKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Active Setup\Installed Components", arrSubKeys32
			objReg.EnumKey HKEY_LOCAL_MACHINE, "Software\Wow6432Node\Microsoft\Active Setup\Installed Components", arrSubKeys64
			arrSubKeys = joinArray(arrSubKeys32, arrSubKeys64)
			
			For Each strSubKey In arrSubKeys
				If InStr(strSubKey,Keyname) Then	'by AA, KS, MG
					subKey32 = "HKLM\Software\Microsoft\Active Setup\Installed Components\" & strSubKey & "\"
					subKey64 = "HKLM\Software\Wow6432Node\Microsoft\Active Setup\Installed Components\" & strSubKey & "\"
					On Error Resume Next
					wshShell.RegDelete subKey32
					fnAddLog("INFO: Trying to delete 64bit AS.")
					If Err <> 0 Then
						Err.Clear
						wshShell.RegDelete subKey64
						fnAddLog("INFO: Trying to delete 32bit AS.")
					End If
					If Err = 0 then
						fnAddLog("INFO: Active setup registry " & strSubKey & " was removed.")
						intRVal = 0
					Else
						fnAddLog("FUNCTION ERROR: Removing active setup registry " & strSubKey & " failed.") 'intRVal = 1
					End If
					On Error GoTo 0
				Else
					fnAddLog("INFO/!WARNING: Active setup for this strGUID not found. The path to AS can be wrong or AS has been deleted already.")
					intRVal = 0
				End If
			Next

		Else
			objReg.EnumKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Active Setup\Installed Components", arrSubKeys
			
			For Each strSubKey In arrSubKeys
				If InStr(strSubKey,Keyname) Then	'by AA, KS, MG
					subKey32 = "HKLM\Software\Microsoft\Active Setup\Installed Components\" & strSubKey & "\"
							
					On Error Resume Next
					wshShell.RegDelete subKey32
					fnAddLog("INFO: Trying to delete Active Setup registry strRegKey.")
					If Err = 0 then
						fnAddLog("INFO: Active setup registry " & strSubKey & " was removed.")
						intRVal = 0
					Else
						fnAddLog("FUNCTION ERROR: Removing active setup registry " & strSubKey & " failed.")
					End If
					On Error GoTo 0
				Else
					fnAddLog("INFO/!WARNING: Active setup for this strGUID not found. The path to AS can be wrong or AS has been deleted already.")
					intRVal = 0
				End If
			Next	
		End If		

		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnRemoveActiveSetup - fail.") 
			
			strLogTAB = strLogTABBak
			fnRemoveActiveSetup = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnRemoveActiveSetup - success.") 
			
			strLogTAB = strLogTABBak
			fnRemoveActiveSetup = intRVal
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! empty string
'!
'! @param  strGUID        			empty string
'! @param  strLocalPKGName      		empty string
'! @param  UninstallString      empty string
'!
'! @return 0,1
Function fnARPSwap(strGUID, strLocalPKGName, UninstallString)	' strGUID often exists in '{ }', in this case it is necessary to add these parentheses
													' If adding uninstallstring or hiding native strRegKey failed, function not remove previously added keys (function not undo its action)
													' Function on end its action hide nomodify button and remove system component of new strGUID if its exists.
													' Check if there is no exists the same keys 32 bit and 64 bit (not common situation)
	Dim strKey, strKey64
    Dim strCurrentFNName                   : strCurrentFNName = "fnARPSwap"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strGUID - " & strGUID & ", strLocalPKGName - " & strLocalPKGName & ", UninstallString - " & UninstallString & DotFN(0))
		
		If FnParamChecker(strGUID, "strGUID") = 1 Then
            fnAddLog(VarsFNErr(0) & "strGUID" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnARPSwap = 1
		End If
		
		If FnParamChecker(strLocalPKGName, "strLocalPKGName") = 1 Then
            fnAddLog(VarsFNErr(0) & "strLocalPKGName" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnARPSwap = 1
		End If
		
		If FnParamChecker(UninstallString, "UninstallString") = 1 Then
            fnAddLog(VarsFNErr(0) & "UninstallString" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnARPSwap = 1
		End If
		
		strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID
		strKey64 = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID
			
		On Error Resume Next
			wshShell.RegRead "HKEY_LOCAL_MACHINE\" & strKey & "\"
			fnAddLog("INFO: Trying to find 32bit UninstallString")
			If Err <> 0 Then
				Err.Clear
				wshShell.RegRead "HKEY_LOCAL_MACHINE\" & strKey64 & "\"
				fnAddLog("INFO: Trying to find 64bit UninstallString")
				If Err <> 0 Then
					fnAddLog("FUNCTION ERROR: UninstallString was not found, check spelling.") 'intRVal = 1
					fnAddLog("FUNCTION INFO: ending fnARPSwap - fail.")
					fnFinalize(1)
					fnARPSwap = 1
				Else
					intRVal = 0
					fnAddLog("INFO: UninstallString 64bit was found.")
					strKey = strKey64
				End If
			Else
				intRVal = 0
				fnAddLog("INFO: UninstallString 32bit was found.")
				' strKey = strKey
			End If
		On Error GoTo 0
			
		Dim i, arrSubKeys, arrSubTypes, Root, strSubKeyVal
		Root = &H80000002
		If objReg.EnumValues(Root, strKey, arrSubKeys, arrSubTypes) = 0 AND IsArray(arrSubKeys) Then
			For i = 0 To UBound(arrSubKeys)
				Select Case arrSubTypes(i)
					Case 1		: arrSubTypes(i) = "REG_SZ"
					Case 2		: arrSubTypes(i) = "REG_EXPAND_SZ"
					Case 3		: arrSubTypes(i) = "REG_BINARY"
					Case 4		: arrSubTypes(i) = "REG_DWORD"
					Case 7		: arrSubTypes(i) = "REG_MULTI_SZ"
					Case Else: fnAddLog("FUNCTION ERROR: Invalid reg type! End function fnChangeKey. End code: 1")
								fnFinalize(1)
				End Select
				strSubKeyVal = wshShell.RegRead("HKEY_LOCAL_MACHINE\" & strKey & "\" & arrSubKeys(i))
				intRVal = fnAddKey("HKEY_LOCAL_MACHINE\" & strKey & "_" & strLocalPKGName & "\" & arrSubKeys(i), strSubKeyVal, arrSubTypes(i))
				If intRVal <> 0 Then
					fnRemoveKey("HKEY_LOCAL_MACHINE\" & strKey & "_" & strLocalPKGName)												'	Function activity cleaning
					fnAddLog("FUNCTION ERROR: copy of strRegKey " & strKey & " error.")			' 	check, attention this step!
					fnAddLog("FUNCTION INFO: ending fnARPSwap - fail.")
					'fnFinalize(1)
					fnARPSwap = 1
				End If
			Next
		Else
			fnAddLog("FUNCTION ERROR: Registry strRegKey path not exist or there is no value of strRegKey!")
		End If
		
		If intRVal = 0 Then
			intRVal = fnAddKey("HKEY_LOCAL_MACHINE\" & strKey & "_" & strLocalPKGName & "\UninstallString", UninstallString, "REG_EXPAND_SZ")	' change uninstallstring in new copy of ARP information about package
			If intRVal = 0 Then
			intRVal = fnARPHide(strGUID)	' change visibility on hidden in old copy of ARP information about package
				If intRVal = 0 Then
					intRVal = fnARPNoMod(strGUID & "_" & strLocalPKGName)	' hidden nomodify button in new copy of strGUID in ARP information about package
					If intRVal = 0 Then
						intRVal = fnCheckIfKeyExist("HKEY_LOCAL_MACHINE\" & strKey & "_" & strLocalPKGName & "\SystemComponent")	' remove system component of new strGUID if its exists
						If intRVal = 0 Then
							fnRemoveKey("HKEY_LOCAL_MACHINE\" & strKey & "_" & strLocalPKGName & "\SystemComponent")
						End If
					End If
				End If
			End If
		End If
		
		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnARPSwap - fail.") 
			
			strLogTAB = strLogTABBak
			fnARPSwap = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnARPSwap - success.") 
			
			strLogTAB = strLogTABBak
			fnARPSwap = intRVal
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will hide ARP entry.
'!
'! @param  strGUID        			Application strGUID
'!
'! @return 0,1
Function fnARPHide(strGUID)
	Dim strRegType, strRegValue, strKeyPath, strRegKey
    Dim strCurrentFNName                   : strCurrentFNName = "fnARPHide"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strGUID - " & strGUID & DotFN(0))
		
		If FnParamChecker(strGUID, "strGUID") = 1 Then
            fnAddLog(VarsFNErr(0) & "strGUID" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnARPHide = 1
		End If
		
		strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID & "\"
		intRVal = fnCheckIfKeyExist(strKeyPath)
		
		if intRVal <> 0 Then
			strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID & "\"
			intRVal = fnCheckIfKeyExist(strKeyPath)
		End If

		If intRVal = 0 Then
			strRegKey = strKeyPath & "SystemComponent"
			strRegType = "REG_DWORD"
			strRegValue = 1
			intRVal = fnAddKey(strRegKey, strRegValue, strRegType)
			fnAddLog("INFO: strGUID is correct, ARP hiding has been proceeding")
		Else
			fnAddLog("INFO: Wrong reg path or strGUID!")
		End If
		
		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnARPHide - fail.") 
			
			strLogTAB = strLogTABBak
			fnARPHide = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnARPHide - success.") 
			
			strLogTAB = strLogTABBak
			fnARPHide = intRVal
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will remove ARP repair button in ARP entry.
'!
'! @param  strGUID        			Application strGUID
'!
'! @return 0,1
Function fnARPNoRep(strGUID)
	Dim strRegType, strRegValue, strKeyPath
    Dim strCurrentFNName                   : strCurrentFNName = "fnARPNoRep"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strGUID - " & strGUID & DotFN(0))
		
		If FnParamChecker(strGUID, "strGUID") = 1 Then
            fnAddLog(VarsFNErr(0) & "strGUID" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnARPNoRep = 1
		End If
		
		
		strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID & "\"
		intRVal = fnARPDisablePointedOption(strKeyPath, "NoRepair")
		
		if intRVal <> 0 Then
			strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID & "\"
			intRVal = fnARPDisablePointedOption(strKeyPath, "NoRepair")
		End If
		
		
		
		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnARPNoRep - fail.") 
			
			strLogTAB = strLogTABBak
			fnARPNoRep = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnARPNoRep - success.") 
			
			strLogTAB = strLogTABBak
			fnARPNoRep = intRVal
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will remove ARP modifity button in ARP entry.
'!
'! @param  strGUID        			Application strGUID
'!
'! @return 0,1
Function fnARPNoMod(strGUID)
	Dim strRegType, strRegValue, strKeyPath
    Dim strCurrentFNName                   : strCurrentFNName = "fnARPNoMod"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strGUID - " & strGUID & DotFN(0))
		
		If FnParamChecker(strGUID, "strGUID") = 1 Then
            fnAddLog(VarsFNErr(0) & "strGUID" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnARPNoMod = 1
		End If
		
		
		strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID & "\"
		intRVal = fnARPDisablePointedOption(strKeyPath, "NoModify")
		
		if intRVal <> 0 Then
			strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID & "\"
			intRVal = fnARPDisablePointedOption(strKeyPath, "NoModify")
		End If
		
		
		
		
		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnARPNoMod - fail.") 
			
			strLogTAB = strLogTABBak
			fnARPNoMod = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnARPNoMod - success.")

			strLogTAB = strLogTABBak			
			fnARPNoMod = intRVal
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will remove ARP remove/uninstall button in ARP entry.
'!
'! @param  strGUID        			Application strGUID
'!
'! @return 0,1
Function fnARPNoRemove(strGUID)
	Dim strRegType, strRegValue, strKeyPath
    Dim strCurrentFNName                   : strCurrentFNName = "fnARPNoRemove"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strGUID - " & strGUID & DotFN(0))
		
		If FnParamChecker(strGUID, "strGUID") = 1 Then
            fnAddLog(VarsFNErr(0) & "strGUID" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnARPNoRemove = 1
		End If
		
		
		strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID & "\"
		intRVal = fnARPDisablePointedOption(strKeyPath, "NoRemove")
		
		if intRVal <> 0 Then
			strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID & "\"
			intRVal = fnARPDisablePointedOption(strKeyPath, "NoRemove")
		End If
		
	
		
		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnARPNoRemove - fail.") 
			
			strLogTAB = strLogTABBak
			fnARPNoRemove = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnARPNoRemove - success.") 
			
			strLogTAB = strLogTABBak
			fnARPNoRemove = intRVal
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will remove ARP remove/uninstall, repail, modifity buttons in ARP entry.
'!
'! @param  strGUID        			Application strGUID
'!
'! @return 0,1
Function fnARPDisableOptions(strGUID)
    Dim strCurrentFNName                   : strCurrentFNName = "fnARPDisableOptions"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strGUID - " & strGUID & DotFN(0))
		
		If FnParamChecker(strGUID, "strGUID") = 1 Then
            fnAddLog(VarsFNErr(0) & "strGUID" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnARPDisableOptions = 1
		End If
		
		If fnARPNoRep(strGUID) <> 0 OR fnARPNoMod(strGUID) <> 0 OR fnARPNoRemove(strGUID) <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnARPDisableOptions - fail.") 
			
			strLogTAB = strLogTABBak
			fnARPDisableOptions = 1
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnARPDisableOptions - success.")

			strLogTAB = strLogTABBak			
			fnARPDisableOptions = 0
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will add ARP remove/uninstall, repail, modifity buttons in ARP entry if chosen.
'!
'! @param  strRegKey        		Application strGUID
'! @param  strOptionName        	Option Name to revoke.
'!
'! @return 0,1
Function fnARPDisablePointedOption(strRegKey, strOptionName)
    Dim strCurrentFNName                   : strCurrentFNName = "fnARPDisablePointedOption"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKey - " & strRegKey & ", strOptionName - " & strOptionName & DotFN(0))
    
    'TODO param checking
		
		intRVal = fnCheckIfKeyExist(strRegKey)
		
		If intRVal = 0 Then
			Dim strRegKey :	strRegKey = strRegKey & strOptionName
				strRegType = "REG_DWORD"
				strRegValue = 1
				intRVal = fnAddKey(strRegKey, strRegValue, strRegType)
				fnAddLog("INFO: strGUID is correct, ARP " & strOptionName & " has been proceeding.")
				strLogTAB = strLogTABBak	
				fnARPDisablePointedOption = 0
		Else
				fnAddLog("INFO: strGUID is missing, ARP " & strOptionName & " has not been proceeded.")
				strLogTAB = strLogTABBak	
				fnARPDisablePointedOption = 1	
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will clear cusomizations done by fnARPDisableOptions function.
'!
'! @param  strGUID        			Application strGUID
'!
'! @return 0,1
Function fnARPClearCustoms(strGUID)
	Dim strRegKey
    Dim strCurrentFNName                   : strCurrentFNName = "fnARPClearCustoms"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strGUID - " & strGUID & DotFN(0))
		
		If FnParamChecker(strGUID, "strGUID") = 1 Then
            fnAddLog(VarsFNErr(0) & "strGUID" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnARPClearCustoms = 1
		End If
		
	Dim pathV 							: pathV = Array("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID & "\", "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID & "\")
		
		
		for each elementR in pathV 
			strRegKey = elementR
			intRVal = fnCheckIfKeyExist(strRegKey)
		
			If intRVal = 0 Then
				intRVal = fnCheckIfKeyExist(strRegKey & "NoRepair")
				If intRVal = 0 Then
					intRVal = fnRemoveKey(strRegKey & "NoRepair")
					
					If intRVal <> 0 Then
						fnAddLog("INFO: strGUID NoRepair ARP " & strRegKey & " has not been proceeded.")
					Else 
						fnAddLog("INFO: strGUID NoRepair ARP " & strRegKey & " has been proceeded.")
					End If
				Else 	
					fnAddLog("INFO: strGUID NoRepair is missing, ARP: " & strRegKey & "NoRepair" & ". Going next.")
				End If
				
				intRVal = fnCheckIfKeyExist(strRegKey & "NoModify")
				If intRVal = 0 Then
					intRVal = fnRemoveKey(strRegKey & "NoModify")
					
					If intRVal <> 0 Then
						fnAddLog("INFO: strGUID NoModify ARP " & strRegKey & " has not been proceeded.")
					Else 
						fnAddLog("INFO: strGUID NoModify ARP " & strRegKey & " has been proceeded.")
					End If
				Else 	
					fnAddLog("INFO: strGUID NoModify is missing, ARP: " & strRegKey & "NoModify" & ". Going next.")
				End If
				
				intRVal = fnCheckIfKeyExist(strRegKey & "NoRemove")
				If intRVal = 0 Then
					intRVal = fnRemoveKey(strRegKey & "NoRemove")
					
					If intRVal <> 0 Then
						fnAddLog("INFO: strGUID NoRemove ARP " & strRegKey & " has not been proceeded.")
					Else 
						fnAddLog("INFO: strGUID NoRemove ARP " & strRegKey & " has been proceeded.")
					End If
				Else 	
					fnAddLog("INFO: strGUID NoRemove is missing, ARP: " & strRegKey & "NoRemove" & ". Going next.")
				End If
			Else
				fnAddLog("INFO: strGUID is missing, ARP: " & strRegKey & " . Exiting function.")
				
				strLogTAB = strLogTABBak
				fnARPClearCustoms = 0
			End If
		next
		
		fnAddLog("INFO: strGUID customs ARP removed: " & strRegKey & " . Exiting function.")
		
		strLogTAB = strLogTABBak
		fnARPClearCustoms = 0		
End Function
'-------------------------------------------------------------------------------------------
'! Will add custom registry entry to all logged users.
'!
'! @param  strRegKey        			Registry strRegKey to add.
'! @param  strRegKeyValue        		strRegKeyValue for registry strRegKey.
'! @param  sType        		Data type for registry strRegKey.
'!
'! @return 0,1
Function AddLoggedUserEntry(strRegKey, strRegKeyValue, sType)	'AO 
	Const HKEY_USERS = &H80000003
	Dim arrSubKeys, strSubKey, bKey
    Dim strCurrentFNName                   : strCurrentFNName = "AddLoggedUserEntry"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
	strLogTAB = strLogTAB & strLogSeparator
            
    fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKey - " & strRegKey & ", strRegKeyValue - " & strRegKeyValue & ", sType - " & sType & DotFN(0))
	
	If FnParamChecker(strRegKey, "strRegKey") = 1 Then
        fnAddLog(VarsFNErr(0) & "strRegKey" & VarsFNErr(1))
        fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
		
		strLogTAB = strLogTABBak
		fnFinalize(1)
		AddLoggedUserEntry = 1
	End If
	
	If FnParamChecker(strRegKeyValue, "strRegKeyValue") = 1 Then
        fnAddLog(VarsFNErr(0) & "strRegKeyValue" & VarsFNErr(1))
        fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
		
		strLogTAB = strLogTABBak
		fnFinalize(1)
		AddLoggedUserEntry = 1
	End If
	
	'sType is optional!!!
	
	objReg.EnumKey HKEY_USERS, "", arrSubKeys
	For Each strSubKey In arrSubKeys
		On Error resume Next
		If InStr(strSubKey, "_Classes") = 0 Then ' To exclude registry strRegKey with _Classes - T.D.
			bKey = objShell.RegRead("HKEY_USERS\" & strSubKey & "\Volatile Environment\USERNAME") ' if there is an USERNAME value, it means that user currently logged on a machine
			If bKey <> "" then
				intRVal = fnAddKey("HKEY_USERS\" & strSubKey & "\" & strRegKey, strRegKeyValue, sType) ' added \ between strSubKey and strRegKey - T.D.
			Else
				fnAddLog("INFO: No logged user found!")
				intRVal = 0
			End If
		End If
	Next
	
	If intRVal <> 0 Then
		fnAddLog("FUNCTION INFO: ending AddLoggedUserEntry - fail.") 
			
		strLogTAB = strLogTABBak
		AddLoggedUserEntry = intRVal
		Exit Function
	Else
		fnAddLog("FUNCTION INFO: ending AddLoggedUserEntry - success.")

		strLogTAB = strLogTABBak			
		AddLoggedUserEntry = intRVal
		Exit Function
	End If
End Function
'-------------------------------------------------------------------------------------------
'! Will remove custom registry entry for all logged users.
'!
'! @param  strRegKey        			Registry strRegKey to add.
'!
'! @return 0,1
Function RemoveLoggedUserEntry(strRegKey)	'AO
	Const HKEY_USERS = &H80000003
	Dim arrSubKeys, strSubKey, bKey
    Dim strCurrentFNName                   : strCurrentFNName = "RemoveLoggedUserEntry"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
	strLogTAB = strLogTAB & strLogSeparator
            
    fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKey - " & strRegKey & DotFN(0))
	
	If FnParamChecker(strRegKey, "strRegKey") = 1 Then
        fnAddLog(VarsFNErr(0) & "strRegKey" & VarsFNErr(1))
        fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
		
		strLogTAB = strLogTABBak
		fnFinalize(1)
		RemoveLoggedUserEntry = 1
	End If
	
	objReg.EnumKey HKEY_USERS, "", arrSubKeys
	For Each strSubKey In arrSubKeys
		On Error resume Next
		If InStr(strSubKey, "_Classes") = 0 Then ' To exclude registry strRegKey with _Classes
			bKey = objShell.RegRead("HKEY_USERS\" & strSubKey & "\Volatile Environment\USERNAME") ' if there is an USERNAME value, it means that user currently logged on a machine
			If bKey <> "" then
				intRVal = fnRemoveKey("HKEY_USERS\" & strSubKey & "\" & strRegKey) ' added \ between strSubKey and strRegKey - T.D.
			Else
				fnAddLog("INFO: No logged user found!")
				intRVal = 0
			End If
		End If
	Next
	
	If intRVal <> 0 Then
		fnAddLog("FUNCTION INFO: ending RemoveLoggedUserEntry - fail.") 
			
		strLogTAB = strLogTABBak
		RemoveLoggedUserEntry = intRVal
		Exit Function
	Else
		fnAddLog("FUNCTION INFO: ending RemoveLoggedUserEntry - success.")

		strLogTAB = strLogTABBak			
		RemoveLoggedUserEntry = intRVal
		Exit Function
	End If
End Function
'-------------------------------------------------------------------------------------------
'! Will change custom registry entry for all logged users.
'!
'! @param  strRegKey        			Registry strRegKey to add.
'! @param  strRegKeyNewValue        	New registry strRegKey value.
'!
'! @return 0,1
Function ChangeLoggedUserEntry(strRegKey, strRegKeyNewValue)	'AO
	Const HKEY_USERS = &H80000003
	Dim arrSubKeys, strSubKey, bKey
    Dim strCurrentFNName                   : strCurrentFNName = "ChangeLoggedUserEntry"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
	strLogTAB = strLogTAB & strLogSeparator
            
    fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKey - " & strRegKey & ", strRegKeyNewValue - " & strRegKeyNewValue & DotFN(0))
	
	If FnParamChecker(strRegKey, "strRegKey") = 1 Then
        fnAddLog(VarsFNErr(0) & "strRegKey" & VarsFNErr(1))
        fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
		
		strLogTAB = strLogTABBak
		fnFinalize(1)
		ChangeLoggedUserEntry = 1
	End If
	
	If FnParamChecker(strRegKeyNewValue, "strRegKeyNewValue") = 1 Then
        fnAddLog(VarsFNErr(0) & "strRegKeyNewValue" & VarsFNErr(1))
        fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))  
		
		strLogTAB = strLogTABBak
		fnFinalize(1)
		ChangeLoggedUserEntry = 1
	End If
	
	objReg.EnumKey HKEY_USERS, "", arrSubKeys
	For Each strSubKey In arrSubKeys
		On Error resume Next
		If InStr(strSubKey, "_Classes") = 0 Then ' To exclude registry strRegKey with _Classes
			bKey = objShell.RegRead("HKEY_USERS\" & strSubKey & "\Volatile Environment\USERNAME") ' if there is an USERNAME value, it means that user currently logged on a machine
			If bKey <> "" then
				intRVal = fnChangeKey("HKEY_USERS\" & strSubKey & "\" & strRegKey, strRegKeyNewValue) ' added \ between strSubKey and strRegKey - T.D.
			Else
				fnAddLog("INFO: No logged user found!")
				intRVal = 0
			End If
		End If
	Next
	
	If intRVal <> 0 Then
		fnAddLog("FUNCTION INFO: ending ChangeLoggedUserEntry - fail.") 
			
		strLogTAB = strLogTABBak
		ChangeLoggedUserEntry = intRVal
		Exit Function
	Else
		fnAddLog("FUNCTION INFO: ending ChangeLoggedUserEntry - success.")

		strLogTAB = strLogTABBak			
		ChangeLoggedUserEntry = intRVal
		Exit Function
	End If
End Function
'-------------------------------------------------------------------------------------------
'! Will remove all registry strRegKey tree.
'!
'! @param  KeyHive        		Registry strRegKey main hive.
'! @param  strRegKey        			Registry strRegKey Name.
'!
'! @return 0,1
Function fnRemoveKeyAndSubkeys(KeyHive, strRegKey)
	Dim strSubkey, arrSubkeys
    Dim strCurrentFNName                   : strCurrentFNName = "fnRemoveKeyAndSubkeys"
	Dim intRVal							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "KeyHive - " & KeyHive & ", strRegKey - " & strRegKey & DotFN(0))
		
		If FnParamChecker(KeyHive, "KeyHive") = 1 Then
            fnAddLog(VarsFNErr(0) & "KeyHive" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRemoveKeyAndSubkeys = 1
		End If
		
		If FnParamChecker(strRegKey, "strRegKey") = 1 Then
            fnAddLog(VarsFNErr(0) & "strRegKey" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRemoveKeyAndSubkeys = 1
		End If
		
		
	'----- Converting KeyHive parameter into value:
		If KeyHive = "HKCR" OR KeyHive = "HKEY_CLASSES_ROOT" Then
			KeyHive = &H80000000
		ElseIf KeyHive = "HKCU" OR KeyHive = "HKEY_CURRENT_USER" Then
			KeyHive = &H80000001
		ElseIf KeyHive = "HKLM" OR KeyHive = "HKEY_LOCAL_MACHINE" Then
			KeyHive = &H80000002
		ElseIf KeyHive = "HKU" OR KeyHive = "HKEY_USERS" Then
			KeyHive = &H80000003
		ElseIf KeyHive = "HKCC" OR KeyHive = "HKEY_CURRENT_CONFIG" Then
			KeyHive = &H80000005
		End If
		
		
		objReg.EnumKey KeyHive, strRegKey, arrSubkeys
		
		If IsArray(arrSubkeys) Then
			For Each strSubkey In arrSubkeys
				Call fnRemoveKeyAndSubkeys(KeyHive, strRegKey & "\" & strSubkey)
			Next
		End If
		
		
		On Error Resume Next
			objReg.DeleteKey KeyHive, strRegKey
		
		If Err = 0 Then
			fnAddLog("INFO: Deleting strRegKey: " & strRegKey & ".")
			intRVal = 0			' DeleteKey success
		End If
		
		Err.Clear
			wshShell.RegRead(strRegKey)
		If Err <> 0 Then		' strRegKey not exist
			If intRVal = 0 Then	' strRegKey not exist and DeleteKey success
				fnAddLog("INFO: strRegKey has deleted successfully.")
			Else				' strRegKey not exist and DeleteKey fail
				fnAddLog("INFO/!WARNING: Cannot delete already not existent strRegKey: " & strRegKey & ".")
				intRVal = 0
			End If
		Else					' strRegKey still exists
			fnAddLog("FUNCTION ERROR: Cannot delete strRegKey: " & strRegKey & ".")
			intRVal = 1
		End If
		On Error GoTo 0
		
		
		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnRemoveKeyAndSubkeys - fail.")
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRemoveKeyAndSubkeys = intRVal
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnRemoveKeyAndSubkeys - success.")
			
			strLogTAB = strLogTABBak
			fnRemoveKeyAndSubkeys = intRVal
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will add dummy ARP registry entries for packages without ARP entry.
'!
'!
'! @return 0,1
Function fnAddDummyARP()
	Dim strOKey
	Dim strCurrentFNName					: strCurrentFNName = "fnAddDummyARP"
	Dim strFail							: strFail = 0
	Dim intRVal							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
		
		
		If fnIs64BitOS(False) = 1 Then
			strOKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strPkgGUID & "\"
		Else
			strOKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & strPkgGUID & "\"
		End If
		
		
	'----- Adding dummy ARP registry entries:
		intRVal = fnAddKey(strOKey & "DisplayName", strAppName, "REG_SZ")
		If intRVal = 1 Then
			strFail = 1
		End If
		
		intRVal = fnAddKey(strOKey & "DisplayVersion", strAppNumber, "REG_SZ")
		If intRVal = 1 Then
			strFail = 1
		End If
		
		intRVal = fnAddKey(strOKey & "SystemComponent", "1", "REG_DWORD")
		If intRVal = 1 Then
			strFail = 1
		End If
		
		
	'----- Performing function self check:
		If strFail = 0 Then
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0))	'Success
			
			strLogTAB = strLogTABBak
			fnAddDummyARP = 0
		Else
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))	'Fail
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddDummyARP = 1
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will remove dummy ARP registry strRegKey.
'!
'!
'! @return 0,1
Function fnRemoveDummyARP()
	Dim strOKey
	Dim strCurrentFNName					: strCurrentFNName = "fnRemoveDummyARP"
	Dim intRVal							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
		
		
		If fnIs64BitOS(False) = 1 Then
			strOKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strPkgGUID & "\"
		Else
			strOKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & strPkgGUID & "\"
		End If
		
		
		intRVal = fnRemoveKey(strOKey)
		
		If intRVal = 0 Then
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0))	'Success
			
			strLogTAB = strLogTABBak
			fnRemoveDummyARP = 0
		Else
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))	'Fail
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRemoveDummyARP = 1
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will change ARP entry with custom (customized uninstall string), and hide original one.
'!
'! @param  strGUID        			Application strGUID.
'!
'! @return 0,1
Function fnSubstituteARP(strGUID)
	Dim strBitness, strOKey, strNKey, intRVal

	strBitness = fnAppBitness ()

	strOKey = "HKEY_LOCAL_MACHINE\SOFTWARE\" & strBitness & "Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID & "\"
	strNKey = "HKEY_LOCAL_MACHINE\SOFTWARE\" & strBitness & "Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID & "_" & strPKGName

	intRVal = fnBackupKey(strOKey, strNKey)
	
		intRVal = fnCheckIfKeyExist(strNKey & "\WindowsInstaller") 
		If intRVal = 0 Then
			intRVal = fnRemoveKey(strNKey & "\WindowsInstaller")
		End If
		
		intRVal = fnCheckIfKeyExist(strNKey & "\SystemComponent") 
		If intRVal = 0 Then
			intRVal = fnRemoveKey(strNKey & "\SystemComponent")
		End If
	intRVal = fnAddKey(strNKey & "\NoModify", "1", "REG_DWORD")
	intRVal = fnAddKey(strNKey & "\NoRepair", "1", "REG_DWORD")
	intRVal = fnAddKey(strNKey & "\UninstallString", "wscript.exe " & chr(34) & strTempForPKGSrc & "\" & strPKGName & "\" & strPKGName & ".vbs" & chr(34) & " /remove", "REG_EXPAND_SZ")
	intRVal = fnAddKey(strNKey & "\DisplayIcon", strTempForPKGSrc & "\" & strPKGName & "\_Source\ico.ico", "REG_SZ")
	
	intRVal = fnAddKey(strOKey & "SystemComponent", "1", "REG_DWORD")
	
	fnSubstituteARP = 0
End Function 
'-------------------------------------------------------------------------------------------
'! Will return installed application bitness using strPkgGUID.
'!
'! @return 0,1
Function fnAppBitness ()
	Dim strBitness
	
	If fnIs64BitOS(false) = 0 Then
		If fnCheckIfKeyExist("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strPkgGUID & "\") = 0 Then
			strBitness = ""
		ElseIf fnCheckIfKeyExist("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & strPkgGUID & "\") = 0 Then
			strBitness = "Wow6432Node\"
		Else 
			Exit Function
		End If
	Else
		If fnCheckIfKeyExist("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strPkgGUID & "\") = 0 Then
			strBitness = ""
		Else 
			Exit Function
		End If
	End If
	
	fnAppBitness = strBitness
End Function
'-------------------------------------------------------------------------------------------


'-=====-======-=====-=====-=====-=====-
'! empty string.
'!
'! @return N/A
Function joinArray(arrData1, arrData2) ' DJ
    Dim counter, oldUarrData1
        oldUarrData1 = UBound(arrData1) + 1
    ReDim Preserve arrData1(UBound(arrData1) + UBound(arrData2) + 1)

    For counter = oldUarrData1 to UBound(arrData1)
        arrData1(counter) = arrData2(counter - oldUarrData1)
    Next

    joinArray = arrData1
End Function
'-=====-======-=====-=====-=====-=====-
'-------------------------