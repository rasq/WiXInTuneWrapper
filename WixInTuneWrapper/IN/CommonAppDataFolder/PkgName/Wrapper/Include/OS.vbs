'****************************************************************************************************************************************************
' FUNCTIONS: 
'	fnIs64BitShellEnv(boolIsLog) 
'	fnIsServer(boolIsLog)
'	fnIs64BitOS(boolIsLog)
'	fnOSVersion(boolIsLog)
'	fnPrintMachineDetails()
'	fnOSVariableTranslation(strOSVariable, boolIsLog)  
'	fnIsFreeHDDSpace(intIsNeededInMB)
'	fnFreeHDDSpace()
'	fnLocalUsersProfiles(boolIsLog)
'	fnLoggedUsers(boolIsLog)
'	fnOSLanguage()
'   fnOSLanguageCode()
'	fnAddEnvVar(strVarName, strVarValue)
'	fnEditEnvVar(strVarName, strVarValue)
'	fnIsExistEnvVar(strVarName, strVarValue)
'	fnCheckIfBatteryExists()
'	fnCheckMachineChassisType()
'   fnSCCMCache(strInfType)
'
' FUNCTION EXAMPLE USAGE (IN and OUT):
'
'	fnIs64BitShellEnv()
'	Return: 0,1
'
' 	fnIsServer(IsLog)
' 	Return: true, false
'
'	fnIs64BitOS(true/false)
'	Return: 0,1
'
'	fnPrintMachineDetails()
'	Return: Computer Name: IBMIBM-TUP30AM1
'			Registered user:  IBM
'	
'	fnOSVariableTranslation("%windir%", false)
'	Return: C:\Windows
'	fnOSVariableTranslation("%programfiles(x86)%")
'	Return: C:\Program Files (x86) for 64-bit OS or C:\Program Files for 32-bit OS
'
'	fnIsFreeHDDSpace(25400)
'	Return: 0,1
'	fnIsFreeHDDSpace()
'	Return: free disk space in MB
'
'	fnFreeHDDSpace()
'	Return: free disk space in MB
'
'
'	fnLocalUsersProfiles(boolIsLog)
'	Return: string with User Profiles in machine divided by ";"
'
'
'	fnLoggedUsers(boolIsLog)
'	Return: string with logged in Users in machine divided by ";"
'
'
'	fnOSLanguage()
'	Return: language Name string (English, Spanish... etc - listing in function), if language is not in our list then return English.
'
'
'	fnOSLanguageCode()
'	Return: language code string (1033, 1045... etc - listing in function), if language is not in our list then return 1033.
'
'	fnAddEnvVar(strVarName, strVarValue)
'	Return: 0, 1
'
'
'	fnEditEnvVar(strVarName, strVarValue)
'	Return: 0, 1
'
'
'	fnIsExistEnvVar(strVarName, strVarValue)
'	Return: 0, 1
'
'	fnCheckIfBatteryExists()
'	Return: 0 - battery found, 1 - battery not found
'
'	fnCheckMachineChassisType()
'	Return: 0 - notebook/laptop, 1 - workstation
'
'***************************************************************************************************************************************************
Public Function fnIs64BitShellEnv(boolIsLog) 
	Dim strShellType, colProcessEnvVars, intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnIs64BitShellEnv"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		If fnIsNull(boolIsLog) = True or boolIsLog = "" or isEmpty(boolIsLog) = True Then
			boolIsLog = true
		End If
		
		If boolIsLog = true Then
            fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
		End If
		
		
		Set colProcessEnvVars = objShell.Environment("Process") 
		strShellType = colProcessEnvVars("PROCESSOR_ARCHITECTURE")

		If strShellType = "AMD64" Or strShellType = "x64" Or strShellType = "X64" Then
			If boolIsLog = true Then
				fnAddLog("FUNCTION INFO: 64-bit shell is running.")
			End If
			intRVal = 0
		Else
			If boolIsLog = true Then
				fnAddLog("FUNCTION INFO: 32-bit shell is running.")
			End If
			intRVal = 1
		End If
		
		If boolIsLog = true Then
			call Const_FnEndrVal(strCurrentFNName, intRVal)
		End If

		strLogTAB = strLogTABBak
		fnIs64BitShellEnv = intRVal
End Function
'-------------------------------------------------------------------------------------------
Function fnIsServer(boolIsLog)
	Dim colItems, objItem, strProductType, intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnIsServer"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		If fnIsNull(boolIsLog) = True or boolIsLog = "" or isEmpty(boolIsLog) = True Then
			boolIsLog = true
		End If
		
		If boolIsLog = true Then
            fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
		End If
		
		On Error Resume Next
		
		Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
		
		For Each objItem in colItems
			strProductType = objItem.ProductType
		Next
		
		If strProductType = 1 Then 'workstation
			If boolIsLog = true Then
				fnAddLog("FUNCTION INFO: OS identified as a WORKSTATION. strProductType = " & strProductType & ". Installation will continue.")
			End If
			intRVal = false
		Else 'server
			If boolIsLog = true Then
				fnAddLog("FUNCTION INFO: OS identified as a SERVER. strProductType = " & strProductType & ". Installation halted.")
			End If
			intRVal = true
		End If
		
		If boolIsLog = true Then
			call Const_FnEndrVal(strCurrentFNName, intRVal)
		End If
		
		strLogTAB = strLogTABBak
		fnIsServer = intRVal
End Function 
'-------------------------------------------------------------------------------------------
Function fnIs64BitOS(boolIsLog)
    Dim strCurrentFNName                   : strCurrentFNName = "fnIs64BitOS"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		If fnIsNull(boolIsLog) = True or boolIsLog = "" or isEmpty(boolIsLog) = True Then
			boolIsLog = true
		End If
		
		if boolIsLog = true Then
            fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
		End If
		
		Dim strKeyPath, strValueName, strValue, intRVal

		Const HKEY_LOCAL_MACHINE = &H80000002
			strKeyPath = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
			strValueName = "Identifier"

		objReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue
		
		If (instr(strValue,"64")) Then
			if boolIsLog = true Then
				fnAddLog("FUNCTION INFO: 64-bit OS detected.")
			End If
			intRVal=0
		Else
			if boolIsLog = true Then
				fnAddLog("FUNCTION INFO: 64-bit OS not detected.")
			End If
			intRVal=1
		End If
	  
			if boolIsLog = true Then
				call Const_FnEndrVal(strCurrentFNName, intRVal)
			End If
		
		strLogTAB = strLogTABBak
		fnIs64BitOS = intRVal
		
		
		'If GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth = 64 Then ' a to nie bedzie lepsze jako samo wykrywanie? 
		'    fnIs64BitOS = True
		'Else
		'    fnIs64BitOS = False
		'End If
End Function
'-------------------------------------------------------------------------------------------
Function fnPrintMachineDetails()
	Dim colItems, objOS
    Dim strCurrentFNName                   : strCurrentFNName = "fnPrintMachineDetails"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
		
		Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
		
		For Each objOS in colItems
			fnAddLog("FUNCTION INFO: Computer Name:    " & objOS.CSName)
			fnAddLog("FUNCTION INFO: Registered user:  " & objOS.RegisteredUser)
			fnAddLog("FUNCTION INFO: OS version:       " & objOS.Caption )
			fnAddLog("FUNCTION INFO: OS architecture:  " & objOS.OSArchitecture)
			fnAddLog("FUNCTION INFO: OS language:      " & objOS.fnOSLanguage)
			fnAddLog("FUNCTION INFO: Service pack:     " & objOS.ServicePackMajorVersion)
			fnAddLog("FUNCTION INFO: Free disk space:  " & fnFreeHDDSpace() & " MB")
		Next

		fnAddLog("FUNCTION INFO: Ending function fnPrintMachineDetails.")
		
		strLogTAB = strLogTABBak
End Function 
'-------------------------------------------------------------------------------------------
Function fnOSVariableTranslation(strOSVariable, boolIsLog)
	Dim intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnOSVariableTranslation"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		If fnIsNull(boolIsLog) = True or boolIsLog = "" or isEmpty(boolIsLog) = True Then
			boolIsLog = true
		End If
	
    'TODO variable checking
    
    
		if boolIsLog = true Then
            fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strOSVariable - " & strOSVariable & DotFN(0)) 
		End If
		
		If strOSVariable = "%PROGRAMFILES%" or strOSVariable = "%programfiles%" Then
			If fnIs64BitOS(false) = 0 Then '64-bit OS
				intRVal = objShell.ExpandEnvironmentStrings("%PROGRAMW6432%") ''C:\Program Files - for 64-bit applications
			Else
				intRVal = objShell.ExpandEnvironmentStrings("%PROGRAMFILES%") ''C:\Program Files - for 32-bit applications
			End If	
		ElseIf strOSVariable = "%COMMONPROGRAMFILES%" or strOSVariable = "%commonprogramfiles%" Then
			If fnIs64BitOS(false) = 0 Then '64-bit OS
				intRVal = objShell.ExpandEnvironmentStrings("%COMMONPROGRAMW6432%") 'C:\Program Files\Common File - for 64-bit applications
			Else
				intRVal = objShell.ExpandEnvironmentStrings("%COMMONPROGRAMFILES%") 'C:\Program Files\Common File - for 32-bit applications
			End If
		ElseIf strOSVariable = "%PROGRAMFILES(x86)%" or strOSVariable = "%programfiles(x86)%" Then	
			If fnIs64BitOS(false) = 0 Then
				intRVal = objShell.ExpandEnvironmentStrings("%PROGRAMFILES(x86)%") 'C:\Program Files - for 32-bit applications
			Else
				intRVal = objShell.ExpandEnvironmentStrings("%PROGRAMFILES%") 'C:\Program Files - for 32-bit applications
			End If
		ElseIf strOSVariable = "%COMMONPROGRAMFILES(x86)%" or strOSVariable = "%commonprogramfiles(x86)%" Then
			If fnIs64BitOS(false) = 0 Then
				intRVal = objShell.ExpandEnvironmentStrings("%COMMONPROGRAMFILES(x86)%") 'C:\Program Files\Common File - for 32-bit applications
			Else
				intRVal = objShell.ExpandEnvironmentStrings("%COMMONPROGRAMFILES%") 'C:\Program Files\Common File - for 32-bit applications
			End If
		Else
			intRVal = objShell.ExpandEnvironmentStrings(strOSVariable)
		End If
		
		if boolIsLog = true Then
			call Const_FnEndrVal(strCurrentFNName, intRVal)
		End If
		
		strLogTAB = strLogTABBak
		fnOSVariableTranslation = intRVal
End Function 
'-------------------------------------------------------------------------------------------
'! Reads HDD free space data and compare it with parameter. 
'!
'! @param  intIsNeededInMB     How much space package require for installation.
'!
'! @return 1,0
Function fnIsFreeHDDSpace(intIsNeededInMB)
	Dim objWMIServiceL, objLogicalDisk, strDiskSpaceInMB, intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnIsFreeHDDSpace"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		If FnParamChecker(intIsNeededInMB, "intIsNeededInMB") = 1 Then
			fnAddLog("FUNCTION INFO: Starting function fnIsFreeHDDSpace with parameter NULL.")
		Else
			fnAddLog("FUNCTION INFO: Starting function fnIsFreeHDDSpace with parameter " & intIsNeededInMB & ".")
			fnAddLog("FUNCTION INFO: Required free space on C drive: " & intIsNeededInMB & " MB.")
		End If

		Set objWMIServiceL = GetObject("winmgmts:")
		Set objLogicalDisk = objWMIServiceL.Get("Win32_LogicalDisk.DeviceID='c:'")
		
		strDiskSpaceInMB = Clng(objLogicalDisk.FreeSpace / 1024 / 1024)
		fnAddLog("FUNCTION INFO: Free space on C drive: " & strDiskSpaceInMB & " MB.")
		
		If fnIsNull (intIsNeededInMB) Then
			intRVal = strDiskSpaceInMB
		ElseIf strDiskSpaceInMB > intIsNeededInMB Then
			intRVal = 0
		Else
			intRVal = 1
		End If

		call Const_FnEndrVal(strCurrentFNName, intRVal)
		
		strLogTAB = strLogTABBak
		fnIsFreeHDDSpace = intRVal
End Function
'-------------------------------------------------------------------------------------------
'! Reads HDD free space data. 
'!
'! @return fnFreeHDDSpace
Function fnFreeHDDSpace()
	Dim objWMIServiceL, objLogicalDisk, strDiskSpaceInMB, intRVal
		
	Set objWMIServiceL = GetObject("winmgmts:")
	Set objLogicalDisk = objWMIServiceL.Get("Win32_LogicalDisk.DeviceID='c:'")
	
	intRVal = Clng(objLogicalDisk.FreeSpace / 1024 / 1024)
	
	fnFreeHDDSpace = intRVal
End Function
'-------------------------------------------------------------------------------------------
Function fnLocalUsersProfiles(boolIsLog)
	Dim strUsersProfiles
    Dim strCurrentFNName                   : strCurrentFNName = "fnLocalUsersProfiles"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
	Const HKEY_LOCAL_MACHINE = &H80000002

		If fnIsNull(boolIsLog) = True or boolIsLog = "" or isEmpty(boolIsLog) = True Then
			boolIsLog = true
		End If
		
		If boolIsLog = true Then
            fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
		End If
			
			strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
			objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys

			For Each objSubkey In arrSubkeys
				strValueName = "ProfileImagePath"
				strSubPath = strKeyPath & "\" & objSubkey
				objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue
				
				if objSubkey <> "S-1-5-18" and objSubkey <> "S-1-5-19" and objSubkey <> "S-1-5-20" Then
						if strUsersProfiles <> "" Then
							strUsersProfiles = strValue & ";" & strUsersProfiles
						Else
							strUsersProfiles = strValue & ";"
						End If
					If boolIsLog = true Then
						fnAddLog("FUNCTION INFO: registry info about local accounts: " & objSubkey & " -> " & strValue & ".")
					End If
				End If
			Next
	
		If boolIsLog = true Then
			fnAddLog("FUNCTION INFO: Ending function fnLocalUsersProfiles.")
		End If
		
		strLogTAB = strLogTABBak
		fnLocalUsersProfiles = strUsersProfiles
End Function
'-------------------------------------------------------------------------------------------
Function fnLoggedUsers(boolIsLog)
	Dim strUserName, objUserItems
    Dim strCurrentFNName                   : strCurrentFNName = "fnLoggedUsers"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		Set objUserItems = CreateObject( "Scripting.System" )

		If fnIsNull(boolIsLog) = True or boolIsLog = "" or isEmpty(boolIsLog) = True Then
			boolIsLog = true
		End If
		
		If boolIsLog = true Then
            fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
		End If
			
		if objUserItems.Count = 0 Then	
			If boolIsLog = true Then
				fnAddLog("FUNCTION INFO: Ending function fnLoggedUsers with empty string.")
			End If
			strUserName = ""
			fnLoggedUsers = strUserName
			Exit Function
		Else
			For Each objItem in objUserItems 
				if strUserName <> "" Then
					strUserName = objItem.UserName & ";" & strUserName
				Else
					strUserName = objItem.UserName & ";"
				End If
					
				If boolIsLog = true Then
					fnAddLog("FUNCTION INFO: registry info about logged in accounts: " & objItem.UserName & ".")
				End If
			Next
		End If
	
		If boolIsLog = true Then
			fnAddLog("FUNCTION INFO: Ending function fnLoggedUsers.")
		End If
		
		strLogTAB = strLogTABBak
		fnLoggedUsers = strUserName
End Function
'-------------------------------------------------------------------------------------------
Function fnOSLanguage()
	Dim intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnOSLanguage"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
	
	Dim rEnglish							: rEnglish = Array(1033, 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 12297)
	Dim rGerman								: rGerman = Array(1031, 3079, 5127, 4103, 2055)
	Dim rDutch								: rDutch = Array(1043, 2067)
	Dim rJapanese							: rJapanese = Array(1041)
	Dim rFrench								: rFrench = Array(1036, 2060, 3084, 5132, 6156, 4108)
	Dim rSpanish							: rSpanish = Array(3082, 8202, 14346, 1034, 20490, 10250, 15370, 6154, 19466, 2058, 18442, 4106, 17418, 12298, 7178, 5130, 9226, 13322, 16394, 11274)
	Dim rItalian							: rItalian = Array(1040, 2064)
	Dim rTurkish							: rTurkish = Array(1055)
	
	Dim rLang								: rLang = Array("English", "German", "Dutch", "Japanese", "French", "Spanish", "Italian", "Turkish")
            
            fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
			
			
			intRVal = GetLocale

			for each lang in rEnglish
				If lang = intRVal Then
					intRVal = "English"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguage = intRVal
					Exit Function
				End If
			Next
			
			for each lang in rGerman
				If lang = intRVal Then
					intRVal = "German"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguage = intRVal
					Exit Function
				End If
			Next
			
			for each lang in rDutch
				If lang = intRVal Then
					intRVal = "Dutch"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguage = intRVal
					Exit Function
				End If
			Next
			
			for each lang in rJapanese
				If lang = intRVal Then
					intRVal = "Japanese"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguage = intRVal
					Exit Function
				End If
			Next
			
			for each lang in rFrench
				If lang = intRVal Then
					intRVal = "French"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguage = intRVal
					Exit Function
				End If
			Next
			
			for each lang in rSpanish
				If lang = intRVal Then
					intRVal = "Spanish"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguage = intRVal
					Exit Function
				End If
			Next
			
			for each lang in rItalian
				If lang = intRVal Then
					intRVal = "Italian"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguage = intRVal
					Exit Function
				End If
			Next
			
			for each lang in rTurkish
				If lang = intRVal Then
					intRVal = "Turkish"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguage = intRVal
					Exit Function
				End If
			Next
	
		
		intRVal = "English" 'default lang as English
		call Const_FnEndrVal(strCurrentFNName, intRVal)
		strLogTAB = strLogTABBak
		
	fnOSLanguage = intRVal
End Function
'-------------------------------------------------------------------------------------------
Function fnOSLanguageCode()
	Dim intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnOSLanguageCode"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
	
	Dim rEnglish							: rEnglish = Array(1033, 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 12297)
	Dim rGerman								: rGerman = Array(1031, 3079, 5127, 4103, 2055)
	Dim rDutch								: rDutch = Array(1043, 2067)
	Dim rJapanese							: rJapanese = Array(1041)
	Dim rFrench								: rFrench = Array(1036, 2060, 3084, 5132, 6156, 4108)
	Dim rSpanish							: rSpanish = Array(3082, 8202, 14346, 1034, 20490, 10250, 15370, 6154, 19466, 2058, 18442, 4106, 17418, 12298, 7178, 5130, 9226, 13322, 16394, 11274)
	Dim rItalian							: rItalian = Array(1040, 2064)
	Dim rTurkish							: rTurkish = Array(1055)
	
	Dim rLang								: rLang = Array("English", "German", "Dutch", "Japanese", "French", "Spanish", "Italian", "Turkish")
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
			
			
			intRVal = GetLocale

			for each lang in rEnglish
				If lang = intRVal Then
					intRVal = "1033"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguageCode = intRVal
					Exit Function
				End If
			Next
			
			for each lang in rGerman
				If lang = intRVal Then
					intRVal = "1031"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguageCode = intRVal
					Exit Function
				End If
			Next
			
			for each lang in rDutch
				If lang = intRVal Then
					intRVal = "1043"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguageCode = intRVal
					Exit Function
				End If
			Next
			
			for each lang in rJapanese
				If lang = intRVal Then
					intRVal = "1041"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguageCode = intRVal
					Exit Function
				End If
			Next
			
			for each lang in rFrench
				If lang = intRVal Then
					intRVal = "1036"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguageCode = intRVal
					Exit Function
				End If
			Next
			
			for each lang in rSpanish
				If lang = intRVal Then
					intRVal = "3082"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguageCode = intRVal
					Exit Function
				End If
			Next
			
			for each lang in rItalian
				If lang = intRVal Then
					intRVal = "1040"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguageCode = intRVal
					Exit Function
				End If
			Next
			
			for each lang in rTurkish
				If lang = intRVal Then
					intRVal = "1055"
					call Const_FnEndrVal(strCurrentFNName, intRVal)
					strLogTAB = strLogTABBak
					fnOSLanguageCode = intRVal
					Exit Function
				End If
			Next
	
		
		intRVal = "1033" 'default lang as English
		call Const_FnEndrVal(strCurrentFNName, intRVal)
		strLogTAB = strLogTABBak
		
	fnOSLanguageCode = intRVal
End Function
'-------------------------------------------------------------------------------------------
Function fnAddEnvVar(strVarName, strVarValue)
	Dim intRVal, objVar
    Dim strCurrentFNName                   : strCurrentFNName = "fnAddEnvVar"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strVarName - " & strVarName & "strVarValue - " & strVarValue & DotFN(0)) 
		
			If FnParamChecker(strVarName, "strVarName") = 1 Then
                call fnConst_VarFail("strVarName", strCurrentFNName)
				
				strLogTAB = strLogTABBak
				fnFinalize(1)
				fnAddEnvVar = 1
			End If
			
			If FnParamChecker(strVarValue, "strVarValue") = 1 Then
                fnAddLog(VarsFNErr(0) & "strVarValue" & VarsFNErr(1))

                    strVarValue = ""

                fnAddLog(SetVarFN(0) & "strVarValue" & SetVarFN(1) & "empty string" & DotFN(0))
			End If
			
			Set objVar      			= objVarClass.SpawnInstance_
				objVar.Name          	= strVarName
				objVar.VariableValue 	= strVarValue
				objVar.UserName      	= "<SYSTEM>"	
				objVar.Put_
				
			fnAddLog("FUNCTION INFO: Created environment variable " & strVarName & " with Value: " & strVarValue & ".")
			
			
			Set objVar      = Nothing
			'Set objVarClass = Nothing
			
			intRVal = fnIsExistEnvVar(strVarName, strVarValue)
		
		call Const_FnEndrVal(strCurrentFNName, intRVal)
		strLogTAB = strLogTABBak
		
	fnAddEnvVar = intRVal
End Function
'-------------------------------------------------------------------------------------------
Function fnEditEnvVar(strVarName, strVarValue)
	Dim intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnEditEnvVar"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strVarName - " & strVarName & "strVarValue - " & strVarValue & DotFN(0)) 
		
			If FnParamChecker(strVarName, "strVarName") = 1 Then
                call fnConst_VarFail("strVarName", strCurrentFNName)
				
				strLogTAB = strLogTABBak
				fnFinalize(1)
				fnEditEnvVar = 1
			End If
			
			If FnParamChecker(strVarValue, "strVarValue") = 1 Then
                call fnConst_VarFail("strVarValue", strCurrentFNName)
				
				strLogTAB = strLogTABBak
				fnFinalize(1)
				fnEditEnvVar = 1
			End If

			intRVal = fnIsExistEnvVar(strVarName, strVarValue)
			
			if intRVal <> 0 Then
				intRVal = fnAddEnvVar(strVarName, strVarValue)
			End If
			
			
			
		
		call Const_FnEndrVal(strCurrentFNName, intRVal)
		strLogTAB = strLogTABBak
		
	fnEditEnvVar = intRVal
End Function
'-------------------------------------------------------------------------------------------
Function fnIsExistEnvVar(strVarName, strVarValue)
	Dim intRVal, strQuery, colItems
    Dim strCurrentFNName                   : strCurrentFNName = "fnIsExistEnvVar"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		intRVal = 1
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strVarName - " & strVarName & "strVarValue - " & strVarValue & DotFN(0)) 
		
			If FnParamChecker(strVarName, "strVarName") = 1 Then
                call fnConst_VarFail("strVarName", strCurrentFNName)
				
				strLogTAB = strLogTABBak
				fnFinalize(1)
				fnIsExistEnvVar = 1
			End If
			
			If FnParamChecker(strVarValue, "strVarValue") = 1 Then
                fnAddLog(VarsFNErr(0) & "strVarValue" & VarsFNErr(1))

                    strVarValue = ""

                fnAddLog(SetVarFN(0) & "strVarValue" & SetVarFN(1) & "empty string" & DotFN(0))
			End If
			
				strQuery = "SELECT * FROM Win32_Environment WHERE Name='" & strVarName & "'"
				
				Set colItems = objWMIService.ExecQuery( strQuery, "WQL", 48 )

					For Each objItem In colItems
						fnAddLog("Caption        : " & objItem.Caption)
						fnAddLog("Description    : " & objItem.Description)
						fnAddLog("Name           : " & objItem.Name)
						fnAddLog("Status         : " & objItem.Status)
						fnAddLog("SystemVariable : " & objItem.SystemVariable)
						fnAddLog("UserName       : " & objItem.UserName)
						fnAddLog("VariableValue  : " & objItem.VariableValue)
						
						if objItem.Name = strVarName Then
							if strVarValue = objItem.VariableValue Then
								fnAddLog("FUNCTION INFO: System variable exist with this same value.")					
								intRVal = 0
							Else
								fnAddLog("FUNCTION INFO: System variable exist without this same value.")
								intRVal = 0
							End If
						End If
					Next

				Set colItems      = Nothing
		
		call Const_FnEndrVal(strCurrentFNName, intRVal)
		strLogTAB = strLogTABBak
		
	fnIsExistEnvVar = intRVal
End Function
'-------------------------------------------------------------------------------------------
'! Reads informations about computer power type and return info if it's use battery or not. 
'!
'! @return 0,1
Function fnCheckIfBatteryExists()
	Dim intRVal, objBattery, colBattery
    Dim strCurrentFNName                   : strCurrentFNName = "fnCheckIfBatteryExists"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
    
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
		
		On Error Resume Next
			intRVal = 1
			Set colBattery = objWMIService.ExecQuery("SELECT * FROM Win32_Battery",,48)
			
			'Searching for battery:
				fnAddLog("FUNCTION INFO: Checking, if battery exists to determine machine type...")
				For Each objBattery In colBattery
					intRVal = 0
				Next
			
			'Checking machine chassis type in case battery has not been found:
				If intRVal = 1 Then
					fnAddLog("FUNCTION INFO: Battery not found. Ensuring that machine is a workstation by checking chassis type...")
					If fnCheckMachineChassisType() = 0 Then
						intRVal = 0
					End If
				Else
					fnAddLog("FUNCTION INFO: Battery found. Machine is recognized as notebook/laptop.")
				End If
			
			Set colBattery = Nothing
			If Err Then Err.Clear
		On Error GoTo 0
		
        call Const_FnEndrVal(strCurrentFNName, intRVal)
		strLogTAB = strLogTABBak
		fnCheckIfBatteryExists = intRVal
		Exit Function
End Function
'-------------------------------------------------------------------------------------------
'! Reads informations about computer chasis type and return info if it's laptop or desktop. 
'!
'! @return 0,1
Function fnCheckMachineChassisType()
	Dim intRVal, objChassis, colChassis, intChassisValue, strChassisType
    Dim strCurrentFNName                   : strCurrentFNName = "fnCheckMachineChassisType"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
    
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
		
		Set colChassis = objWMIService.ExecQuery("SELECT * FROM Win32_SystemEnclosure")
	
		'Searching for machine chassis type:
			For Each objChassis In colChassis
				For Each intChassisValue In objChassis.ChassisTypes
					Select Case intChassisValue
						Case 1
							strChassisType = "Other"
							intRVal = 1
						Case 2
							strChassisType = "Unknown"
							intRVal = 1
						Case 3
							strChassisType = "Desktop"
							intRVal = 1
						Case 4
							strChassisType = "Low Profile Desktop"
							intRVal = 1
						Case 5
							strChassisType = "Pizza Box"
							intRVal = 1
						Case 6
							strChassisType = "Mini Tower"
							intRVal = 1
						Case 7
							strChassisType = "Tower"
							intRVal = 1
						Case 8										'LAPTOP FOUND
							strChassisType = "Portable"
							intRVal = 0
						Case 9										'LAPTOP FOUND
							strChassisType = "Laptop"
							intRVal = 0
						Case 10										'LAPTOP FOUND
							strChassisType = "Notebook"
							intRVal = 0
						Case 11
							strChassisType = "Hand Held"
							intRVal = 1
						Case 12										'LAPTOP FOUND
							strChassisType = "Docking Station"
							intRVal = 0
						Case 13
							strChassisType = "All in One"
							intRVal = 1
						Case 14										'LAPTOP FOUND
							strChassisType = "Sub Notebook"
							intRVal = 0
						Case 15
							strChassisType = "Space-Saving"
							intRVal = 1
						Case 16
							strChassisType = "Lunch Box"
							intRVal = 1
						Case 17
							strChassisType = "Main System Chassis"
							intRVal = 1
						Case 18
							strChassisType = "Expansion Chassis"
							intRVal = 1
						Case 19
							strChassisType = "SubChassis"
							intRVal = 1
						Case 20
							strChassisType = "Bus Expansion Chassis"
							intRVal = 1
						Case 21
							strChassisType = "Peripheral Chassis"
							intRVal = 1
						Case 22
							strChassisType = "Storage Chassis"
							intRVal = 1
						Case 23
							strChassisType = "Rack Mount Chassis"
							intRVal = 1
						Case 24
							strChassisType = "Sealed-Case PC"
							intRVal = 1
						Case Else
							strChassisType = "CHASSIS NOT FOUND"
							intRVal = 1
					End Select
				Next
			Next
		
		Set colChassis = Nothing
		
		fnAddLog("FUNCTION INFO: Ending function fnCheckMachineChassisType() with intRVal = " & intRVal & ". Machine is recognized as: " & chr(34) & strChassisType & chr(34) & ".")
		strLogTAB = strLogTABBak
		fnCheckMachineChassisType = intRVal
		Exit Function
End Function
'-------------------------------------------------------------------------------------------
Function fnOSVersion(boolIsLog)
    Dim strCurrentFNName                   : strCurrentFNName = "fnOSVersion"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		If fnIsNull(boolIsLog) = True or boolIsLog = "" or isEmpty(boolIsLog) = True Then
			boolIsLog = true
		End If
		
		if boolIsLog = true Then
            fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
		End If
		
        Dim ver, verSplit, colOperatingSystems, objOperatingSystem
            Set colOperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")

        
            For Each objOperatingSystem in colOperatingSystems
                ver = objOperatingSystem.Version
                verSplit = Split(ver, ".")

                    intRVal = verSplit(0)
            Next

	  
			if boolIsLog = true Then
				fnAddLog("FUNCTION INFO: Ending function fnOSVersion with result main NT version = " & intRVal & ".")
			End If
		
		strLogTAB = strLogTABBak
		fnOSVersion = intRVal
End Function
'-------------------------------------------------------------------------------------------
'! Reads SCCM cache space data. 
'!
'! @param  strInfType          total/free
'!
'! @return -1,fnSCCMCache
Function fnSCCMCache(strInfType)
    On Error Resume Next
    Dim objUIResManager 
    Dim objCache
        
        Set objUIResManager = CreateObject("UIResource.UIResourceMgr")
        Set objCache = objUIResManager.GetCacheInfo()
    
    If strInfType = "total" Then
        fnSCCMCache = objCache.TotalSize
    ElseIf strInfType = "free" Then
        fnSCCMCache = objCache.FreeSize
    Else
        fnSCCMCache = -1      
    End If
End Function
'-------------------------------------------------------------------------------------------
'! Translates a SID to user Name. 
'!
'! @param  strSIDIn          user SID
'!
'! @return userName
Function fnSIDTranslate(strSIDIn)
    Dim strCurrentFNName                   : strCurrentFNName = "fnSIDTranslate"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
    Dim objAccount		                
		strLogTAB = strLogTAB & strLogSeparator
		
            If FnParamChecker(strVarName, "strSIDIn") = 1 Then
                call fnConst_VarFail("strSIDIn", strCurrentFNName)
				
				strLogTAB = strLogTABBak
				fnFinalize(1)
				fnIsExistEnvVar = 1
			End If
        
        Set objAccount = objWMIService.Get("Win32_SID.SID='" & strSIDIn & "'")
        Dim strUserName		            : strUserName = objAccount.strAccountName

    fnSIDTranslate = strUserName
End Function
'-------------------------------------------------------------------------------------------