'****************************************************************************************************************************************************
' FUNCTIONS:  
' fnRunCmd(strCmd, strArgs) 
' fnGetDateAndTime()
' fnGetDate()
' fnGetTime()
' fnNoFlags()
' fnRequestPortal()
' fnSetInstall()
' fnSetUninstall()
' fnProgressBarSet(boolInstMode)
' fnFinalizeIfFail(strFunctionName, intRVal)
' fnGetMSOfficeVersions()
' fnGetDotNetVersions()
' fnAllSoftwareList()
' fnCMDParams()
' fnInitialize()
' fnPreSetup()
' fnFinalize(intTempRVal)
' fnRefreshExplorer() 'check and fix to w10
' fnRetCodeChecker(intRetCode, arrRCList) 
' fnLastCharChecker(strValue, strChar) 
'	
' FUNCTION EXAMPLE USAGE (IN and OUT):
' fnRunCmd("notepad.exe", "plik.txt") 
' 	Return 1, 0
' 
' fnGetDateAndTime()
' 	Return xxxx_xx_xx_xx_xx_xx
' 
' fnGetDate()
' 	Return xxxx_xx_xx
' 
' fnGetTime()
' 	Return xx_xx_xx
' 
' fnRetCodeChecker("3010", "0,-123,3010") 
'   Return intRetCode (or will end with fail if intRVal not present in proper RCodes)
'
'
'
'***************************************************************************************************************************************************
'! Generate date and time string.
'!
'! @return Data_Time (xxxx_xx_xx_xx_xx_xx)
Function fnGetDateAndTime()
    Dim strCurrentFNName                   : strCurrentFNName = "fnGetDateAndTime"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2)) 
			fnGetDateAndTime = fnGetDate() & "_" & fnGetTime()
		fnAddLog("FUNCTION INFO: ending fnGetDateAndTime.") 
	
	strLogTAB = strLogTABBak 
End Function
'-------------------------------------------------------------------------------------------
'! Generate date string.
'!
'! @return Data (xxxx_xx_xx)
Function fnGetDate()
    Dim strCurrentFNName                   : strCurrentFNName = "fnGetDate"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2)) 
			fnGetDate = Year(Date) & "_" & Month(Date) & "_" & Day(Date)
		fnAddLog("FUNCTION INFO: ending fnGetDate.") 
	
	strLogTAB = strLogTABBak 
End Function
'-------------------------------------------------------------------------------------------
'! Generate time string.
'!
'! @return Time (xx_xx_xx)
Function fnGetTime()
    Dim strCurrentFNName                   : strCurrentFNName = "fnGetTime"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
			fnGetTime = hour(time) & "_" & minute(time) & "_" & second(time)
		fnAddLog("FUNCTION INFO: ending fnGetTime.") 
	
	strLogTAB = strLogTABBak 
End Function
'-------------------------------------------------------------------------------------------
'! Quit script if triggered without parameter (if needed).
'!
'! @return N/A
Function fnNoFlags()
    Dim strCurrentFNName                   : strCurrentFNName = "fnNoFlags"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		strLogName = strPKGName & "_FAIL_VBS.log" 
		call fnStartLog(strLogName)
	
	fnAddLog("SCRIPT FAIL: fnInitialize without flags or without proper flag.")
	
	strLogTAB = strLogTABBak 
		fnFinalize(1)
End Function
'-------------------------------------------------------------------------------------------
'! Check and initialize request portal if needed.
'!
'! @return 0,1
Function fnRequestPortal()
    Dim strCurrentFNName                   : strCurrentFNName = "fnRequestPortal"
	Dim strLogTABBak	                    : strLogTABBak = strLogTAB
	Dim intRVal		                    : intRVal = 0
	Dim strCmd, strRequestFile, strRequestDir
		strLogTAB = strLogTAB & strLogSeparator
		
    fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
	
		strRequestDir = strScriptDir & "\Common"
		strRequestFile = "EntryRequest_U.wsf"
		strCmd = chr(34) & strRequestDir & chr(92) & strRequestFile & chr(34) & " " & strPKGName & " " & chr(34) & strLogFolder & chr(92) & strPKGName & chr(34)
	
	fnAddLog("SCRIPT INFO: Checking, if RequestPortal is needed.")
	
		If EntryRequestURL = "" OR fnIsNull(EntryRequestURL) = True OR InStr(EntryRequestURL, "http:") = 0 Then
			fnAddLog("SCRIPT INFO: RequestPortal is not needed for this account.")
			intRVal = 0
		ElseIf fnOSVersion(True) <> 10 Then
			fnAddLog("SCRIPT INFO: Operating system is not detected as Windows 10. Skipping RequestPortal, going to installation.")
			intRVal = 0
		Else
		
			If fnIsExist(strRequestFile, strRequestDir) = 0 Then
				fnAddLog("SETUP INFO: starting RequestPortal : " & strCmd & ".")
				intRVal = objShell.Run(strCmd, 1, 1)										'Open request portal for packages for one Client
				
				If intRVal <> 0 Then
					fnAddLog("SETUP ERROR: RequestPortal exited with error code " & intRVal & ".")
					fnAddLog("SCRIPT INFO: ending fnRequestPortal. Exiting script with intRVal = 1.")
					intRVal = 1
					strLogTAB = strLogTABBak
					fnFinalize(intRVal)
				End If
			Else
				fnAddLog("SCRIPT INFO: RequestPortal cannot start due to missing file: " & chr(34) & strRequestDir & chr(92) & strRequestFile & chr(34) & ".")
				intRVal = Const_ReqPortalNotEC
				fnAddLog("SCRIPT INFO: ending fnRequestPortal. Exiting script with intRVal = " & intRVal & ".")
				
				strLogTAB = strLogTABBak
				fnFinalize(intRVal)
			End If
		End If
	
	fnAddLog("SCRIPT INFO: ending fnRequestPortal with intRVal = " & intRVal & ".")
	strLogTAB = strLogTABBak
End Function
'-------------------------------------------------------------------------------------------
'! Set appropriate variables to install mode.
'!
'! @return N/A
Function fnSetInstall()
    Dim strTempPKGName
	Dim strLogTABBak	:	strLogTABBak = strLogTAB
	Dim intStrikeVar	:	intStrikeVar = -2000
	Dim intRVal		:   intRVal = 0
		strLogTAB = strLogTAB & strLogSeparator
    
		strLogName = strPKGName & "_Install_VBS.log" 
		boolThisInstall = true
		
        call fnStartLog(strLogName)
		
           fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
	       fnAddLog("SCRIPT INFO: fnInitialize - starting Install().")
    
                call fnPrintMachineDetails()
    
	       'fnAddLog(PreTxtFN(3) & "free HDD space: " & fnFreeHDDSpace() & DotFN(0))
	       fnAddLog(PreTxtFN(3) & "SCCM cache total space: " & fnSCCMCache("total") & ", SCCM free cache space: " & fnSCCMCache("free") & DotFN(0))
	
		If selfCopy = True Then
			call fnSubstituteSelf()
		End If
	
		if StrikeWindowActive = false and boolSilentInstall = false and PrecheckIsInstalled = true and blnFthStrikeActive = false Then	'AO - one account does not check fnIsIbmMsiPKGInstalled
			if fnIsIbmMsiPKGInstalled(strPkgGUID, strAppNumber, strPKGName) = 0 Then
				fnAddLog("FUNCTION INFO: ending fnSetInstall - PKG installed, exiting script.")
				fnFinalize(0)
			Else
				intRVal = fnRequestPortal()
			End If
			
				if(IsLangWindow = true and strLanguagesToUse <> "" and boolSilentInstall = false) Then 'checking if language window will be displayed
					fnAddLog("SCRIPT INFO: going to fnSelectionLanguageWindow.")
					strSelectedLang = fnSelectionLanguageWindow()
					if strSelectedLang = 0 Then 'beacause this function have proper exit <> 0 - this must be that way. 
						fnAddLog("FUNCTION INFO: ending fnSetInstall - langauge has not been chose.")
						fnFinalize(1)
					End if
				End if
				
				if(IsKillingWindow = true and strProcessessToKill <> "" and boolSilentInstall = false) Then
					fnAddLog("SCRIPT INFO: going to fnProcessWindowHTA.")
					intRVal = fnProcessWindowHTA()
				End if
				
			fnAddLog("SCRIPT INFO: going to fnProgressBarSet.")
			call fnProgressBarSet(boolThisInstall) 
		ElseIf  PrecheckIsInstalled = false Then																'AO 	
			fnAddLog("SCRIPT INFO: going to fnProgressBarSet.")
			call fnProgressBarSet(boolThisInstall) 	
		ElseIf blnFthStrikeActive = true Then
			intRVal = fnIsIbmMsiPKGInstalled(strPkgGUID, strAppNumber, strPKGName)
			
			if intRVal = 0 Then
				fnAddLog("FUNCTION INFO: ending fnSetInstall - PKG installed, exiting script.")
				boolFSilentInstall = true
				fnFinalize(0)
			End If
			
			If boolSilentInstall = false AND boolFSilentInstall = False Then
				intRVal = fnFthStrikeWindowHTA()
			End If
			
			If intRVal = -11 Then
				fnAddLog("SCRIPT INFO: ending script. Installation Postponed")
				fnFinalize(Const_StrikeEC)
			ElseIf intRVal = 1 Then
				fnFinalize(1)
			Else
				fnAddLog("SCRIPT INFO: going to next step.")
					intRVal = fnKillProcess(strProcessessToKill)
					If intRVal <> 0 Then
						fnFinalize(intRVal)
					End If
				If boolFSilentInstall = False Then
					call fnBannerWindowHTA()
				End If
			End If
		Else
			if fnIsIbmMsiPKGInstalled(strPkgGUID, strAppNumber, strPKGName) = 0 Then
				fnAddLog("FUNCTION INFO: ending fnSetInstall - PKG installed, exiting script.")
				fnFinalize(0)
			Else 
				intRVal = fnRequestPortal()
			End If
			
			fnAddLog("SCRIPT INFO: going to fnProcessesWindow.")
            if strFriendlyPKGName = "" or strFriendlyPKGName = "strFriendlyPKGName" Then
                strTempPKGName = strPKGName
            Else
                strTempPKGName = strFriendlyPKGName
            End If
        
			intStrikeVar = fnProcessesWindow(HowManyStrikesVar, AccountNameVar, TitleVar, intTimeVar, strTempPKGName, IsForceKillingVar, strFriendlyProcessName, IsRestartMessageVar)
		End If
		
		
		REM If blnFthStrikeActive = true Then
			REM intRVal = fnIsIbmMsiPKGInstalled(strPkgGUID, strAppNumber, strPKGName)
			REM msgbox intRVal
			
			REM if intRVal = 0 Then
				REM fnAddLog("FUNCTION INFO: ending fnSetInstall - PKG installed, exiting script.")
				REM fnFinalize(0)
			REM End If
			
			REM If boolSilentInstall = false AND boolFSilentInstall = False Then
				REM intRVal = fnFthStrikeWindowHTA()
			REM End If
			
			REM If intRVal = -11 Then
				REM fnAddLog("SCRIPT INFO: ending script. Installation Postponed")
				REM fnFinalize(Const_StrikeEC)
			REM ElseIf intRVal = 1 Then
				REM fnFinalize(1)
			REM Else
				REM fnAddLog("SCRIPT INFO: going to next step.")
					REM intRVal = fnKillProcess(strProcessessToKill)
					REM If intRVal <> 0 Then
						REM fnFinalize(intRVal)
					REM End If
				REM If boolFSilentInstall = False Then
					REM call fnBannerWindowHTA()
				REM End If
			REM End If
		REM End If
		
		
		if intStrikeVar = -2000 then
			fnAddLog("FUNCTION INFO: ending fnSetInstall - without fnProcessesWindow.")
		Elseif intStrikeVar = 0 Then
			fnAddLog("FUNCTION INFO: ending fnSetInstall.")
		Elseif intStrikeVar = Const_StrikeEC Then
			fnAddLog("FUNCTION INFO: ending fnSetInstall - fnProcessesWindow postponed installation.")
			fnFinalize(Const_StrikeEC)
		End If
	
	strLogTAB = strLogTABBak 
End Function
'-------------------------------------------------------------------------------------------
'! Set appropriate variables to uninstall mode.
'!
'! @return N/A
Function fnSetUninstall()
    Dim strTempPKGName, intRVal
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		strLogName = strPKGName & "_Uninstall_VBS.log" 
		boolThisInstall = false
		call fnStartLog(strLogName)
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
    
	fnAddLog("SCRIPT INFO: fnInitialize - starting Uninstall().") 
               
        call fnPrintMachineDetails()
    
        If StrikeWindowActiveUninstall = false and boolSilentInstall = false and UninstallProgressBarActive = true Then
			if fnIsIbmMsiPKGInstalled(strPkgGUID, strAppNumber, strPKGName) = 2 Then											
				fnAddLog("FUNCTION INFO: ending fnSetUninstall - PKG uninstalled, exiting script.")
				fnFinalize(0)
			End If
			
			if(IsUninstallKillingWindow = true and strProcessessToKill <> "" and boolSilentInstall = false) Then
				fnAddLog("SCRIPT INFO: going to fnProcessWindowHTA.")
				intRVal = fnProcessWindowHTA()
            End if
        
            fnAddLog("SCRIPT INFO: going to fnProgressBarSet.")
			call fnProgressBarSet(boolThisInstall) 
		ElseIf StrikeWindowActiveUninstall = false and boolSilentInstall = false and UninstallProgressBarActive = false Then						
			if fnIsIbmMsiPKGInstalled(strPkgGUID, strAppNumber, strPKGName) = 2 Then											
				fnAddLog("FUNCTION INFO: ending fnSetUninstall - PKG uninstalled, exiting script.")
				fnFinalize(0)
			End If
        
            if(IsUninstallKillingWindow = true and strProcessessToKill <> "" and boolSilentInstall = false) Then   'copied into top section, for progress bar and killing process window fix
				fnAddLog("SCRIPT INFO: going to fnProcessWindowHTA.")
				intRVal = fnProcessWindowHTA()
            End if
		Else
			if fnIsIbmMsiPKGInstalled(strPkgGUID, strAppNumber, strPKGName) = 2 Then											
				fnAddLog("FUNCTION INFO: ending fnSetUninstall - PKG uninstalled, exiting script.")
				fnFinalize(0)
			End If
			fnAddLog("SCRIPT INFO: going to fnProcessesWindow.")
            if strFriendlyPKGName = "" or strFriendlyPKGName = "strFriendlyPKGName" Then
                strTempPKGName = strPKGName
            Else
                strTempPKGName = strFriendlyPKGName
            End If
        
			call fnProcessesWindow(HowManyStrikesVar, AccountNameVar, TitleVar, intTimeVar, strTempPKGName, IsForceKillingVar, strFriendlyProcessName, IsRestartMessageVar)
		End If
		
		If blnFthStrikeActive = true Then
			intRVal = fnKillProcess(strProcessessToKill)
			If intRVal <> 0 Then
				fnFinalize(intRVal)
			End If
		End If
	fnAddLog("FUNCTION INFO: ending fnSetUninstall.")
	
	strLogTAB = strLogTABBak 
End Function
'-------------------------------------------------------------------------------------------
'! Set progress bar caption into proper mode (install/uninstall), check if display of it is allowed.
'!
'! @return N/A
Function fnProgressBarSet(boolInstMode)	
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		 
		fnAddLog("FUNCTION INFO: starting ProgressBarSetBarCaption with parameter - " & boolInstMode & ".") 
		
			If boolInstMode = true Then
				If ProgressBarActive Then
					fnAddLog("FUNCTION INFO: checking string parameter for progressbar, ini value = " & InstallProgressBarCaption & ".") 
                    If InStr(InstallProgressBarCaption,"<strFriendlyPKGName>") > 0 Then
					   'call fnShowProgressBar(InstallProgressBarTitle, Replace(InstallProgressBarCaption,"<strFriendlyPKGName>", strFriendlyPKGName))
					   call fnShowProgressBar(InstallProgressBarTitle, InstallProgressBarCaption)
                    Else
					   call fnShowProgressBar(strPKGName, InstallProgressBarCaption)
                    End If 
				End If
			Else
				If UninstallProgressBarActive Then
					fnAddLog("FUNCTION INFO: checking string parameter for progressbar, ini value = " & InstallProgressBarCaption & ".") 
                    If InStr(UninstallProgressBarCaption,"<strFriendlyPKGName>") > 0 Then
					   'call fnShowProgressBar(UninstallProgressBarTitle, Replace(UninstallProgressBarCaption,"<strFriendlyPKGName>", strFriendlyPKGName))
					   call fnShowProgressBar(UninstallProgressBarTitle, UninstallProgressBarCaption)
                    Else
					   call fnShowProgressBar(strPKGName, UninstallProgressBarCaption)
                    End If 
				End If
			End If
			
		fnAddLog("FUNCTION INFO: ending fnProgressBarSet.")
	
	strLogTAB = strLogTABBak 
	fnProgressBarSet = 0
End Function
'-------------------------------------------------------------------------------------------
'! Quit script if fail occurred. 
'!
'! @return N/A
Function fnFinalizeIfFail(strFunctionName, intRVal)
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
        fnAddLog(VarsFNErr(0) & "strFunctionName" & VarsFNErr(1))
        fnAddLog(EndingFN(0) & strFunctionName & SucFailFN(1))
	
	strLogTAB = strLogTABBak 
	
	fnFinalize(intRVal)
	strFunctionName = intRVal
End Function
'-------------------------------------------------------------------------------------------
'! Generate log entry about intRVal returned by function.
'!
'! @param  strCmd     		Command to run in CMD.
'! @param  strArgs          Additional CMD parameters for strCmd.
'!
'! @return 0,1,3010,intRVal
Function fnRunCmd(strCmd, strArgs)
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog("FUNCTION INFO: starting fnRunCmd with parameters: strCmd - " & strCmd & ", strArgs- " & strArgs & ".") 
			Dim CMDCommand, intRVal
			
				If FnParamChecker(strCmd, "strCmd") = 1 Then
			        call fnConst_VarFail ("strCmd", strCurrentFNName)
			
			        strLogTAB = strLogTABBak
                    fnFinalize(1)
                    fnRunCmd = 1
				End If
				
				If FnParamChecker(strArgs, "strArgs") = 1 Then
			        fnAddLog(VarsFNWar(0) & "strArgs" & VarsFNWar(1) & "empty string" & DotFN(0))
					strArgs = ""
				End If
				
				if strArgs <> "" Then
				    CMDCommand = chr(34) & strCmd & chr(34) & chr(32) & strArgs
                Else 
                    CMDCommand = chr(34) & strCmd & chr(34) 
                End If
				
				
					fnAddLog("INFO: CMDCommand: " & CMDCommand)
				intRVal = objShell.Run (CMDCommand, 0, true)
					fnAddLog("INFO: " & fnRunCmd & " installation ended with intRVal = " & intRVal & ".")

				If intRVal<>3010 And intRVal<>0 Then 
					fnAddLog("FUNCTION INFO: ending fnRunCmd - fail.") 	
	
					strLogTAB = strLogTABBak 	
					fnRunCmd = 1
				Else 
					fnAddLog("FUNCTION INFO: ending fnRunCmd.") 
					
					If intRVal = 3010 then boolRebootContainer = True
					
					strLogTAB = strLogTABBak 		
					fnRunCmd = 0
				End If	
End Function
'-------------------------------------------------------------------------------------------
Function fnGetMSOfficeVersions()
	Dim arrColSoft, strVersionO, strOffices
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		Set arrColSoft = objWMIService.ExecQuery("SELECT * FROM Win32_Product WHERE Name Like 'Microsoft Office%'")
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
			
		If arrColSoft.Count = 0 Then
			fnAddLog("FUNCTION INFO: Ending function fnGetMSOfficeVersions, with empty string.")
			
			strOffices = ""
			fnGetMSOfficeVersions = strOffices
			Exit Function
		Else
			For Each objItem In arrColSoft
				strVersionO = Left(objItem.Version, InStr(1,objItem.Version,".")-1)
				fnAddLog("FUNCTION INFO: " & objitem.caption & ", Version " & strVersionO)
					if strOffices <> "" Then
						strOffices = strVersionO & ";" & strOffices
					Else
						strOffices = strVersionO & ";"
					End If
			Next
		End If
	
		fnAddLog("FUNCTION INFO: Ending function fnGetMSOfficeVersions.")
		
		strLogTAB = strLogTABBak
		fnGetMSOfficeVersions = strOffices
End Function
'-------------------------------------------------------------------------------------------
Function fnGetDotNetVersions()
	Dim strDotNets, arrColSoft
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		Set arrColSoft = objWMIService.ExecQuery("SELECT * FROM Win32_Product WHERE Name Like '%%.NET%%'")
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
			
		If arrColSoft.Count = 0 Then
			fnAddLog("FUNCTION INFO: Ending function fnGetMSOfficeVersions, with empty string.")
			
			strDotNets = ""
			fnGetDotNetVersions = strDotNets
			Exit Function
		Else
			For Each objItem In arrColSoft
				fnAddLog("FUNCTION INFO: " & objitem.Name & ", Version " & objItem.Version)
					if strDotNets <> "" Then
						strDotNets = objItem.Version & ";" & strDotNets
					Else
						strDotNets = objItem.Version & ";"
					End If
			Next
		End If
	
		fnAddLog("FUNCTION INFO: Ending function fnGetDotNetVersions.")
		
		strLogTAB = strLogTABBak
		fnGetDotNetVersions = strDotNets
End Function
'-------------------------------------------------------------------------------------------
Function fnAllSoftwareList()
	Dim allSoft, intRet1, strValue1, strValue2, intValue3, intValue4, intValue5
	Const HKLM = &H80000002
	
	Dim strKey							: strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" 
	Dim strKey64						: strKey64 = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" 
	Dim strEntry1a						: strEntry1a = "DisplayName" 
	Dim strEntry1b						: strEntry1b = "QuietDisplayName" 
	Dim strEntry2						: strEntry2 = "InstallDate" 
	Dim strEntry3						: strEntry3 = "VersionMajor" 
	Dim strEntry4						: strEntry4 = "VersionMinor" 
	Dim strEntry5						: strEntry5 = "EstimatedSize" 
	
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
			
		if fnIs64BitOS(false) = 0 Then
			For i = 0 to 2
				if i = 0 Then
					strKey = strKey
					fnAddLog("FUNCTION INFO: 64bit listing from: " & strKey & ".")
				Else
					strKey = strKey64
					fnAddLog("FUNCTION INFO: 32bit listing from: " & strKey & ".")
				End If 
				
				objReg.EnumKey HKLM, strKey, arrSubkeys 	
			
				For Each strSubkey In arrSubkeys 
					intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1) 
					If intRet1 <> 0 Then 
						objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1 
					End If 
					
					objReg.GetStringValue HKLM, strKey & strSubkey, strEntry2, strValue2 
					
					objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry3, intValue3 
					objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry4, intValue4 
					
					objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry5, intValue5 
					
					if strValue1 <> "" And intValue3 <> "" Then 
						fnAddLog("FUNCTION INFO: " & strEntry1a & ": " & strValue1 & ", Version: " & intValue3 & "." & intValue4 & ".")
						if allSoft <> "" Then
							allSoft = intValue3 & "." & intValue4 & "^" & strValue1 & ";" & allSoft
						Else
							allSoft = intValue3 & "." & intValue4 & "^" & strValue1 & ";"
						End If
					End If
				Next 
				
				i = i + 1
			Next
		Else
				strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" 
				
				fnAddLog("FUNCTION INFO: 32bit listing from: " & strKey & ".")
					
				objReg.EnumKey HKLM, strKey, arrSubkeys 	
			
				For Each strSubkey In arrSubkeys 
					intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1) 
					If intRet1 <> 0 Then 
						objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1 
					End If 
					
					objReg.GetStringValue HKLM, strKey & strSubkey, strEntry2, strValue2 
					objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry3, intValue3 
					objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry4, intValue4 
					objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry5, intValue5 
					
					if strValue1 <> "" And intValue3 <> "" Then 
						fnAddLog("FUNCTION INFO: " & strEntry1a & ": " & strValue1 & ", Version: " & intValue3 & "." & intValue4 & ".")
						if allSoft <> "" Then
							allSoft = intValue3 & "." & intValue4 & "^" & strValue1 & ";" & allSoft
						Else
							allSoft = intValue3 & "." & intValue4 & "^" & strValue1 & ";"
						End If
					End If
				Next 
		End If
			
		fnAddLog("FUNCTION INFO: Ending function fnAllSoftwareList.")
		
		strLogTAB = strLogTABBak
		fnAllSoftwareList = allSoft
End Function
'-------------------------------------------------------------------------------------------
'! Check parameters passed to script in CMD.
'!
'! @return N/A
Function fnPreSetup()
		if fnIs64BitOS(false) = 0 and fnIs64BitShellEnv(false) = 1 Then
			if OneFileSetup = false Then
				If Wscript.Arguments.Count = 0 Then
					intRetCode = objShell.Run (strFolder & "\wscript.exe " & chr(34) & Wscript.ScriptFullName & chr(34), 0, True)
				Else
					intRetCode = objShell.Run (strFolder & "\wscript.exe " & chr(34) & Wscript.ScriptFullName & chr(34) & " " & Wscript.Arguments(0), 0, True)
				End If
			Else 
				if Wscript.Arguments.Count = 0 Then
					call fnNoFlags()													'Exiting script
				Elseif Wscript.Arguments.Count = 1 Then
					intRetCode = objShell.Run (strFolder & "\wscript.exe " & Chr(34) & Wscript.ScriptFullName & Chr(34) & " " & Wscript.Arguments(0), 0, True)
				Elseif Wscript.Arguments.Count = 2 Then																										' Add one additional condition
																																							' in order to be able to use more than
																																							' one parameter - DK
					intRetCode = objShell.Run (strFolder & "\wscript.exe " & Chr(34) & Wscript.ScriptFullName & Chr(34) & " " & Wscript.Arguments(0) & " " & Wscript.Arguments(1), 0, True)
				End If
			End If
			WScript.Quit(intRetCode)
		End If
End Function
'-------------------------------------------------------------------------------------------
'! CMD parameters checker.
'!
'! @return N/A
    Function fnCMDParams()
	Dim colArgs
		Set colArgs = WScript.Arguments.Named
       
        If colArgs.Exists("server") Then 
            forCitrix = true 
        End If
        If NOT colArgs.Exists("srs") Then 
            EntryRequestURL = ""
        End If
        If colArgs.Exists("s") Then 
            boolSilentInstall = true  
        End If  
		
        If colArgs.Exists("silent") Then 
            boolSilentInstall = true 
        End If
		
        If colArgs.Exists("fsilent") Then 
            boolFSilentInstall = true 
        End If
End Function
'-------------------------------------------------------------------------------------------
'! All actions triggered before installation/uninstallation.
'!
'! @return N/A
Function fnInitialize()
	Dim intRVal, colArgs
		Set colArgs = WScript.Arguments.Named
	
		if OneFileSetup = true Then						'If OneFileSetup then checking for arguments (flags)
            if Wscript.Arguments.Count = 0 Then
				call fnNoFlags()
            Else 
                
                call fnCMDParams()
                
                If colArgs.Exists("remove") or colArgs.Exists("uninstall") Then 
                    intRVal = fnSetUninstall()	
                End If
                If colArgs.Exists("install") or colArgs.Exists("setup") Then 
                    intRVal = fnSetInstall()
                End If	
            End If
		Else											'If no OneFileSetup then checking script Name
            if LCase(strScriptHostName) = "install.vbs" or LCase(strScriptHostName) = "setup.vbs" or ((InStr(strScriptHostName,chr(45) & "install.vbs"))<>0) or ((InStr(strScriptHostName,chr(95) & "install"))<>0) Then     
                
                call fnCMDParams()
                
                intRVal = fnSetInstall()
            Elseif LCase(strScriptHostName) = "uninstall.vbs" or LCase(strScriptHostName) = "remove.vbs" or ((InStr(strScriptHostName,chr(45) & "uninstall.vbs"))<>0) or ((InStr(strScriptHostName,chr(95) & "uninstall"))<>0) Then
                
                call fnCMDParams()
                
                intRVal = fnSetUninstall()   
            Else
                call fnNoFlags()	                                  
            End If
		End If
        
        
        if intRVal <> 0 Then
            fnAddLog("SETUP ERROR: fnSetInstall() exited with error. Exiting script with intRVal = 1.")
            fnFinalize(1)
        End If
        
        
        If forCitrix = True then							
            If fnIsServer(true) = False then
                fnAddLog("SETUP WARRNING: intended for server installations but the machine is NOT server OS.")
                fnFinalize(1)
            Else
                fnAddLog("SETUP INFO: It's a citrix setup, installation on server OS. Proceed with script.")
            End If
        else
            If fnIsServer(True) = True Then
                fnAddLog("SETUP INFO: ending fnInitialize.") 
                fnAddLog("SETUP WARRNING: exit script, machine is server OS.") 
                fnFinalize(Const_isServEC)
            Else
                fnAddLog("SETUP INFO: it's not a server machine. Proceed with script.")
            end If
        End If
        
        
	
	fnAddLog("SETUP INFO: ending fnInitialize.") 
           
            
	if boolThisInstall = true Then
		Install()										'Going to main Install function
	Else 
		Uninstall()										'Going to main Uninstall function
	End If
End Function
'-------------------------------------------------------------------------------------------
'! All actions triggered after installation/uninstallation if we have success or fail.
'!
'! @param  intTempRVal         Last script return code.
'!
'! @return N/A
Function fnFinalize(intTempRVal)
	Dim intRVal, strNKey, strBitness
	Dim intRVal 								: intRVal = 0
	
	fnAddLog("SETUP INFO: starting fnFinalize. Last return code: " & intTempRVal & ".") 
       
	If blnFthStrikeActive = true AND boolThisInstall = True Then
		If boolRebootContainer = True OR intTempRVal = 3010 Then
			boolRebootNeeded = True
		End If
		
		If boolFSilentInstall = false Then
			Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'mshta.exe'")
			
			For Each objProcess In colProcessList
				If InStr(objProcess.CommandLine, strWindowFile) Then
					objProcess.Terminate()
					fnAddLog("FUNCTION INFO: fnBannerWindowHTA was hidden.") 	
				End If
			Next
			
			If intTempRVal <> 3010 And intTempRVal <> 0 Then
			Else
				call CleanFthStrikeHTA()
				call fnCompletionDialogHTA()
			End IF
		End If
	End If	 
'----------------------------------------------------	
    If fnIs64BitOS(false) = 0 AND blnFthStrikeActive = true Then
		If fnCheckIfKeyExist("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strPkgGUID & "_" & strPKGName & "\") = 0 Then
			strBitness = ""
		ElseIf fnCheckIfKeyExist("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & strPkgGUID & "_" & strPKGName & "\") = 0 Then
			strBitness = "Wow6432Node\"
		Else 
			Exit Function
		End If
    ElseIf blnFthStrikeActive = true Then
		If fnCheckIfKeyExist("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strPkgGUID & "_" & strPKGName & "\") = 0 Then
			strBitness = ""
		Else 
			Exit Function
		End If
	End If
	
	strNKey = "HKEY_LOCAL_MACHINE\SOFTWARE\" & strBitness & "Microsoft\Windows\CurrentVersion\Uninstall\" & strPkgGUID & "_" & strPKGName
'-------------------------------------------------------
	If selfCopy = True AND boolThisInstall = True Then
		call fnSubstituteARP(strPkgGUID)
	ElseIf selfCopy = True AND boolThisInstall = False Then
		If fnCheckIfKeyExist(strNKey & "\") = 0 Then
			intRVal = fnRemoveKey(strNKey & "\")
		End If
		
		If fnIsExistD(strTempForPKGSrc & "\" & strPKGName) = 0 Then
			call fnRemoveDirectory(strTempForPKGSrc & "\" & strPKGName, False)
		End If
	End If
	
		
    If boolRebootContainer = True And intTempRVal <> 3010 And intTempRVal <> 0 Then
            fnAddLog("SETUP INFO: intTempRVal <> 3010 and 0, ignoring boolRebootContainer.")  
            boolRebootContainer = false
            no3010 = false
    End If 
        
	If boolRebootContainer = True then intTempRVal = 3010
	If intTempRVal = 3010 And no3010 = true Then intTempRVal = 0
                
		fnAddLog("SETUP INFO: ending fnFinalize.") 
		fnAddLog("SETUP INFO: exiting with intRVal= " & intTempRVal & ".") 
          
                
    If boolThisInstall = true Then
        fnHideProgressBar(InstallProgressBarTitle)
    Else
        fnHideProgressBar(UninstallProgressBarTitle)        
    End If
	
                
	Set objWMIService = Nothing
	Set colProcessList = Nothing
	
	If IsFailFinalizeMsg = True And intTempRVal <> 0 And intTempRVal <> Const_StrikeEC AND NOT strScriptHostName = "uninstall.vbs" Then		'MR - disabling fail msg popup in case of uninstall
		call fnPopupWindows("120", TitleVar, "Installation of " & strPKGName & " was unsuccessful." & vbCrLf & vbCrLf & "No action is required at this time." , "0", "48")  
		call fnRollback()
			fnAddLog("SETUP INFO: ending log, intRVal = " & intTempRVal) 
		fnStopLog(intTempRVal)
		wscript.quit intTempRVal								'Whole script exit code set as intTempRVal
	ElseIf intTempRVal <> 0 And intTempRVal <> Const_StrikeEC AND NOT strScriptHostName = "uninstall.vbs" Then 								'for fnRollback in case of all installations without finalize MSG
		call fnRollback()	
			fnAddLog("SETUP INFO: ending log, intRVal = " & intTempRVal) 
		fnStopLog(intTempRVal)
		wscript.quit intTempRVal								'Whole script exit code set as intTempRVal
	Else 
		fnStopLog(intTempRVal)
		wscript.quit intTempRVal								'Whole script exit code set as intTempRVal
	End If
End Function
'-------------------------------------------------------------------------------------------
'! Force explorer refresh.
'!
'! @return N/A
Function fnRefreshExplorer()
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
        
   
        Dim colProcess, objProcess 

            Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = 'explorer.exe'")

            For Each objProcess in colProcess
               objProcess.Terminate()
            Next 

			
		fnAddLog("FUNCTION INFO: Ending function fnRefreshExplorer.")
		
		strLogTAB = strLogTABBak
End Function 
'-------------------------------------------------------------------------------------------
'! Check if intLast return code is one of proper exit codes specified in arrRCList.
'!
'! @param  intRetCode          Last return code.
'! @param  arrRCList           Proper exit codes list.
'!
'! @return N/A
Function fnRetCodeChecker(intRetCode, arrRCList) 
    Dim CLRetCode, intRVal, tmpA, tmpC, i, STa, STb
    Dim strCurrentFNName                   : strCurrentFNName = "fnRetCodeChecker"
    Dim failVar                         : failVar = true
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "intRetCode - " & intRetCode & ", arrRCList - " & arrRCList & DotFN(0)) 
                
        If FnParamChecker(arrRCList, "arrRCList") = 1 Then
			fnAddLog(VarsFNErr(0) & "arrRCList" & VarsFNErr(1))
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))    ' change to SET new value - in log. not fail
			
            arrRCList = strSuccessTable 
            
            i = 0
		End If   
                
        If FnParamChecker(intRetCode, "intRetCode") = 1 Then
            call fnConst_VarFail("intRetCode", strCurrentFNName)  
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRetCodeChecker = 1
		End If      
                
        
            CLRetCode = Split (arrRCList, ",")     
            
            For Each tmpA In CLRetCode
                fnAddLog("FUNCTION INFO: Testing if intRetCode (" & intRetCode & ") = tmpA (" & tmpA & ").")
                If CStr(tmpA) = CStr(intRetCode) Then
                    failVar = false
                    fnAddLog("FUNCTION INFO: Setting failVar to " & failVar & ".")
                Else 
                    fnAddLog("FUNCTION INFO: Skipping change failVar.")  
                End If
            Next
                
            if failVar = true Then        
                fnAddLog("FUNCTION INFO: Ending function fnRetCodeChecker, intRVal for fail.")

                strLogTAB = strLogTABBak
                fnFinalize(intRetCode)
                fnRetCodeChecker = intRetCode
            Else   
                fnAddLog("FUNCTION INFO: Ending function fnRetCodeChecker, success.")

                strLogTAB = strLogTABBak
                fnRetCodeChecker = intRetCode
            End If
End Function 
'-------------------------------------------------------------------------------------------
'! Check if intLast character of specifid string is this same as specified char in secound varaible.
'!
'! @param  strValue        String to test.
'! @param  strChar          Char we are testing for.
'!
'! @return 0,1
Function fnLastCharChecker(strValue, strChar) 
    Dim CLRetCode, intRVal, tmpA
    Dim strCurrentFNName                   : strCurrentFNName = "fnRetCodeChecker"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strValue - " & strValue & ", strChar - " & strChar & DotFN(0)) 
                
        If FnParamChecker(strValue, "strValue") = 1 Then
            call fnConst_VarFail("strValue", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnLastCharChecker = 1
		End If   
                
        If FnParamChecker(strChar, "strChar") = 1 Then
            call fnConst_VarFail("strChar", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnLastCharChecker = 1
		End If      
                
            tmpA = Right(strValue, 1)    
            
            if tmpA = strChar Then
                intRVal = 0
            Else
                intRVal = 1   
            End If
                     
        fnAddLog(EndingFN(0) & strCurrentFNName & EndingFN(1) & intRVal & DotFN(0))
                
        strLogTAB = strLogTABBak
        fnLastCharChecker = intRVal
End Function 
'-------------------------------------------------------------------------------------------
'! Check if function parameter have value (is not a empty, null variable)
'!
'! @param  VarParam         Application Version from ARP.
'!
'! @return 0,1
Function FnParamChecker(VarParam, VarName) 
    Dim CLRetCode, intRVal, tmpA
    Dim strCurrentFNName                   : strCurrentFNName = "FnParamChecker"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & VarName & " - " & VarParam & DotFN(0)) 
                
        If fnIsNull(VarParam) = True or VarParam = "" or isEmpty(VarParam) = True Then
			fnAddLog(VarsFNErr(0) & VarName & VarsFNErr(2))
			
			intRVal = 1
        Else
            intRVal = 0 
		End If   
                
                     
        fnAddLog(EndingFN(0) & strCurrentFNName & EndingFN(1) & intRVal & DotFN(0))
                
        strLogTAB = strLogTABBak
        FnParamChecker = intRVal
End Function '-------------------------------------------------------------------------------------------
'! Clear stuff after 5th strike.
'!
'! @return N/A
Function ClearARP()
	Dim strBitness, strNKey, strOKey
	
	fnAddLog("SCRIPT INFO: cleaning ARP.")
	
'-------------------------------------------------------
	strBitness = ""
	
	strNKey = "HKEY_LOCAL_MACHINE\SOFTWARE\" & strBitness & "Microsoft\Windows\CurrentVersion\Uninstall\" & strPkgGUID & "_" & strPKGName
	strOKey = "HKEY_LOCAL_MACHINE\SOFTWARE\" & strBitness & "Microsoft\Windows\CurrentVersion\Uninstall\" & strPkgGUID
'-------------------------------------------------------
		If selfCopy = True Then
			If fnCheckIfKeyExist(strNKey & "\") = 0 Then
				call fnRemoveKey(strNKey & "\")
			End If
			
			If fnCheckIfKeyExist(strOKey & "\") = 0 Then
				call fnRemoveKey(strOKey & "\")
			End If
'-------------------------------------------------------
	strBitness = "Wow6432Node\"
	
	strNKey = "HKEY_LOCAL_MACHINE\SOFTWARE\" & strBitness & "Microsoft\Windows\CurrentVersion\Uninstall\" & strPkgGUID & "_" & strPKGName
	strOKey = "HKEY_LOCAL_MACHINE\SOFTWARE\" & strBitness & "Microsoft\Windows\CurrentVersion\Uninstall\" & strPkgGUID
'-------------------------------------------------------
			If fnCheckIfKeyExist(strNKey & "\") = 0 Then
				call fnRemoveKey(strNKey & "\")
			End If
			
			If fnCheckIfKeyExist(strOKey & "\") = 0 Then
				call fnRemoveKey(strOKey & "\")
			End If
			
			If fnIsExistD(strTempForPKGSrc & "\" & strPKGName) = 0 Then
				call fnRemoveDirectory(strTempForPKGSrc & "\" & strPKGName, False)
			End If
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Clear stuff after 5th strike.
'!
'! @return N/A
Function Clear5thStr()
	fnAddLog("SCRIPT INFO: cleaning 5thStr.")

	call ClearARP()
	call CleanFthStrikeHTA()
End Function
'-------------------------------------------------------------------------------------------