'****************************************************************************************************************************************************
' FUNCTIONS: 
'	fnShowProgressBar(strWindowTitle, strProgressBarCaption) 
'	fnHideProgressBar(strWindowTitle) 
' 	fnBannerWindowHTA()
'	
' FUNCTION EXAMPLE USAGE (IN and OUT):
' fnHideProgressBar "Package Name" 
'	Return 0

' fnShowProgressBar "Package Name", "Installation in progress. Please wait ..."
' fnShowProgressBar "Package Name", "Uninstallation in progress. Please wait ..."
'	Return 0

'***************************************************************************************************************************************************
'! Disable running VBS progress bar started by fnShowProgressBar(strWindowTitle, strProgressBarCaption).
'!
'! @param  strWindowTitle            Window strWindowTitle, agreed (it's format) per client, the same as strWindowTitle from running window.
'!
'! @return N/A
Function fnHideProgressBar(strWindowTitle) 
	Dim objProcess, strProcessName, strTempPKGName
    Dim strCurrentFNName                   : strCurrentFNName = "fnHideProgressBar"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		'strProcessName = "'mshta.exe'"
		strProcessName = "'ProgressBar.exe'"
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strWindowTitle - " & strWindowTitle & DotFN(0)) 
		
		strTempPKGName = strPKGName
		
		If (strProgressBarCaption = InstallProgressBarCaption AND ProgressBarActive = false) Or (strProgressBarCaption = UninstallProgressBarCaption AND UninstallProgressBarActive = false) Then
			fnAddLog("FUNCTION INFO: ending fnHideProgressBar - ProgressBarActive/UninstallProgressBarActive -> false.") 
			
			strLogTAB = strLogTABBak
			fnHideProgressBar = 0
			Exit Function
		End If 
		
		Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = " & strProcessName) 
    
                if strWindowTitle = "strFriendlyPKGName" Then
				    fnAddLog("FUNCTION INFO: variable strWindowTitle is string strFriendlyPKGName. Setting to strFriendlyPKGName.")
				    strWindowTitle = strFriendlyPKGName
                ElseIf strWindowTitle = "strPKGName" Then
			         fnAddLog("FUNCTION INFO: variable strWindowTitle is string strPKGName. Setting to strPKGName.")
			         strWindowTitle = strPKGName
                End If
				
				If InStr(strWindowTitle, "<strFriendlyPKGName>") > 0 Then
					If strFriendlyPKGName <> "" Then
						strWindowTitle = Replace(strWindowTitle, "<strFriendlyPKGName>", strFriendlyPKGName)
					Else
						strWindowTitle = Replace(strWindowTitle, "<strFriendlyPKGName>", strPKGName)
					End If
                End If
				
    
		If FnParamChecker(strWindowTitle, "strWindowTitle") = 1 Then
			     fnAddLog(VarsFNErr(0) & "strWindowTitle" & VarsFNErr(1))
			         
                    strWindowTitle = " "
        
			     fnAddLog(SetVarFN(0) & "strWindowTitle" & SetVarFN(1) & "empty string" & DotFN(0))
		End If
		
		For Each objProcess In colProcessList
			If InStr(objProcess.CommandLine, strWindowTitle) Then
				objProcess.Terminate()
				fnAddLog("FUNCTION INFO: ProgressBar was hidden.") 	
			End If
		Next
		
		fnAddLog("FUNCTION INFO: ending fnHideProgressBar.") 	
		strPKGName = strTempPKGName
		
		strLogTAB = strLogTABBak
		fnHideProgressBar = 0
End Function
'-------------------------------------------------------------------------------------------
'! Enable VBS progress bar for package.
'!
'! @param  strWindowTitle            Window strWindowTitle, agreed (it's format) per client.
'! @param  strProgressBarCaption Main text information aboout instaling package.
'!
'! @return 0,1
Function fnShowProgressBar(strWindowTitle, strProgressBarCaption) 
	Dim intRVal, strWindowFile, strExeDir, strExeCMD, strTempPKGName ',strHtaCMD, strHtaDir
    Dim strCurrentFNName                   : strCurrentFNName = "fnShowProgressBar"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		'strHtaDir = strScriptDir & "\Include\HTA"
		strExeDir = strScriptDir & "\Include\arrEXE"
		'strWindowFile = "ProgressBar.hta"
		strWindowFile = "ProgressBar.exe"
		intRVal = 0
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strWindowTitle - " & strWindowTitle & ", strProgressBarCaption - " & strProgressBarCaption & DotFN(0)) 
			
			'intRVal = fnIsExist(strWindowFile, strHtaDir) 
			intRVal = fnIsExist(strWindowFile, strExeDir) 
			strTempPKGName = strPKGName
	
			If (strProgressBarCaption = InstallProgressBarCaption AND ProgressBarActive = false) Or (strProgressBarCaption = UninstallProgressBarCaption AND UninstallProgressBarActive = false) Then
				fnAddLog("FUNCTION INFO: ending fnShowProgressBar - ProgressBarActive/UninstallProgressBarActive -> false.") 
				
				strLogTAB = strLogTABBak
				fnShowProgressBar = 0
				Exit Function
			End If
		
			If intRVal = 1 Then
				'fnAddLog("FUNCTION ERROR: HTA file missing or bad path.")
				fnAddLog("FUNCTION ERROR: arrEXE file missing or bad path.")
				fnAddLog("FUNCTION ERROR: ending fnShowProgressBar, intRVal: " & intRVal & ".")
				
				strLogTAB = strLogTABBak
				fnShowProgressBar = intRVal
			End If
			
			
			
			
				If InStr(InstallProgressBarTitle, "<strFriendlyPKGName>") > 0 Then
					If strFriendlyPKGName <> "" Then
						strWindowTitle = Replace(InstallProgressBarTitle, "<strFriendlyPKGName>", strFriendlyPKGName)
					Else
						strWindowTitle = Replace(InstallProgressBarTitle, "<strFriendlyPKGName>", strPKGName)
					End If
                End If
				
				
                    'if strWindowTitle = "strFriendlyPKGName" Then
                    '    fnAddLog("FUNCTION INFO: variable strWindowTitle is string strFriendlyPKGName. Setting to strFriendlyPKGName.")
                    '    strWindowTitle = strFriendlyPKGName
                    'ElseIf strWindowTitle = "strPKGName" Then
                    '     fnAddLog("FUNCTION INFO: variable strWindowTitle is string strPKGName. Setting to strPKGName.")
                    '     strWindowTitle = strPKGName
                    'End If


            If FnParamChecker(strWindowTitle, "strWindowTitle") = 1 Then
			         fnAddLog(VarsFNErr(0) & "bitness" & VarsFNErr(1))
			         
                        strWindowTitle = " "
        
			         fnAddLog(SetVarFN(0) & "strWindowTitle" & SetVarFN(1) & "empty string" & DotFN(0))
            End If
					
    
			If FnParamChecker(strProgressBarCaption, "strProgressBarCaption") = 1 Then
			         fnAddLog(VarsFNErr(0) & "strProgressBarCaption" & VarsFNErr(1))
			         
                        strProgressBarCaption = " "
        
			         fnAddLog(SetVarFN(0) & "strWindowTitle" & SetVarFN(1) & "empty string" & DotFN(0))
			End If
    
    
    
    
                If InStr(strProgressBarCaption,"<strFriendlyPKGName>") > 0 Then
                    strProgressBarCaption = Replace(strProgressBarCaption,"<strFriendlyPKGName>", strFriendlyPKGName)
                End If
				
                If InStr(strProgressBarCaption,"<strFriendlyProcessName>") > 0 Then
                    strProgressBarCaption = Replace(strProgressBarCaption,"<strFriendlyProcessName>", strFriendlyProcessName)
                End If
		
		
			
			'strHtaCMD = "mshta.exe " & chr(34) & strHtaDir & "\" & strWindowFile & chr(34) & " " & chr(34) & strProgressBarCaption & chr(34) & " " & chr(34) & strAccountName & chr(34) & " " & chr(34) & strWindowTitle & chr(34)
			strExeCMD = chr(34) & strExeDir & "\" & strWindowFile & chr(34) & " /" & chr(34) & strProgressBarCaption & chr(34) & "/" & chr(34) & strAccountName & chr(34) & "/" & chr(34) & strWindowTitle & chr(34) & "/" & chr(34) & PBMode & chr(34)
			
				'call objShell.Exec (strHtaCMD)
				call objShell.Exec (strExeCMD)
		
			'fnAddLog("FUNCTION INFO: fnShowProgressBar, strHtaCMD: " & strHtaCMD & ".")
			fnAddLog("FUNCTION INFO: fnShowProgressBar, strExeCMD: " & strExeCMD & ".")
			
			strPKGName = strTempPKGName
			
		fnAddLog("FUNCTION INFO: ending fnShowProgressBar.")
		strLogTAB = strLogTABBak
	fnShowProgressBar = intRVal
End Function
'-------------------------------------------------------------------------------------------
'! Disable/enable simple banner for installed package.
'!
'! @return intRVal from fnKillProcess (disable Banner.hta). 0,1 (enable Banner.hta)
Function fnBannerWindowHTA()
	'Dim intRVal, colProcessList, strWindowFile, strExeCMD, strExeDir ',strHtaCMD, strHtaDir
	Dim intRVal, colProcessList, strWindowFile, strHtaCMD, strHtaDir
    Dim strCurrentFNName                   : strCurrentFNName = "fnBannerWindowHTA"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		strHtaDir = strScriptDir & "\Include\HTA"
		strWindowFile = "Banner.hta"
		intRVal = 0
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
			
			intRVal = fnIsExist(strWindowFile, strHtaDir) 
	
			if intRVal = 1 Then
				fnAddLog("FUNCTION ERROR: HTA file missing or bad path.")
				fnAddLog("FUNCTION ERROR: ending fnBannerWindowHTA, intRVal: " & intRVal & ".")
				
				strLogTAB = strLogTABBak
				fnBannerWindowHTA = intRVal
			End If
			
			Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'mshta.exe'")
			
			For Each objProcess In colProcessList
				If InStr(objProcess.CommandLine, strWindowFile) Then
					objProcess.Terminate()
					fnAddLog("FUNCTION INFO: fnBannerWindowHTA was hidden.") 	
					
					fnAddLog("FUNCTION INFO: ending fnBannerWindowHTA, intRVal: " & 0 & ".")
					strLogTAB = strLogTABBak
					fnBannerWindowHTA = 0
					Exit Function
				End If
			Next
			
			strHtaCMD = "mshta.exe " & chr(34) & strHtaDir & "\" & strWindowFile & chr(34) & " " & chr(34) & AccountNameVar & chr(34) & " " & chr(34) & strFriendlyPKGName & chr(34) & " " & chr(34) & strAppNumber & chr(34)	
			
				call objShell.Exec (strHtaCMD)
			
		fnAddLog("FUNCTION INFO: ending fnBannerWindowHTA, intRVal: " & intRVal & ".")
		strLogTAB = strLogTABBak
	fnBannerWindowHTA = intRVal
End Function
'-------------------------------------------------------------------------------------------