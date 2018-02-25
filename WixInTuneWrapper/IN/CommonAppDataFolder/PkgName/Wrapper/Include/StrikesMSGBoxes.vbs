'****************************************************************************************************************************************************
' FUNCTIONS: 
'	fnSelectionLanguageWindow()
' 	fnProcessWindowHTA()
' 	fnFthStrikeWindowHTA()
' 	fnCompletionDialogHTA()
'	fnPopupWindows(Time, strWindowTitle, strMessage, intButton, intIcon)
'	fnProcessesWindow(intHowManyStrikes, strAccountName, strWindowTitle, Time,  strWhatInstalls, boolIsForceKilling, strProcessesToCloseVar, boolIsRestartMessage)
'	
' FUNCTION EXAMPLE USAGE (IN and OUT):
''-------------------------------------------------------------------------------------------
' 	fnPopupWindows(Time, strWindowTitle, strMessage, intButton, intIcon)
'	EXAMPLE: 
'			fnPopupWindows(0, "LEO IT Communication", "Displayed message", 0, 32)
'			fnPopupWindows(60, "LEO IT Communication", "Displayed message", NULL, NULL)
' 		
'-------------------------------------------------------------------------------------------
'	fnProcessesWindow(intHowManyStrikes, strAccountName, strWindowTitle, Time,  strWhatInstalls, boolIsForceKilling, strProcessesToCloseVar, boolIsRestartMessage)
'	EXAMPLE: 
'		fnProcessesWindow(3, "Astellas", "Astellas IT Communication", 7200, "Oracle_JavaRuntimeEnvironment_7.0.510_AST_W7_EN_B1" , true, "Internet Browser and Java", false)
'		fnProcessesWindow(3, "Astellas", "Astellas IT Communication", 7200, "Astellas_USComplianceGuide_2014.04_AST_W7_EN_B1" , NULL, NULL, NULL)
'			
'-------------------------------------------------------------------------------------------
'***************************************************************************************************************************************************
'! Display window with supported languages list by package.
'!
'! @return 1-n - Index of selected language.
Function fnSelectionLanguageWindow()
	Dim intRVal, strHtaDir, strWindowFile
    Dim strCurrentFNName                   : strCurrentFNName = "fnSelectionLanguageWindow"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		strHtaDir = strScriptDir & "\Include\HTA"
		strWindowFile = "LanguageWindow.hta"
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
			
			intRVal = fnIsExist(strWindowFile, strHtaDir) 
	
			if intRVal = 1 Then
				fnAddLog("FUNCTION ERROR: HTA file missing or bad path.")
				fnAddLog("FUNCTION ERROR: ending fnSelectionLanguageWindow, intRVal: " & intRVal & ".")
				
				strLogTAB = strLogTABBak
				fnSelectionLanguageWindow = intRVal
			End If
			
			Do
				intRVal = objShell.Run (chr(34) & strHtaDir & "\" & strWindowFile & " " & chr(34) & chr(34) & strLanguagesToUse & chr(34) & " " & chr(34) & AccountNameVar & chr(34) & " " & chr(34) & strFriendlyPKGName & chr(34), 0, true)
			Loop while intRVal = 0
			
		fnAddLog("FUNCTION INFO: ending fnSelectionLanguageWindow, intRVal: " & intRVal & ".")
		strLogTAB = strLogTABBak
	fnSelectionLanguageWindow = intRVal
End Function
'-------------------------------------------------------------------------------------------
'! Display window with information about running processes that need to be killed before installation/uninstallation.
'!
'! @return intRVal from fnKillProcess. fnProcessWindowHTA don't have intRVal for itself. 
Function fnProcessWindowHTA()
	Dim intRVal, strHtaDir, strWindowFile, intUninstString, fromDateH, intVTimer
    Dim strCurrentFNName                   : strCurrentFNName = "fnProcessWindowHTA"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		strHtaDir = strScriptDir & "\Include\HTA"
		
		If StrikeWindowIsPersistent = False Then
			strWindowFile = "Processes.hta"
		Else 
			strWindowFile = "ProcessesPersistent.hta"
			fromDateH = Now
		End If
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
		
            If boolThisInstall = false Then
                fnAddLog("FUNCTION INFO: setting as Uninstall.")
                intUninstString = "uninstall"
            Else
                fnAddLog("FUNCTION INFO: setting as Install.")
                intUninstString = "install"
            End If
        
			intRVal = fnIsExist(strWindowFile, strHtaDir) 
	
			if intRVal = 1 Then
				fnAddLog("FUNCTION ERROR: HTA file missing or bad path.")
				fnAddLog("FUNCTION ERROR: ending fnProcessWindowHTA, intRVal: " & intRVal & ".")
				
				strLogTAB = strLogTABBak
				fnProcessWindowHTA = intRVal
			End If
        
            If strFriendlyPKGName = "" Then
                strFriendlyPKGName = strPKGName
            End If
        
            If strFriendlyProcessName = "" Then
                strFriendlyProcessName = strProcessessToKill
            End If
	
            if fnIsProcessRunning(strProcessessToKill) = 1 Then
                fnAddLog("FUNCTION INFO: ending fnProcessWindowHTA, now running processes intRVal: 0.")
		        strLogTAB = strLogTABBak
	            fnProcessWindowHTA = 0
                Exit Function
            End If
        
			Do
                fnAddLog("FUNCTION INFO: running hta as:" & chr(34) & strHtaDir & "\" & strWindowFile & " " & chr(34) & chr(34) & intProcessHTAWindowTime & chr(34) & " " & chr(34) & AccountNameVar & chr(34) & " " & chr(34) & strFriendlyPKGName & chr(34) & " " & chr(34) & strFriendlyProcessName & chr(34) & " " & chr(34) & intUninstString & chr(34))
            
				If StrikeWindowIsPersistent = False Then
					intRVal = objShell.Run (chr(34) & strHtaDir & "\" & strWindowFile & " " & chr(34) & chr(34) & intProcessHTAWindowTime & chr(34) & " " & chr(34) & AccountNameVar & chr(34) & " " & chr(34) & strFriendlyPKGName & chr(34) & " " & chr(34) & strFriendlyProcessName & chr(34) & " " & chr(34) & intUninstString & chr(34), 0, true)
				Else 
					intRVal = objShell.Run (chr(34) & strHtaDir & "\" & strWindowFile & " " & chr(34) & chr(34) & intProcessHTAWindowTime & chr(34) & " " & chr(34) & AccountNameVar & chr(34) & " " & chr(34) & strAppName & chr(34) & " " & chr(34) & strFriendlyProcessName & chr(34) & " " & chr(34) & intUninstString & chr(34), 0, true)
				End If
				
				If StrikeWindowIsPersistent = true Then
					if fnIsProcessRunning(strProcessessToKill) = 1 Then
						fnAddLog("FUNCTION INFO: ending fnProcessWindowHTA, now running processes intRVal: 0.")
						strLogTAB = strLogTABBak
						fnProcessWindowHTA = 0
						Exit Function
					Else
						intVTimer = TimerCheck(fromDateH, Now)
						If intVTimer >= 11 Then
							fnFinalize(Const_toRetryEC)
						Else 
							intRVal = 0
						End If
					End If
				End If
			Loop while intRVal = 0
				
				If IsForceKillingVar = true Then
					intRVal = fnKillProcess(strProcessessToKill)
				Else
					fnAddLog("FUNCTION INFO: skipping fnKillProcess - client doesn't support it.")
				End If
			
		call Const_FnEndrVal(strCurrentFNName, intRVal)
		strLogTAB = strLogTABBak
	fnProcessWindowHTA = intRVal
End Function
'-------------------------------------------------------------------------------------------
'! Display window with informations about Package: running processes, package Name etc. It's providing timer and counter for deferring installation.
'!
'! @return intRVal from fnKillProcess. fnFthStrikeWindowHTA don't have intRVal for itself.
Function fnFthStrikeWindowHTA()
	Dim intRVal, strHtaDir, strWindowFile, intStrikesNumber, intWindowType, strTimerVar, intStrikesNumberREG, boolRebootWarning, boolLogoffWarning, strRegStrikeCountPath, strRegTimer, strBitness, strVarBit, strCommandHTA, strTmpValue
    Dim strCurrentFNName                   : strCurrentFNName = "fnFthStrikeWindowHTA"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		strHtaDir = strScriptDir & "\Include\HTA"
		strWindowFile = "5thStrikeWindow.hta"
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
			
			intRVal = fnIsExist(strWindowFile, strHtaDir) 
	
			if intRVal = 1 Then
				fnAddLog("FUNCTION ERROR: HTA file missing or bad path.")
				fnAddLog("FUNCTION ERROR: ending fnFthStrikeWindowHTA, intRVal: " & intRVal & ".")
				
				strLogTAB = strLogTABBak
				fnFthStrikeWindowHTA = intRVal
			End If
			
			If fnIs64BitOS(false) = 0 Then 
				strBitness = "\Wow6432Node"
				strVarBit = "64"
			Else
				strBitness = ""
				strVarBit = "32"
			End If
			
			strRegStrikeCountPath = "HKEY_LOCAL_MACHINE\SOFTWARE" & strBitness & "\IBM\Packages\" & strPKGName & "\strikecount"
			strRegTimer = strPKGName
			
			If fnCheckIfKeyExist(strRegStrikeCountPath) = 0 Then
				intStrikesNumberREG = fnReadKey(strRegStrikeCountPath)
			Else 
				intStrikesNumberREG = 0
				intRVal = fnAddKey(strRegStrikeCountPath, "0", "REG_SZ")
			End If
			
			
			intStrikesNumber = CInt(HowManyStrikesVar) - CInt(intStrikesNumberREG)

			strTimerVar = CStr(intTimeVar*60)
			
			boolRebootWarning = boolRebootNeeded
			boolLogoffWarning = boolLogOffNeeded
			
			
			If intStrikesNumber = 0 Then
				intWindowType = 7 'intLast window
				fnAddLog("FUNCTION INFO: intLast 5th strike window.")
				
				strTimerVar = "600"
			Else
				If strProcessessToKill <> "" Then
					intRVal = fnIsProcessRunning(strProcessessToKill)
					
					If boolRebootWarning = True Then 
						If intRVal = 0 Then
							intWindowType = 4
							fnAddLog("FUNCTION INFO: standard 5th strike with reboot info and with process info.")
						Else 
							intWindowType = 3
							fnAddLog("FUNCTION INFO: standard 5th strike with reboot info.")
						End If
					ElseIf boolLogoffWarning = True Then
						If intRVal = 0 Then
							intWindowType = 6
							fnAddLog("FUNCTION INFO: standard 5th strike with log off info and with process info.")
						Else 
							intWindowType = 5
							fnAddLog("FUNCTION INFO: standard 5th strike with log off.")
						End If
					Else
						If intRVal = 0 Then
							intWindowType = 2
							fnAddLog("FUNCTION INFO: standard 5th strike with process info.")
						Else 
							intWindowType = 1
							fnAddLog("FUNCTION INFO: standard 5th strike.")
						End If
					End If
				Else
					If boolRebootWarning = True Then 
						intWindowType = 3 
						fnAddLog("FUNCTION INFO: standard 5th strike with reboot info.")
					ElseIf boolLogoffWarning = True Then
						intWindowType = 5
						fnAddLog("FUNCTION INFO: standard 5th strike with log off.")
					Else
						intWindowType = 1 
						fnAddLog("FUNCTION INFO: standard 5th strike.")
					End If
				End If
			End If
			
			call fnChangeKey("HKEY_LOCAL_MACHINE\SOFTWARE" & strBitness & "\IBM\Packages\" & strRegTimer & "\remaningTime", strTimerVar)
			
			Do
				strTmpValue = fnReadKey("HKEY_LOCAL_MACHINE\SOFTWARE" & strBitness & "\IBM\Packages\" & strRegTimer & "\remaningTime")
				if strTimerVar < strTmpValue OR fnIsNull(strTmpValue) OR isEmpty(strTmpValue) Then
					if intWindowType = 7 Then
						strTimerVar = "600"
					Else 
						strTimerVar = CStr(intTimeVar*60)
					End If
				Else 
					strTimerVar = strTmpValue
				End If
				
				strCommandHTA = chr(34) & strHtaDir & "\" & strWindowFile & " " & chr(34) & " " & chr(34) & AccountNameVar & chr(34) & " " & chr(34) & strFriendlyPKGName & chr(34) & " " & chr(34) & strAppNumber & chr(34) & " " & chr(34) & intStrikesNumber & chr(34) & " " & chr(34) & intWindowType & chr(34) & " " & chr(34) & strFriendlyProcessName & chr(34) & " " & chr(34) & strTimerVar & chr(34) & " " & chr(34) & strRegTimer & chr(34) & " " & chr(34) & strVarBit & chr(34)
			
				fnAddLog("FUNCTION INFO: starting: " & strCommandHTA & ".")
				intRVal = objShell.Run (strCommandHTA, 0, true)
					If intRVal = -11 Then
						fnAddLog("FUNCTION INFO: ending fnFthStrikeWindowHTA, intRVal: " & intRVal & ".")
						
						If fnCheckIfKeyExist(strRegStrikeCountPath) = 0 Then
							intStrikesNumberREG = fnReadKey(strRegStrikeCountPath)
							intStrikesNumberREG = CInt(intStrikesNumberREG)
							call fnChangeKey(strRegStrikeCountPath, intStrikesNumberREG + 1)
						End If
						
						call fnRemoveKey ("HKEY_LOCAL_MACHINE\SOFTWARE" & strBitness & "\IBM\Packages\" & strPKGName & "\remaningTime")
						
						strLogTAB = strLogTABBak
						fnFthStrikeWindowHTA = intRVal
						Exit Function
					ElseIf intRVal <> 0 Then
						fnAddLog("FUNCTION INFO: ending fnFthStrikeWindowHTA, intRVal: " & intRVal & ".")
					
						call fnRemoveKey ("HKEY_LOCAL_MACHINE\SOFTWARE" & strBitness & "\IBM\Packages\" & strPKGName & "\remaningTime")
					
						strLogTAB = strLogTABBak
						fnFthStrikeWindowHTA = intRVal
						Exit Function
					End If
			Loop while intRVal = 0 or fnReadKey("HKEY_LOCAL_MACHINE\SOFTWARE" & strBitness & "\IBM\Packages\" & strRegTimer & "\remaningTime") <> 0
			
		fnAddLog("FUNCTION INFO: ending fnFthStrikeWindowHTA, intRVal: " & intRVal & ".")
		strLogTAB = strLogTABBak
	fnFthStrikeWindowHTA = intRVal
End Function
'-------------------------------------------------------------------------------------------
            '! Display window with simple information after installation.
'!
'! @return intRVal from fnKillProcess. fnProcessWindowHTA don't have intRVal for itself. 
Function fnCompletionDialogHTA()
	Dim strHtaCMD, strHtaDir, strWindowFile, intWindowType, intRVal
	Dim strCurrentFNName                   : strCurrentFNName = "fnCompletionDialogHTA"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		strHtaDir = strScriptDir & "\Include\HTA"
		strWindowFile = "CompletionDialog.hta"
		intRVal = 0
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
	
			intRVal = fnIsExist(strWindowFile, strHtaDir) 
	
			if intRVal = 1 Then
				fnAddLog("FUNCTION ERROR: HTA file missing or bad path.")
				fnAddLog("FUNCTION ERROR: ending fnCompletionDialogHTA, intRVal: " & intRVal & ".")
				
				strLogTAB = strLogTABBak
				fnCompletionDialogHTA = intRVal
			End If
			
				If boolRebootNeeded = True Then
					intWindowType = 2
				ElseIf boolLogOffNeeded = True Then
					intWindowType = 3
				Else
					intWindowType = 1
				End If
			
					strHtaCMD = "mshta.exe " & chr(34) & strHtaDir & "\" & strWindowFile & chr(34) & " " & chr(34) & AccountNameVar & chr(34) & " " & chr(34) & strFriendlyPKGName & chr(34) & " " & chr(34) & strAppNumber & chr(34) & " " & chr(34) & intWindowType & chr(34)
					
				call objShell.Exec (strHtaCMD)
			
		fnAddLog("FUNCTION INFO: ending fnCompletionDialogHTA.")
		strLogTAB = strLogTABBak
	fnCompletionDialogHTA = intRVal
End Function'-------------------------------------------------------------------------------------------
'! Clean 5th strike registry after installation.
'!
'! @return N/A
Function CleanFthStrikeHTA()
	If fnIs64BitOS(false) = 0 Then 
		call fnRemoveKey ("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\IBM\Packages\" & strPKGName & "\")
	Else
		call fnRemoveKey ("HKEY_LOCAL_MACHINE\SOFTWARE\IBM\Packages\" & strPKGName & "\")
	End If
End Function
'-------------------------------------------------------------------------------------------
'! Compare start and end time provided as parameters.
'!
'! @param  fromDate      	Start date with time.
'! @param  toDate      		Current date with current time.
'!
'! @return deltaTime 
Function TimerCheck(fromDate, toDate)
	Dim VTime
		fnAddLog("FUNCTION INFO: fromDate - " & fromDate & " toDate - " & toDate & ".")
		
			VTime = DateDiff("h",fromDate,toDate)
		
		fnAddLog("FUNCTION INFO: elapsed time: VTime - " & VTime & " h.")
	TimerCheck = VTime
End Function
'-------------------------------------------------------------------------------------------
'! Simple information window, can be configured in many ways to catch specific input from users.
'!
'! @param  Time             Receives an integer containing the number of seconds that the box will display for until dismissed. If zero or omitted, the message box stays until the user dismisses it. Default value: 0
'! @param  strWindowTitle            Receives a string containing the title that will appear as the title of the message box. Default value - empty string
'! @param  strMessage          Receives a string containing the message that will be displayed to the user in the message box. Default value- empty string
'! @param  intButton           Receives an integer that corresponds to displayed buttons, Default value - 0 Values that can be used: 0 (OK), 1 (OK, Cancel), 2 (Abort, Ignore, Retry), 3 (Yes, No, Cancel), 4 (Yes, No), 5 (Retry, Cancel)
'! @param  intIcon             Receives an integer that corresponds to a look. Default value - NULL. Values that can be used: 16 (Critical), 32 (Question), 48 (Exclamation), 64 (Information)
'!
'! @return Return the button which the user clicked to dismiss the pop-up message box. ReturnValue: 1 (OK), 2 (Cancel), 3 (Abort), 4 (Retry), 5 (Ignore), 6 (Yes), 7 (No), -1 (None, message box was dismissed automatically (timeout))        
Function fnPopupWindows(Time, strWindowTitle, strMessage, intButton, intIcon)
	if boolSilentInstall = false Then
		Dim intShowMessage, boolButtonSelect, boolIconSelect
        Dim strCurrentFNName                   : strCurrentFNName = "fnPopupWindows"
		Dim strLogTABBak						: strLogTABBak = strLogTAB
			strLogTAB = strLogTAB & strLogSeparator
            
            fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "Time - " & Time & ", strWindowTitle - " & strWindowTitle & ", strMessage - " & strMessage & ", intButton - " & intButton & ", intIcon - " & intIcon & DotFN(0))
			
			boolButtonSelect = False
			boolIconSelect = False
			
			If FnParamChecker(Time, "Time") = 1 Then
				fnAddLog("FUNCTION INFO: variable Time is NULL. Setting to 0 value.")
				Time = 0
			End If
			
			If FnParamChecker(strWindowTitle, "strWindowTitle") = 1 Then
				fnAddLog("FUNCTION INFO: variable strWindowTitle is NULL. Setting to empty string.")
				strWindowTitle = " "
			End If
			
			If FnParamChecker(strMessage, "strMessage") = 1 Then
				fnAddLog("FUNCTION INFO: variable strMessage is NULL. Setting to empty string.")
				strMessage = " "
			End If
			
			If FnParamChecker(intButton, "intButton") = 1 Then
				fnAddLog("FUNCTION INFO: variable intButton is NULL.")
				boolButtonSelect = False
			End If
			
			If FnParamChecker(intIcon, "intIcon") = 1 Then
				fnAddLog("FUNCTION INFO: variable intIcon is NULL.")
				boolIconSelect = False
			End If
			
			If intButton=0 or intButton=1 or intButton=2 or intButton=3 or intButton=4 or intButton=5 Then 
				boolButtonSelect = True
			Else
				boolButtonSelect = False
				If fnIsNull(intButton) = False or intButton <> "" Then
					fnAddLog("FUNCTION INFO: variable intButton is wrong.")
				End If
			End If
			
			If intIcon=16 or intIcon=32 or intIcon=48 or intIcon=64 Then 
				boolIconSelect = True
			Else
				boolIconSelect = False
				If fnIsNull(intIcon) = False or intIcon <> "" Then
					fnAddLog("FUNCTION INFO: variable intIcon is wrong.")
				End If
			End If
			'show message
			If boolIconSelect = True And boolButtonSelect = True Then 
				intShowMessage = wshShell.Popup(strMessage, Time, strWindowTitle, intButton + intIcon)	
				fnAddLog("FUNCTION INFO: Display strMessage.")
			Elseif boolIconSelect = True Then 
				intShowMessage = wshShell.Popup(strMessage, Time, strWindowTitle, intIcon)
				fnAddLog("FUNCTION INFO: Display strMessage.")
			Elseif boolButtonSelect = True Then
				intShowMessage = wshShell.Popup(strMessage, Time, strWindowTitle, intButton)
				fnAddLog("FUNCTION INFO: Display strMessage.")
			Else 
				intShowMessage = wshShell.Popup(strMessage, Time, strWindowTitle)
				fnAddLog("FUNCTION INFO: Display strMessage.")
			End if
			
			' Result Section
			If intShowMessage = 1 Then
				fnAddLog("FUNCTION INFO: The user has chosen the OK button." )
			Elseif intShowMessage = 2 Then
				fnAddLog("FUNCTION INFO: The user has chosen the Cancel button." )
			Elseif intShowMessage = 3 Then
				fnAddLog("FUNCTION INFO: The user has chosen the Abort button." )
			Elseif intShowMessage = 4 Then
				fnAddLog("FUNCTION INFO: The user has chosen the Retry button." )
			Elseif intShowMessage = 5 Then
				fnAddLog("FUNCTION INFO: The user has chosen the Ignore button." )
			Elseif intShowMessage = 6 Then
				fnAddLog("FUNCTION INFO: The user has chosen the Yes button." )
			Elseif intShowMessage = 7 Then
				fnAddLog("FUNCTION INFO: The user has chosen the No button." )
			Elseif intShowMessage = -1 Then
				fnAddLog("FUNCTION INFO: The time for a user response has been expired." )
			End If

			fnAddLog("FUNCTION INFO: ending fnPopupWindows.") 
			
			strLogTAB = strLogTABBak
			fnPopupWindows = intShowMessage
	Else 	
			fnAddLog("FUNCTION INFO: ending fnPopupWindows - silent installation.")
			fnPopupWindows = 0
	End If
End Function
'-------------------------------------------------------------------------------------------
'! N-strike window.
'!
'! @param  intHowManyStrikes   Receives an integer containing the number of strikes. Default value: 3 - it means that client can choose to postpone installation a maximum of 2 times. 
'! @param  strAccountName      Receives a string containing the account Name which will be displayed in message: "AcountName IT is attempting to install...."
'! @param  strWindowTitle            Receives a string containing the title that will appear as the title of the message box. Default value - empty string
'! @param  TimeTmp          Receives an integer containing the number of seconds that the box will display for until dismissed. If zero or omitted, the message box stays until the user dismisses it. Default value: 0
'! @param  strWhatInstalls     Receives a string containing the Name of application which will be displayed in the message. It could be Friendly names. Default value: PKGname
'! @param  boolIsForceKilling   Receives a value: True or False. Default value: False. True - should be display message about killing process. False - should not display message about kill process.
'! @param  strProcessesToCloseVar   Receives a string containing the Name of application which will be killed during installation. It should be friendly message. example: "Internet Browser and Java"
'! @param  boolIsRestartMessage      Receives a value: True or False. Default value: False. True - should be display message that reboot will be required. False - should not display message about reboot. Reboot is not required. 
'!
'! @return ReturnValue: 0 (The installation could be started), -1000 (Installation postponed)   
Function fnProcessesWindow(intHowManyStrikes, strAccountName, strWindowTitle, TimeTmp,  strWhatInstalls, boolIsForceKilling, strProcessesToCloseVar, boolIsRestartMessage)
	if boolSilentInstall = false Then
		Dim intRVal, intStrikeRC
        Dim strCurrentFNName                   : strCurrentFNName = "fnProcessesWindow"
		Dim strLogTABBak						: strLogTABBak = strLogTAB
			strLogTAB = strLogTAB & strLogSeparator
            
            fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "intHowManyStrikes - " & intHowManyStrikes & ", strAccountName - " & strAccountName & ", strWindowTitle - " & strWindowTitle & ", Time - " & Time & ", strWhatInstalls - " & strWhatInstalls & ", boolIsForceKilling - " & boolIsForceKilling & ", strProcessesToCloseVar - " & strProcessesToCloseVar & ", boolIsRestartMessage - " & boolIsRestartMessage & DotFN(0))
			
			intRVal = 1603
			intStrikeRC = -1000
			
			'intHowManyStrikes
			If FnParamChecker(intHowManyStrikes, "intHowManyStrikes") = 1 Then
				fnAddLog("FUNCTION INFO: variable intHowManyStrikes is NULL. Setting to 3 value.")
				intHowManyStrikes = 3
			ElseIf IsNumeric(intHowManyStrikes) = False Then
				fnAddLog("FUNCTION INFO: variable intHowManyStrikes is not numeric. Setting to 3 value.")
				intHowManyStrikes = 3
			End If
			'strAccountName
			If FnParamChecker(strAccountName, "strAccountName") = 1 Then
				fnAddLog("FUNCTION INFO: variable strAccountName is NULL. Setting to empty string.")
				strAccountName = ""
			End If
			'strWindowTitle
			If FnParamChecker(strWindowTitle, "strWindowTitle") = 1 Then
				fnAddLog("FUNCTION INFO: variable strWindowTitle is NULL. Setting to empty string.")
				strWindowTitle = " "
			End If
			'Time
			If FnParamChecker(TimeTmp, "TimeTmp") = 1 Then
				fnAddLog("FUNCTION INFO: variable TimeTmp is NULL. Setting to 0 value.")
				TimeTmp = 0
			End If
			'strWhatInstalls
                
                if strWhatInstalls = "strFriendlyPKGName" Then
				    fnAddLog("FUNCTION INFO: variable strWhatInstalls is string strFriendlyPKGName. Setting to strFriendlyPKGName.")
				    strWhatInstalls = strFriendlyPKGName
                Elseif strWhatInstalls = "strPKGName" Then
				    fnAddLog("FUNCTION INFO: variable strWhatInstalls is string strPKGName. Setting to PKGname.")
				    strWhatInstalls = strPKGName
                End If
                
			If FnParamChecker(strWhatInstalls, "strWhatInstalls") = 1 Then
				fnAddLog("FUNCTION INFO: variable strWhatInstalls is NULL. Setting to PKGname.")
				strWhatInstalls = strPKGName
			End If
			'boolIsForceKilling
			If FnParamChecker(boolIsForceKilling, "boolIsForceKilling") = 1 Then
				fnAddLog("FUNCTION INFO: variable boolIsForceKilling is NULL. Setting to False.")
				boolIsForceKilling = False 
			Elseif  Not ((boolIsForceKilling = True) or (boolIsForceKilling=False)) Then
				fnAddLog("FUNCTION INFO: variable boolIsForceKilling is wrong. Setting to False.")
				boolIsForceKilling = False 
			End If
			'strProcessesToCloseVar
			If FnParamChecker(strProcessesToCloseVar, "strProcessesToCloseVar") = 1 Then
				fnAddLog("FUNCTION INFO: variable strProcessesToCloseVar is NULL. Setting to empty string.")
				strProcessesToCloseVar = ""	'M.R. changed from Processes = "APPLIACTIONS TO CLOSE"
			End If
			'boolIsRestartMessage
			If FnParamChecker(boolIsRestartMessage, "boolIsRestartMessage") = 1 Then
				fnAddLog("FUNCTION INFO: variable boolIsRestartMessage is NULL. Setting to False.")
				boolIsRestartMessage = False
			Elseif  Not ((boolIsRestartMessage = True) or (boolIsRestartMessage=False)) Then
				fnAddLog("FUNCTION INFO: variable boolIsRestartMessage is wrong. Setting to False.")
				boolIsRestartMessage = False 
			End If
			
			'Messages section
			Dim MessageInstallOptional1, MessageInstallOptional2, MessageInstallOptional3, MessageInstallOptional4, MessageInstallOptional5, MessageInstallOptional6, MessageInstallOptional7, MessageInstallOptional8 
			Dim MessageInstallRequired1, MessageInstallRequired2, MessageInstallRequired3, MessageInstallRequired4, MessageInstallRequired5, MessageInstallRequired6, MessageInstallRequired7
			Dim MessageInstallIfRebootRequired
            Dim DisplayMessage, Cancel_Info, InstallMessage, Number_Postponed, reminingTime
                    
            reminingTime = TimeTmp/60

			Number_Postponed = Array( "once", "twice", "three times", "four times", "five times", "six times", "seven times", "eight times", "nine times", "ten times")
            
                    
                    
                    If G_MessageInstallOptional1 <> "N/A" AND FnParamChecker(G_MessageInstallOptional1, "G_MessageInstallOptional1") = 0 Then
                            MessageInstallOptional1 = G_MessageInstallOptional1
                        If InStr(MessageInstallOptional1,"<strAccountName>") >0 Then
                            MessageInstallOptional1 = Replace(MessageInstallOptional1,"<strAccountName>",strAccountName)
                        End If
                        If InStr(MessageInstallOptional1,"<strFriendlyPKGName>") >0 Then
                            MessageInstallOptional1 = Replace(MessageInstallOptional1,"<strFriendlyPKGName>",strFriendlyPKGName)
                        End If
                        If InStr(MessageInstallOptional1,"<instTime>") >0 Then
                            MessageInstallOptional1 = Replace(MessageInstallOptional1,"<instTime>",intInstallationTime)
                        End If
                    Else
                        MessageInstallOptional1	= strAccountName & " IT is attempting to install the " & strWhatInstalls & " application."
                    End If
                    
                    If G_MessageInstallOptional2 <> "N/A" AND FnParamChecker(G_MessageInstallOptional2, "G_MessageInstallOptional2") = 0 Then
                            MessageInstallOptional2 = G_MessageInstallOptional2
                        If InStr(MessageInstallOptional2,"<strFriendlyProcessName>") >0 Then
                            MessageInstallOptional2 = Replace(MessageInstallOptional2,"<strFriendlyProcessName>",strFriendlyProcessName)
                        End If
                    Else
                        MessageInstallOptional2 = "Please close all " & strProcessesToCloseVar & " windows, saving any work. "
                    End If
                    
                    If G_MessageInstallOptional3 <> "N/A" AND FnParamChecker(G_MessageInstallOptional3, "G_MessageInstallOptional3") = 0 Then
                            MessageInstallOptional3 = G_MessageInstallOptional3
                        If InStr(MessageInstallOptional3,"<strFriendlyProcessName>") >0 Then
                            MessageInstallOptional3 = Replace(MessageInstallOptional3,"<strFriendlyProcessName>",strFriendlyProcessName)
                        End If
                    Else
                        MessageInstallOptional3	= "When you click OK, any remaining " & strProcessesToCloseVar & " sessions will be closed automatically."
                    End If
                    
                    If G_MessageInstallOptional4 <> "N/A" AND FnParamChecker(G_MessageInstallOptional4, "G_MessageInstallOptional4") = 0 Then
                            MessageInstallOptional4 = G_MessageInstallOptional4
                        If InStr(MessageInstallOptional4,"<reminingTime>") >0 Then
                            MessageInstallOptional4 = Replace(MessageInstallOptional4,"<reminingTime>",reminingTime)
                        End If
                    Else
                        MessageInstallOptional4	= "You can choose to postpone this installation a maximum of " & intHowManyStrikes-1 & " times, after which it will be become mandatory. You currently have "
                    End If
                    
                    If G_MessageInstallOptional5 <> "N/A" AND FnParamChecker(G_MessageInstallOptional5, "G_MessageInstallOptional5") = 0 Then
                        MessageInstallOptional5 = G_MessageInstallOptional5
                    Else
                        MessageInstallOptional5	= " remaining opportunities to postpone."
                    End If
                    
                    If G_MessageInstallOptional6 <> "N/A" AND FnParamChecker(G_MessageInstallOptional6, "G_MessageInstallOptional6") = 0 Then
                        MessageInstallOptional6 = G_MessageInstallOptional6
                    Else
                        MessageInstallOptional6	= "Please click OK to proceed with the installation, or CANCEL to postpone the installation at this time."
                    End If
                    
                    If G_MessageInstallOptional7 <> "N/A" AND FnParamChecker(G_MessageInstallOptional7, "G_MessageInstallOptional7") = 0 Then
                        MessageInstallOptional7 = G_MessageInstallOptional7
                    End If
                    
                    If G_MessageInstallOptional8 <> "N/A" AND FnParamChecker(G_MessageInstallOptional8, "G_MessageInstallOptional8") = 0 Then
                        MessageInstallOptional8 = G_MessageInstallOptional8
                    End If
                    
			
			
			
			
			
                    If G_MessageInstallRequired1 <> "N/A" AND FnParamChecker(G_MessageInstallRequired1, "G_MessageInstallRequired1") = 0 Then
                            MessageInstallRequired1 = G_MessageInstallRequired1
                        If InStr(MessageInstallRequired1,"<strAccountName>") >0 Then
                            MessageInstallRequired1 = Replace(MessageInstallRequired1,"<strAccountName>",strAccountName)
                        End If
                        If InStr(MessageInstallRequired1,"<strFriendlyPKGName>") >0 Then
                            MessageInstallRequired1 = Replace(MessageInstallRequired1,"<strFriendlyPKGName>",strFriendlyPKGName)
                        End If
                        If InStr(MessageInstallRequired1,"<instTime>") >0 Then
                            MessageInstallRequired1 = Replace(MessageInstallRequired1,"<instTime>",intInstallationTime)
                        End If
                    Else
                        MessageInstallRequired1	= strAccountName & " IT is attempting to install the " & strWhatInstalls & " application."
                    End If
                    
                    If G_MessageInstallRequired2 <> "N/A" AND FnParamChecker(G_MessageInstallRequired2, "G_MessageInstallRequired2") = 0 Then
                            MessageInstallRequired2 = G_MessageInstallRequired2
                        If InStr(MessageInstallRequired2,"<strFriendlyProcessName>") >0 Then
                            MessageInstallRequired2 = Replace(MessageInstallRequired2,"<strFriendlyProcessName>",strFriendlyProcessName)
                        End If
                    Else
                        MessageInstallRequired2 = "Please close all " & strProcessesToCloseVar & " windows, saving any work. "
                    End If
                    
                    If G_MessageInstallRequired3 <> "N/A" AND FnParamChecker(G_MessageInstallRequired3, "G_MessageInstallRequired3") = 0 Then
                            MessageInstallRequired3 = G_MessageInstallRequired3
                        If InStr(MessageInstallRequired3,"<strFriendlyProcessName>") >0 Then
                            MessageInstallRequired3 = Replace(MessageInstallRequired3,"<strFriendlyProcessName>",strFriendlyProcessName)
                        End If
                    Else
                        MessageInstallRequired3 = "When you click OK, any remaining " & strProcessesToCloseVar & " sessions will be closed automatically."
                    End If
                    
                    If G_MessageInstallRequired4 <> "N/A" AND FnParamChecker(G_MessageInstallRequired4, "G_MessageInstallRequired4") = 0 Then
                            MessageInstallRequired4 = G_MessageInstallRequired4
                        If InStr(MessageInstallRequired4,"<Number_Postponed((intHowManyStrikes-2))>") >0 Then
                            MessageInstallRequired4 = Replace(MessageInstallRequired4,"<Number_Postponed((intHowManyStrikes-2))>",(intHowManyStrikes-1))
                        End If
                    Else
                        MessageInstallRequired4 = "The installation is now mandatory, as it has already been postponed " & Number_Postponed((intHowManyStrikes-2)) & "."
                    End If

                    If G_MessageInstallRequired5 <> "N/A" AND FnParamChecker(G_MessageInstallRequired5, "G_MessageInstallRequired5") = 0 Then
                            MessageInstallRequired5 = G_MessageInstallRequired5
                    Else
			             MessageInstallRequired5 = "Please click OK to proceed with the installation at this time."
                    End If

                    If G_MessageInstallRequired6 <> "N/A" AND FnParamChecker(G_MessageInstallRequired6, "G_MessageInstallRequired6") = 0 Then
                            MessageInstallRequired6 = G_MessageInstallRequired6
                    End If

                    If G_MessageInstallRequired7 <> "N/A" AND FnParamChecker(G_MessageInstallRequired7, "G_MessageInstallRequired7") = 0 Then
                            MessageInstallRequired7 = G_MessageInstallRequired7
                    End If
			
			
          
            If G_MessageInstallIfRebootRequired <> "N/A" AND FnParamChecker(G_MessageInstallIfRebootRequired, "G_MessageInstallIfRebootRequired") = 0 Then
                MessageInstallIfRebootRequired = G_MessageInstallIfRebootRequired
            Else
                MessageInstallIfRebootRequired = "It will be necessary to restart your machine following this installation."
            End If          
        

			' Strike Section - read strike count 
			Dim value, ifexist
			Set objReg 		= GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
			Const HKEY_LOCAL_MACHINE= &H80000002
			
			'Set proper registry keys according to OS architecture
			If fnIs64BitOS(True) = 0 Then
				strKeyPath		= "SOFTWARE\Wow6432Node\IBM\Packages\"
				strKeyPath2		= "HKLM\SOFTWARE\Wow6432Node\IBM\Packages\" & PKGname & "\StrikeCount"
			Else
				strKeyPath		= "SOFTWARE\IBM\Packages\"
				strKeyPath2		= "HKLM\SOFTWARE\IBM\Packages\" & PKGname & "\StrikeCount"
			End If
						
			'New Three Strike registry key creation
			objReg.CreateKey HKEY_LOCAL_MACHINE, strKeyPath & PKGname
			On Error Resume Next
				ifexist = objShell.RegRead(strKeyPath2)
				If ifexist = "" Then
					value = 0
					objShell.RegWrite "HKEY_LOCAL_MACHINE\" & strKeyPath & PKGname & "\StrikeCount" , value
					fnAddLog("FUNCTION INFO: The Three Strike registry: HKEY_LOCAL_MACHINE\" & strKeyPath & PKGname & "\StrikeCount has been added with value: " & value & ".")
				Else
					value= objShell.RegRead( "HKEY_LOCAL_MACHINE\" & strKeyPath & PKGname & "\StrikeCount")
				End If
			On Error GoTo 0
			
			' Result Section
			If ( (value = ("" & (intHowManyStrikes-1)))) Then
				'Prepare message
				DisplayMessage = MessageInstallRequired1
				If boolIsForceKilling = True AND NOT strProcessesToCloseVar = "" Then
					DisplayMessage = DisplayMessage & Chr(32) & MessageInstallRequired2 & vbCrLf & vbCrLf & MessageInstallRequired3
				End If
				DisplayMessage = DisplayMessage & vbCrLf & vbCrLf & MessageInstallRequired4
				If boolIsRestartMessage = True Then DisplayMessage = DisplayMessage & vbCrLf & vbCrLf & MessageInstallIfRebootRequired End If
				DisplayMessage = DisplayMessage & vbCrLf & vbCrLf & MessageInstallRequired5
				'Show message
				InstallMessage = wshShell.Popup( DisplayMessage, TimeTmp, strWindowTitle, 0 + 48)
				'Close message
				If InstallMessage = -1 then
						fnAddLog("FUNCTION INFO: The time for a user response has been expired. The installation could be started.")
				Elseif InstallMessage = 1 then
						fnAddLog("FUNCTION INFO: The user has chosen the OK button. The installation could be started.")
				Else 
						fnAddLog("FUNCTION INFO: The user closed the msgbox button. The installation could be started.")
				End if
				fnAddLog("FUNCTION INFO: Installation can be started.")
				intRVal = 0 
			Else
				Cancel_Info =((intHowManyStrikes-1)-value)
				'Prepare message
				DisplayMessage = MessageInstallOptional1 
				If boolIsForceKilling = True AND NOT strProcessesToCloseVar = "" Then
					DisplayMessage = DisplayMessage & Chr(32) & MessageInstallOptional2 & vbCrLf & vbCrLf & MessageInstallOptional3
				End If
				DisplayMessage = DisplayMessage & vbCrLf & vbCrLf  & MessageInstallOptional4 & " " & Cancel_Info & " " & MessageInstallOptional5 & vbCrLf & vbCrLf
				If boolIsRestartMessage = True Then DisplayMessage = DisplayMessage & MessageInstallIfRebootRequired & vbCrLf & vbCrLf End If 
				DisplayMessage = DisplayMessage & MessageInstallOptional6 
				'Show message
				InstallMessage = wshShell.Popup( DisplayMessage, TimeTmp, strWindowTitle, 1 + 48)	
				'Close message
				If InstallMessage = 1 Then
						fnAddLog("FUNCTION INFO: The user has chosen the OK button. The installation could be started.")
						fnAddLog("FUNCTION INFO: Installation can be started.")
						intRVal = 0 
				Elseif InstallMessage = -1 then
						value = value + 1
						objShell.RegWrite "HKEY_LOCAL_MACHINE\" & strKeyPath & PKGname & "\StrikeCount" , value 
						fnAddLog("FUNCTION INFO: The time for a user response has been expired. The new value count is now: " & value & ".")
						fnAddLog("FUNCTION INFO: Installation cannot be started.")
						intRVal = intStrikeRC
				Else
						value = value + 1
						objShell.RegWrite "HKEY_LOCAL_MACHINE\" & strKeyPath & PKGname & "\StrikeCount" , value 
						fnAddLog("FUNCTION INFO: The user closed the msgbox button or choose the Cancel option. The new value count is now: " & value & ".")
						fnAddLog("FUNCTION INFO: Installation cannot be started.")
						intRVal = intStrikeRC
				End If 
			End If
			fnAddLog("FUNCTION INFO: ending fnProcessesWindow.") 
			fnProcessesWindow = intRVal
					fnAddLog("SCRIPT INFO: going to fnProgressBarSet.")
					
					strLogTAB = strLogTABBak
					If Not intRVal = -1000 then call fnProgressBarSet(boolThisInstall) 
	Else 	
			fnAddLog("FUNCTION INFO: ending fnProcessesWindow - silent installation.")
			fnProcessesWindow = 0
	End If
End Function
'-------------------------------------------------------------------------------------------
'! Clean registry from n-strike window counter.
'!
'! @param  PKGNameV   		Package Name according naming convention.
'!
'! @return 0,1
Function ProcessesWindowRegClean(PKGNameV)
    Dim strCurrentFNName                   : strCurrentFNName = "ProcessesWindowRegClean"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "PKGNameV - " & PKGNameV & DotFN(0))
		
				
		If FnParamChecker(PKGNameV, "PKGNameV") = 1 Then
			fnAddLog("FUNCTION ERROR: variable PKGNameV cannot be NULL.")
			fnAddLog("FUNCTION INFO: ending ProcessesWindowRegClean - fail.") 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			ProcessesWindowRegClean = 1
		End If
		
		
	Dim value, ifexist
		Set objReg 		= GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
		
		'Set proper registry keys according to OS architecture
			If fnIs64BitOS(True) = 0 Then
				strKeyPath		= "HKLM\SOFTWARE\Wow6432Node\IBM\"
			Else
				strKeyPath		= "HKLM\SOFTWARE\IBM\"
			End If
		
		On Error Resume Next
			ifexist = objShell.RegRead(strKeyPath & "Packages\" & PKGnameV & "\StrikeCount")
			If ifexist = "" Then
				fnAddLog("FUNCTION INFO: The Three Strike registry: " & strKeyPath & "Packages\" & PKGnameV & "\StrikeCount not present. Exiting." & ".")
				
				strLogTAB = strLogTABBak
				ProcessesWindowRegClean = 0
				Exit Function
			End If
		On Error GoTo 0
		
	if fnCheckIfKeyExist(strKeyPath & "Packages\" & PKGnameV & "\StrikeCount") = 0 then 
		On Error Resume Next
		wshShell.RegDelete(strKeyPath & "Packages\" & PKGnameV & "\StrikeCount")
		if Err = 0 then
			fnAddLog "Key " & strKeyPath & "Packages\" & PKGnameV & "\StrikeCount has been deleted"
			wshShell.RegDelete(strKeyPath & "Packages\" & PKGnameV & "\")
			if Err = 0 then
				fnAddLog "Key " & strKeyPath & "Packages\" & PKGnameV & " has been deleted"
				wshShell.RegDelete(strKeyPath & "Packages\")
				if Err = 0 then
					fnAddLog "Key " & strKeyPath & "Packages has been deleted"
					wshShell.RegDelete(strKeyPath)
					if Err = 0 then
						fnAddLog "Key " & strKeyPath & " has been deleted"
					else
						fnAddLog "Unable to delete: " & strKeyPath & " - not empty"
					end if
				else
					fnAddLog "Unable to delete: " & strKeyPath & "Packages" & " - not empty"
				end if
			else
				fnAddLog "Unable to delete: " & strKeyPath & "Packages\" & PKGnameV & " - not empty"
			end if
		end if
	end if
	
	
	strLogTAB = strLogTABBak
	ProcessesWindowRegClean = 0
	Exit Function
End Function