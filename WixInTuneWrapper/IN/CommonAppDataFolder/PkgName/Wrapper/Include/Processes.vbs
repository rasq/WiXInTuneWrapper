'****************************************************************************************************************************************************
' FUNCTIONS: 
'	fnIsProcessRunning(strProcessName)
'	fnKillProcess(strProcessName)
'	fnKillProcessCriterium (strProcessName)  
'	fnWaitUntilProcessIsRunning(strProcessName)
'   fnKillProcessCMDLine(strProcessName, strProcessCMDLine) 
'   IsProcessRunningComandLine(strProcessName, strProcessCMDLine)
'	
' FUNCTION EXAMPLE USAGE (IN and OUT):
'	fnIsProcessRunning("notepad.exe,mspaint.exe")
'	Return: 0,1
'
'	fnKillProcess("svchost.exe,calc.exe")
'	Return: 0,1
'
'	fnKillProcessCriterium ("word,snipp") 
'	Return: 0,1
'
'	fnWaitUntilProcessIsRunning("")
'	Return: 0
'MsgBox "fnIsProcessRunning() = " & fnIsProcessRunning("notepad.exe,mspaint.exe")
'MsgBox "fnKillProcess() = " & fnKillProcess("mspaint.exe")
'MsgBox "fnKillProcess() = " & fnKillProcess("svchost.exe,calc.exe")
'MsgBox "fnKillProcessCriterium() = " & fnKillProcessCriterium ("pad,snipp")
'MsgBox "fnWaitUntilProcessIsRunning() = " & fnWaitUntilProcessIsRunning("StikyNot.exe,mspaint.exe")
'***************************************************************************************************************************************************
'! Return information about processes status (are (or not) running).
'!
'! @param  strProcessName        Process Name - as is in ProcessExplorer. It can be list of processes with separator "'".
'!
'! @return 0,1
Function fnIsProcessRunning(strProcessName)
	Dim colProcess, intRVal, arrProcesses, strProcess
    Dim strCurrentFNName                   : strCurrentFNName = "fnIsProcessRunning"
	Dim strCount 						: strCount = 0
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strProcessName - " & strProcessName & DotFN(0)) 
		
		If FnParamChecker(strProcessName, "strProcessName") = 1 Then
            call fnConst_VarFail("strProcessName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnIsProcessRunning = 1
		End If
		
		arrProcesses = Split (strProcessName, ",")

		For Each strProcess In arrProcesses
			Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & strProcess & "'")
		 
			If colProcess.Count > 0 Then
				fnAddLog("FUNCTION INFO: Process " & strProcess & " is running. There is/are " & colProcess.Count & " instance/s of the process.")
				strCount = strCount + 1
			Else
				fnAddLog("FUNCTION INFO: Process " & strProcess & " is not running.")
			End If
		Next
		
		If strCount > 0 Then
			intRVal = 0
		Else
			intRVal = 1
		End If
		
        call Const_FnEndrVal(strCurrentFNName, intRVal)
		
		strLogTAB = strLogTABBak
		fnIsProcessRunning = intRVal
End Function
'-------------------------------------------------------------------------------------------
'! Will kill running processes.
'!
'! @param  strProcessName			Name of the process to be closed (case insensitive). It can be list of processes separated by commas.
'!
'! @return 0,1
Function fnKillProcess(strProcessName)
	Dim arrProcesses, strProcess, strProcessFail, objProcess, colProcess, strKillCMD, boolBlockade		'Unknown - added Blockade variable in case there are multiple instances of the same process
	Dim strCurrentFNName					: strCurrentFNName = "fnKillProcess"
	Dim strFail							: strFail = 0
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strProcessName" & DotFN(3) & strProcessName & DotFN(0))
		
		If FnParamChecker(strProcessName, "strProcessName") = 1 Then
			Call fnConst_VarFail("strProcessName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnKillProcess = 1
		End If
		
		
		arrProcesses = Split(strProcessName, Chr(44))
		
		For Each strProcess In arrProcesses
			Set colProcess = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & strProcess & "'")
			For Each objProcess In colProcess
				fnAddLog(PreTxtFN(0) & "Trying to kill process" & DotFN(3) & objProcess.Name & DotFN(0))
				
				If fnIsProcessRunning(objProcess.Name) = 1 Then
					fnAddLog(PreTxtFN(0) & "Process " & objProcess.Name & " is not running, no need to terminate.")
					boolBlockade = True
				Else
					fnAddLog(PreTxtFN(0) & "Process " & objProcess.Name & " is running, it will be terminated.")
					boolBlockade = False
				End If
				
				If boolBlockade = False Then
					Err.Clear
					strKillCMD = "taskkill /t /f /im " & Chr(34) & objProcess.Name & Chr(34)
					objShell.Run strKillCMD, 0, True
					
					WScript.Sleep(2000)
					
					If Err.Number = 0 Then
						fnAddLog(PreTxtFN(0) & "Process " & objProcess.Name & " has been terminated.")
					Else
						fnAddLog(PreTxtFN(1) & "Process " & objProcess.Name & " has not been terminated. Error code: " & Err.Number & DotFN(0))
					End If
				End If
			Next
		Next
		
		WScript.Sleep(10000)
		
		fnAddLog(SelfCheckFN)
		For Each strProcess In arrProcesses
			If fnIsProcessRunning(strProcess) = 0 Then
				strProcessFail = strProcessFail & strProcess & Chr(44) & Chr(32)
				strFail = strFail + 1
			End If
		Next
		
		If strFail = 0 Then
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0))
			
			strLogTAB = strLogTABBak
			fnKillProcess = 0
		Else
			fnAddLog(EndingFN(2) & strCurrentFNName & SucFailFN(1) & " Some processes are still running: " & Left(strProcessFail, Len(strProcessFail) - 2) & DotFN(0))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnKillProcess = 1
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Will kill specified process with a specified command line.
'!
'! @param  strProcessName        Process Name - as is in ProcessExplorer. It can be list of processes with separator "'".
'! @param  strProcessCMDLine     Process CMD - as is in ProcessExplorer. This can separate valid process when two or more have this same Name.
'!
'! @return 0,1
Function fnKillProcessCMDLine(strProcessName, strProcessCMDLine)  ' P.W.
	Dim arrProcesses, strProcess, objProcess, colProcess, intRVal, strTermCode, i, strKillCMD, strTmpVar, strPidVar
    Dim strCurrentFNName                   : strCurrentFNName = "fnKillProcessCMDLine"
	Dim strFail 						: strFail = 0
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strProcessName - " & strProcessName & ", strProcessCMDLine - " & strProcessCMDLine & DotFN(0)) 
		
		If FnParamChecker(strProcessName, "strProcessName") = 1 Then
            call fnConst_VarFail("strProcessName", strCurrentFNName) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnKillProcessCMDLine = 1
		End If
		
		If FnParamChecker(strProcessCMDLine, "strProcessCMDLine") = 1 Then
            call fnConst_VarFail("strProcessCMDLine", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnKillProcessCMDLine = 1
		End If
    
            strTmpVar = InStr(strProcessCMDLine, "\\")

            if strTmpVar = 0 Then
                strProcessCMDLine = Replace(strProcessCMDLine,"\","\\")
            End If
    
		'on error resume next
			
		Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & strProcessName & "'" & " and CommandLine Like '%%" & strProcessCMDLine & "%%'")
		
		
			If IsProcessRunningComandLine(strProcessName, strProcessCMDLine) = 1 Then
				fnAddLog("FUNCTION INFO: ending fnKillProcessCMDLine, objProcess - " & strProcessName & " and Command Line include: " & strProcessCMDLine & " - doesen't exist.") 
				fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0))
					
				strLogTAB = strLogTABBak
				fnKillProcessCMDLine = 0
				Exit Function
			End If
			
			
		For Each objProcess in colProcess
			fnAddLog("FUNCTION INFO: Process: " & objProcess.Name & " and Command Line include: " & objProcess.CommandLine & " ,with ID = " & objProcess.Processid & " will be terminated.")
			
			strPidVar = strPidVar & " /PID " & objProcess.Processid
		Next		
				
			
			strKillCMD = "taskkill /t /f" & strPidVar
				fnAddLog("FUNCTION INFO: Process taskkill cmd: " & strKillCMD & ".")
			
			objShell.Run strKillCMD,0, True
		
			WScript.Sleep(2000)
				
			If IsProcessRunningComandLine(strProcessName, strProcessCMDLine) = 1 Then
				fnAddLog("FUNCTION INFO: Process " & strProcessName & " and Command Line include: " & strProcessCMDLine & " has been terminated.")
			Else
				fnAddLog("FUNCTION INFO: Process " & strProcessName & " and Command Line include: " & strProcessCMDLine &" has not been terminated.")
				strFail = strFail + 1
			End If
			
		
					If strFail = 0 Then
						intRVal = 0
					Else
						intRVal = 1
					End If
			
			
        call Const_FnEndrVal(strCurrentFNName, intRVal)
		
		strLogTAB = strLogTABBak
		fnKillProcessCMDLine = intRVal
End Function
'--------------------------------------------------------------------------------------------------------------------------
'! It can wait until specified process with specified command line will be running.
'!
'! @param  strProcessName        Process Name - as is in ProcessExplorer. It can be list of processes with separator "'".
'! @param  strProcessCMDLine     Process CMD - as is in ProcessExplorer. This can separate valid process when two or more have this same Name.
'!
'! @return 0,1
Function WaitUntilProcessIsRunningCMDLine(strProcessName, strProcessCMDLine)
	Dim colProcess, intRVal, strTmpVar
    Dim strCurrentFNName                   : strCurrentFNName = "WaitUntilProcessIsRunningCMDLine"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strProcessName - " & strProcessName & ", strProcessCMDLine - " & strProcessCMDLine & DotFN(0)) 
		
			If FnParamChecker(strProcessName, "strProcessName") = 1 Then
            call fnConst_VarFail("strProcessName", strCurrentFNName) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			WaitUntilProcessIsRunningCMDLine = 1
		End If
		
		If FnParamChecker(strProcessCMDLine, "strProcessCMDLine") = 1 Then
            call fnConst_VarFail("strProcessCMDLine", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			WaitUntilProcessIsRunningCMDLine = 1
		End If
			
			
			strTmpVar = InStr(strProcessCMDLine, "\\")

            if strTmpVar = 0 Then
                strProcessCMDLine = Replace(strProcessCMDLine,"\","\\")
            End If
			
			Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & strProcessName & "'" & " and CommandLine Like '%%" & strProcessCMDLine & "%%'")
			
				If colProcess.Count > 0 Then
					fnAddLog("FUNCTION INFO: Process " & Chr(34) & strProcessName & Chr(34) & " is running. Pauses script execution until the process exists.")
				End If
				
				Do While colProcess.Count > 0
					WScript.Sleep 5000
					Set colProcess = Nothing
					Set colProcess = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & strProcessName & "'")
				Loop
				
				fnAddLog("FUNCTION INFO: Process " & Chr(34) & strProcessName & Chr(34) & " is not running.")
				
		
		intRVal = 0
		
		fnAddLog("FUNCTION INFO: Ending function WaitUntilProcessIsRunningCMDLine with result " & intRVal & ".")
		
		strLogTAB = strLogTABBak
		WaitUntilProcessIsRunningCMDLine = intRVal
  End Function
'--------------------------------------------------------------------------------------------------------------------------
'! Return information about processes status (are (or not) running).
'!
'! @param  strProcessName        Process Name - as is in ProcessExplorer. It can be list of processes with separator "'".
'! @param  strProcessCMDLine     Process CMD - as is in ProcessExplorer. This can separate valid process when two or more have this same Name.
'!
'! @return 0,1
Function IsProcessRunningComandLine(strProcessName, ProcessCMDLineV)   ' P.W.
	Dim colProcess, intRVal, arrProcesses, strProcess, i, strTmpVar, strProcessCMDLine
    Dim strCurrentFNName                   : strCurrentFNName = "IsProcessRunningComandLine"
	Dim strCount 						: strCount = 0
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		strProcessCMDLine = ProcessCMDLineV
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strProcessName - " & strProcessName & ", strProcessCMDLine - " & strProcessCMDLine & DotFN(0)) 
		
		If FnParamChecker(strProcessName, "strProcessName") = 1 Then
            call fnConst_VarFail("strProcessName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			IsProcessRunningComandLine = 1
		End If
		
		If FnParamChecker(strProcessCMDLine, "strProcessCMDLine") = 1 Then
            call fnConst_VarFail("strProcessCMDLine", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			IsProcessRunningComandLine = 1
		End If
    
            strTmpVar = InStr(strProcessCMDLine, "\\")

            if strTmpVar = 0 Then
                strProcessCMDLine = Replace(strProcessCMDLine,"\","\\")
            End If
    
		Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name Like '%%" & strProcessName & "%%'" & " and CommandLine Like '%%" & strProcessCMDLine & "%%'")' 

		If colProcess.Count > 0 Then
			fnAddLog("FUNCTION INFO: Process " & strProcessName & " and Command Line include: " & strProcessCMDLine & " is running. There is/are " & colProcess.Count & " instance/s of the process.")
			strCount = strCount + 1
		Else
			fnAddLog("FUNCTION INFO: Process " & strProcessName & " and Command Line include: " & strProcessCMDLine & " is not running.")
		End If
	
		
		If strCount > 0 Then
			intRVal = 0
		Else
			intRVal = 1
		End If
		
        call Const_FnEndrVal(strCurrentFNName, intRVal)
		
		strLogTAB = strLogTABBak
		IsProcessRunningComandLine = intRVal
End Function

'--------------------------------------------------------------------------------------------------------------------------
'! Can kill any process with string in Name that is specified as strProcessName.
'!
'! @param  strProcessName        Process Name - it can be only part of real Name. It can be list of processes with separator "'".
'!
'! @return 0,1
Function fnKillProcessCriterium (strProcessName)
	Dim arrProcesses, strProcess, objProcess, strTermCode, intRVal, i
    Dim strCurrentFNName                   : strCurrentFNName = "fnKillProcessCriterium"
	Dim strFail 						: strFail = 0
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strProcessName - " & strProcessName & DotFN(0)) 
		
		If FnParamChecker(strProcessName, "strProcessName") = 1 Then
            call fnConst_VarFail("strProcessName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnKillProcessCriterium = 1
		End If

		arrProcesses = Split (strProcessName, ",")
		i = 0
		For Each strProcess In arrProcesses
			Set colProcessList = objWMIService.ExecQuery ("SELECT * FROM Win32_Process WHERE Name like '%" & strProcess & "%'")
			For Each objProcess in colProcessList
				fnAddLog("FUNCTION INFO: trying to kill process: " & objProcess.Name & ".")
				
				If fnIsProcessRunning(objProcess.Name) = 1 Then
					fnAddLog("FUNCTION INFO: ending fnKillProcessCriterium, objProcess - " & objProcess.Name & " doesen't exist.") 
					fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0))
					
					strLogTAB = strLogTABBak
					fnKillProcessCriterium = 0
					Exit Function
				End If
				
				fnAddLog("FUNCTION INFO: Process " & objProcess.Name & " will be terminated.")
				strTermCode = objProcess.Terminate()
				
				If strTermCode = 0 Then
					fnAddLog("FUNCTION INFO: Process " & objProcess.Name & " terminated.")
				ElseIf strTermCode = 2 Then
					fnAddLog("FUNCTION INFO: Unable to terminate process " & objProcess.Name & ". Access Denied.")
				ElseIf strTermCode = 3 Then
					fnAddLog("FUNCTION INFO: Unable to terminate process '" & objProcess.Name & "'. Insufficient Privilege.")
				ElseIf strTermCode = 8 Then
					fnAddLog("FUNCTION INFO: Unable to terminate process '" & objProcess.Name & "'. Unknown Failure.")
				ElseIf strTermCode = 9 Then
					fnAddLog("FUNCTION INFO: Unable to terminate process '" & objProcess.Name & "'. Path Not Found.")
				ElseIf strTermCode = 21 Then
					fnAddLog("FUNCTION INFO: Unable to terminate process '" & objProcess.Name & "'. Invalid Parameter.")
				End If
				
				If strTermCode <> 0 Then
					strFail = strFail + 1
				End If
				
				i = i + 1
			Next
		Next
		
		If strFail = 0 Then
			intRVal = 0
		Else
			intRVal = 1
		End If
		
        call Const_FnEndrVal(strCurrentFNName, intRVal)
		
		strLogTAB = strLogTABBak
		fnKillProcessCriterium = intRVal
End Function
'-------------------------------------------------------------------------------------------
'! It can wait until specified process will be running.
'!
'! @param  strProcessName        Process Name - as is in ProcessExplorer. It can be list of processes with separator "'".
'!
'! @return 0,1
Function fnWaitUntilProcessIsRunning(strProcessName)
	Dim objProcess, intRVal, arrProcesses, strProcess
    Dim strCurrentFNName                   : strCurrentFNName = "fnWaitUntilProcessIsRunning"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strProcessName - " & strProcessName & DotFN(0)) 
		
		If FnParamChecker(strProcessName, "strProcessName") = 1 Then
            call fnConst_VarFail("strProcessName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnWaitUntilProcessIsRunning = 1
		End If

		arrProcesses = Split (strProcessName, ",")
		For Each strProcess In arrProcesses
			Set objProcess = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & strProcess & "'")
			
			If objProcess.Count > 0 Then
				fnAddLog("FUNCTION INFO: Process " & Chr(34) & strProcess & Chr(34) & " is running. Pauses script execution until the process exists.")
			End If
			
			Do While objProcess.Count > 0
				WScript.Sleep 5000
				Set objProcess = Nothing
				Set objProcess = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & strProcess & "'")
			Loop
			fnAddLog("FUNCTION INFO: Process " & Chr(34) & strProcess & Chr(34) & " is not running.")
		Next
		
		intRVal = 0
		
        call Const_FnEndrVal(strCurrentFNName, intRVal)
		
		strLogTAB = strLogTABBak
		fnWaitUntilProcessIsRunning = intRVal
  End Function