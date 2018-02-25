'****************************************************************************************************************************************************
' FUNCTIONS: 
'	fnStartLog(strLogName)  
'	fnAddLog(Text)   
'	fnStopLog(intRVal)
'	
' FUNCTION EXAMPLE USAGE (IN and OUT):
' 
'***************************************************************************************************************************************************
'! Create log file (install or uninstall) and create log directory if doesen't exist. Add intFirst line of log. 
'! If strLOGType = "DEBUG" will show msgbox instead of log file creation. 
'!
'! @param  strLogName          Log file Name.
'!
'! @return N/A.
Function fnStartLog(strLogName)
	If strLOGType = "FILE" Then
		Const intForAppending = 8
		Dim objFolder
		Dim strLogDir : strLogDir = strLogFolder & "\" & strPKGName
		
		If Not objFSO.FolderExists(strLogDir) Then
			call fnCreateDirectory(strLogDir, false)
		End If
        
		strLogDir = strLogDir & "\" & strLogName	
		Set objLog = objFSO.OpenTextFile(strLogDir, intForAppending, true, 0)
        
			If fnIsObject(objLog) Then
				objLog.WriteLine StartLogFN(0) & Date & " " & Time & " " & StartLogFN(1) & strLogName & StartLogFN(2)
			End If
	ElseIf strLOGType = "DEBUG" Then
		call msgbox(StartLogFN(0) & Date & " " & Time & " " & StartLogFN(1) & strLogName & StartLogFN(2) ,0, "Info")
	End If
End Function
'-------------------------------------------------------------------------------------------
'! Add line of text to log file (install or uninstall).
'! If strLOGType = "DEBUG" will show msgbox instead of write text line into file. 
'!
'! @param  Text             Text to write into log file.
'!
'! @return N/A.
Function fnAddLog(Text)
	If strLOGType = "FILE" Then
		If fnIsObject(objLog) Then 
			objLog.WriteLine "[" & Time & "]:" & strLogTAB & "->" & Text
		End If
	ElseIf strLOGType = "DEBUG" Then
		call msgbox("[" & Time & "]:" & strLogTAB & "->" & Text ,0, "Info")
	End If
End Function
'-------------------------------------------------------------------------------------------
'! Close log file (install or uninstall). Add intLast line of log wiht script intRVal. 
'! If strLOGType = "DEBUG" will show msgbox instead of log file creation.
'!
'! @param  intRVal             Script exit code.
'!
'! @return N/A.
Function fnStopLog(intRVal)
	if intRVal <> 0 and intRVal <> 3010 Then
		if boolThisInstall = true Then
			fnAddLog(SetupFaiFN(0)) 
		Else
			fnAddLog(SetupFaiFN(1)) 
		End If
	Else
		if boolThisInstall = true Then
			fnAddLog(SetupSucFN(0)) 
		Else
			fnAddLog(SetupSucFN(1)) 
		End If
	End If	
	
		fnAddLog(StopLogFN(0) & Date & " " & Time & " " & StopLogFN(1) & strLogName & StopLogFN(2))
		fnAddLog("")
		fnAddLog(LineFN)
	
    If fnIsObject(objLog) Then
	   objLog.Close
    End If
End Function
'-------------------------------------------------------------------------------------------