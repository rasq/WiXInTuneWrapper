'****************************************************************************************************************************************************
' FUNCTIONS:  
' fnRunVBS(strVBSname, strVbsDirectory, strParameters)  
'	
' FUNCTION EXAMPLE USAGE (IN and OUT):
' fnRunVBS("Microsoft-.NETFramework-4.5.50709-EN-1.0-Install.vbs", strRootDrive & "\Microsoft-NETFramework-4.5.50709-EN-1.0.vbs", NULL) 
' 	Return 1, 0
' 
'***************************************************************************************************************************************************
'! Run external VBScript and catch it's exit code. 
'!
'! @param  strVBSname          File Name (*.vbs).
'! @param  strVbsDirectory        Full path to external script.
'! @param  strParameters       Additional CMD parameters to external script.
'!
'! @return 0,1
Function fnRunVBS(strVBSname, strVbsDirectory, strParameters)
	Dim strVBSInstallCommand, intRVal
    Dim strCurrentFNName                    : strCurrentFNName = "fnRunVBS"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strVBSname - " & strVBSname & ", strVbsDirectory - " & strVbsDirectory & ", strParameters - " & strParameters & DotFN(0))
			
				If FnParamChecker(strVBSname, "strVBSname") = 1 Then
                    call fnConst_VarFail("strVBSname", strCurrentFNName)
					
					strLogTAB = strLogTABBak
					fnRunVBS = 1
				End If
				
				If FnParamChecker(strVbsDirectory, "strVbsDirectory") = 1 Then
			         fnAddLog(VarsFNErr(0) & "strVbsDirectory" & VarsFNErr(1))
			         
                        strVbsDirectory = "\"
        
			         fnAddLog(SetVarFN(0) & "strVbsDirectory" & SetVarFN(1) & strVbsDirectory & DotFN(0))
				End If
				
				If FnParamChecker(strParameters, "strParameters") = 1 Then
			         fnAddLog(VarsFNErr(0) & "strParameters" & VarsFNErr(1))
			         
                        strParameters = ""
        
			         fnAddLog(SetVarFN(0) & "strParameters" & SetVarFN(1) & "empty string" & DotFN(0))
				End If
				
					If fnIsExistD(strVbsDirectory) = 0 Then
						fnAddLog("FUNCTION INFO: strVbsDirectory- " & strVbsDirectory & " exist. Going to next step.")
						If fnIsExist(strVBSname, strVbsDirectory) = 0 Then
							fnAddLog("FUNCTION INFO: strVBSname- " & strVBSname & " exist. Going to next step.")
							
										strVBSInstallCommand = "wscript.exe " & Chr(34) & strVbsDirectory & "\" & strVBSname & Chr(34) & " " & strParameters
											fnAddLog("INFO: strVBSInstallCommand: " & strVBSInstallCommand)
										intRVal = objShell.Run(strVBSInstallCommand, 1, TRUE)
											fnAddLog("INFO: " & strVBSname & " installation ended with intRVal = " & intRVal & ".")
											
										If intRVal<>3010 And intRVal<>0 Then 
												fnAddLog("FUNCTION INFO: ending fnRunVBS - fail.") 	
												
												strLogTAB = strLogTABBak
												fnFinalize(intRVal)											
												fnRunVBS = 1
												
										Else 
												If intRVal = 3010 then boolRebootContainer = True
												fnAddLog("FUNCTION INFO: ending fnRunVBS.") 

												strLogTAB = strLogTABBak												
												fnRunVBS = 0
										End If			
						Else
							fnAddLog("FUNCTION INFO: strVBSname- " & strVBSname & " don't exist. Exiting function.")
							fnAddLog("FUNCTION INFO: ending fnRunVBS - fail.") 

							strLogTAB = strLogTABBak
							fnFinalize(1)		
							fnRunVBS = 1
						End If
					Else
						fnAddLog("FUNCTION INFO: strVbsDirectory- " & strVbsDirectory & " don't exist. Exiting function.")
						fnAddLog("FUNCTION INFO: ending fnRunVBS - fail.") 		
						
						strLogTAB = strLogTABBak
						fnFinalize(1)	
						fnRunVBS = 1
					End If
End Function