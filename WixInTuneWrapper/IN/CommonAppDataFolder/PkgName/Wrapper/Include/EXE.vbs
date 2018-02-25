'****************************************************************************************************************************************************
' FUNCTIONS: 
'	fnInstallEXE(strEXEFile, strEXEParams, strVbsDirectory, strSuccessTable, strName) 
'	fnUninstallEXE(strEXEFile, strEXEParams, strVbsDirectory, strSuccessTable, strName) 
'	fnGetARPUninstallString(strGUID, strAppBit)
'	fnARPUninstall(strARPUString, strExeParams, strName)  
'	fnSplitARPStringAndUninstall(strARPUString, strExeSeparator, strParamsSeparator, strOptionalParams, strName)
'	
' FUNCTION EXAMPLE USAGE (IN and OUT):
' fnInstallEXE("notepadpp.exe", "/QUIET /NORESTART", fnCurrentDir, rValSuccessCodeTable, NULL) 						fnCurrentDir is a variable defined in main template, rValSuccessCodeTable is a table containing return codes for successful installation - defined in main template
'	Return 0,1
'
'fnUninstallEXE("notepadpp.exe", "/REMOVE /QUIET /NORESTART", fnCurrentDir, rValSuccessCodeTable, NULL) 
'	Return 0,1
'
'fnGetARPUninstallString("{716E0306-8318-4364-8B8F-0CC4E9376BAC}", "x64")
'	Return C:\Program Files (x86)\Notepad++\uninstall.exe
'
'fnARPUninstall("C:\Program Files (x86)\Notepad++\uninstall.exe", "/REMOVE /QUIET /NORESTART", strPKGName) 
'	Return 0,1
'
'fnSplitARPStringAndUninstall("C:\App\uninstall.exe \silent \allmodules", ".exe", "\", "", strPKGName)
'	Return 0,1,fnUninstallEXE return value
'
'***************************************************************************************************************************************************
'! Install specified exe file with specified parameters. 
'!
'! @param  strEXEFile          Exe filename.
'! @param  strEXEParams        Additional cmd parameters to use with exe file durring installation.
'! @param  strVbsDirectory        strVbsDirectory (path) where exe file is placed.
'! @param  strSuccessTable     Succes exit codes for exe file (all codes separeted with ",").
'! @param  strName             Package Name (according naming convention), will be used to create or check TAGfile.
'!
'! @return intRVal (0,3010)
Function fnInstallEXE(strEXEFile, strEXEParams, strVbsDirectory, strSuccessTable, strName)	
    Dim strEXEInstallCommand, intRVal, item, arrRCList, rValAfterCMD
    Dim strCurrentFNName                   : strCurrentFNName = "fnInstallEXE"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strEXEFile - " & strEXEFile & ", strEXEParams - " & strEXEParams & ", strVbsDirectory - " & strVbsDirectory & DotFN(0))	
		
		If FnParamChecker(strEXEFile, "strEXEFile") = 1 Then
            call fnConst_VarFail("strEXEFile", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnInstallEXE = 1
		End If
				
		If FnParamChecker(strEXEParams, "strEXEParams") = 1 Then
			fnAddLog(VarsFNErr(0) & "strEXEParams" & VarsFNErr(1)) 
			
			     strEXEParams = ""
        
            fnAddLog(SetVarFN(0) & "strEXEParams" & SetVarFN(1) & "empty string" & DotFN(0))
		End If
				
		If FnParamChecker(strVbsDirectory, "strVbsDirectory") = 1 Then
			fnAddLog(VarsFNErr(0) & "strVbsDirectory" & VarsFNErr(1)) 
			
			     strVbsDirectory = "\"
        
			fnAddLog(SetVarFN(0) & "strVbsDirectory" & SetVarFN(1) & strVbsDirectory & DotFN(0))
		End If	
		
		If FnParamChecker(strSuccessTable, "strSuccessTable") = 1 Then
            call fnConst_VarFail("strSuccessTable", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnInstallEXE = 1
		End If
		
		If FnParamChecker(strName, "strName") = 1 Then
			fnAddLog("FUNCTION INFO: variable strName. Setting to strPKGName - " & strPKGName & ".")
			strName = strPKGName
		End If

		If fnIsExistD(strVbsDirectory) = 0 Then
			fnAddLog("FUNCTION INFO: strVbsDirectory- " & strVbsDirectory & " exist. Going to next step.")
			If fnIsExist(strEXEFile, strVbsDirectory) = 0 Then
				fnAddLog("FUNCTION INFO: strEXEFile- " & strEXEFile & " exist. Going to next step.")
				
				strEXEInstallCommand = chr(34) & strVbsDirectory & "\" & strEXEFile & chr(34)  & " " & strEXEParams 
			
				fnAddLog("INFO: strEXEInstallCommand: " & strEXEInstallCommand)
				intRVal = objShell.Run(strEXEInstallCommand, 0, TRUE)
							
				fnAddLog("INFO: " & strEXEFile & " installation ended with intRVal = " & intRVal & ".")	
				
                arrRCList = Split(strSuccessTable,",")
                
                rValAfterCMD = intRVal
				For Each item In arrRCList
                If CStr(item) = CStr(rValAfterCMD) Then
                    If CStr(item) = "3010" then boolRebootContainer = True					
						intRVal = 0
                        
						If fnIsExistTag(strName) = 1 Then
                            call fnAddTag(strName)  
                        End If
                        
                        WScript.Sleep 5000
                        
                        fnAddLog("FUNCTION INFO: ending fnInstallEXE - success.") 
                        strLogTAB = strLogTABBak

                        fnInstallEXE = intRVal
                        Exit Function
					Else
                        WScript.Sleep 5000
						intRVal = 1						
					End If
				Next
				
					fnAddLog("FUNCTION INFO: ending fnInstallEXE - fail.") 	
					
					strLogTAB = strLogTABBak
					fnFinalize(1)
			Else
				fnAddLog("FUNCTION INFO: strEXEFile- " & strEXEFile & " don't exist. Exiting function.")
				fnAddLog("FUNCTION INFO: ending fnInstallEXE - fail.") 		
				
				strLogTAB = strLogTABBak
				fnInstallEXE = 1
				fnFinalize(1)
			End If
			
		Else
			fnAddLog("FUNCTION INFO: strVbsDirectory- " & strVbsDirectory & " don't exist. Exiting function.")
			fnAddLog("FUNCTION INFO: ending fnInstallEXE - fail.") 

			strLogTAB = strLogTABBak		
			fnInstallEXE = 1
			fnFinalize(1)
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Uninstall application using exe file with specified parameters. 
'!
'! @param  strEXEFile          Exe filename.
'! @param  strEXEParams        Additional cmd parameters to use with exe file durring uninstallation.
'! @param  strVbsDirectory        strVbsDirectory (path) wher exe file is placed.
'! @param  strSuccessTable     Succes exit codes for exe file (all codes separeted with ",").
'! @param  strName             Package strName (according naming convention), will be used to create or check TAGfile.
'!
'! @return intRVal (0,3010)
Function fnUninstallEXE(strEXEFile, strEXEParams, strVbsDirectory, strSuccessTable, strName)
        Dim EXEUninstallCommand, intRVal, item, arrRCList, rValAfterCMD
    Dim strCurrentFNName                   : strCurrentFNName = "fnUninstallEXE"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strEXEFile - " & strEXEFile & ", strEXEParams - " & strEXEParams & ", strVbsDirectory - " & strVbsDirectory & DotFN(0))
				
		If FnParamChecker(strEXEFile, "strEXEFile") = 1 Then
            call fnConst_VarFail("strEXEFile", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnUninstallEXE = 1
		End If
				
		If FnParamChecker(strEXEParams, "strEXEParams") = 1 Then
			fnAddLog(VarsFNErr(0) & "strEXEParams" & VarsFNErr(1)) 
			
			     strEXEParams = ""
        
            fnAddLog(SetVarFN(0) & "strEXEParams" & SetVarFN(1) & "empty string" & DotFN(0))
		End If
				
		If FnParamChecker(strVbsDirectory, "strVbsDirectory") = 1 Then
			fnAddLog(VarsFNErr(0) & "strVbsDirectory" & VarsFNErr(1)) 
			
			     strVbsDirectory = "\"
        
			fnAddLog(SetVarFN(0) & "strVbsDirectory" & SetVarFN(1) & strVbsDirectory & DotFN(0))
		End If	
		
		If FnParamChecker(strSuccessTable, "strSuccessTable") = 1 Then
            call fnConst_VarFail("strSuccessTable", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnUninstallEXE = 1
		End If
		
		If FnParamChecker(strName, "strName") = 1 Then
			fnAddLog(VarsFNErr(0) & "strName" & VarsFNErr(1)) 
			
			     strName = strPKGName
        
			fnAddLog(SetVarFN(0) & "strName" & SetVarFN(1) & strName & DotFN(0))
		End If
		
		If fnIsExistD(strVbsDirectory) = 0 Then
			fnAddLog("FUNCTION INFO: strVbsDirectory- " & strVbsDirectory & " exist. Going to next step.")
			If fnIsExist(strEXEFile, strVbsDirectory) = 0 Then
				fnAddLog("FUNCTION INFO: strEXEFile- " & strEXEFile & " exist. Going to next step.")
				
				EXEUninstallCommand = chr(34) & strVbsDirectory & "\" & strEXEFile  & chr(34) & " " & strEXEParams 
				fnAddLog("INFO: EXEUninstallCommand: " & EXEUninstallCommand)
				intRVal = objShell.Run(EXEUninstallCommand, 0, TRUE)
				fnAddLog("INFO: " & strEXEFile & " uninstallation ended with intRVal = " & intRVal & ".")
											
				arrRCList = Split(strSuccessTable,",")
                
                rValAfterCMD = intRVal
				For Each item In arrRCList
                    If CStr(item) = CStr(rValAfterCMD) Then
						If CStr(item) = "3010" then boolRebootContainer = True					
						intRVal = 0
                            
                        If fnIsExistTag(strName) = 0 Then
                            call fnRemoveTag(strName)  
                        End If
                        
                        WScript.Sleep 5000
                            
                        fnAddLog("FUNCTION INFO: ending fnUninstallEXE - success.") 
                        strLogTAB = strLogTABBak

                        fnUninstallEXE = intRVal
                        Exit Function
					Else
                        WScript.Sleep 5000
						intRVal = 1						
					End If
				Next
				
					fnAddLog("FUNCTION INFO: ending fnUninstallEXE - fail.") 

					strLogTAB = strLogTABBak	
                    fnFinalize(1)
			Else
				fnAddLog("FUNCTION INFO: strEXEFile- " & strEXEFile & " don't exist. Exiting script - fail.")
				fnAddLog("FUNCTION INFO: ending fnUninstallEXE - fail.") 

				strLogTAB = strLogTABBak			
				fnUninstallEXE = 1
                fnFinalize(1)
			End If
		Else
			fnAddLog("FUNCTION INFO: strVbsDirectory- " & strVbsDirectory & " don't exist. Exiting script - fail.")
			fnAddLog("FUNCTION INFO: ending fnUninstallEXE - fail.") 	

			strLogTAB = strLogTABBak		
			fnUninstallEXE = 1
            fnFinalize(1)
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Reads application uninstall ARP entry. 
'!
'! @param  strGUID             Registry Name
'! @param  strAppBit           Application bittness (can be different that real app bitteness), where in registry apliaction store it's uninstall data.
'!
'! @return ARPUninstallString
Function fnGetARPUninstallString(strGUID, strAppBit)
	Dim ARPUninstallV, RegKey
    Dim strCurrentFNName                   : strCurrentFNName = "fnGetARPUninstallString"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strGUID - " & strGUID & ", strAppBit - " & strAppBit & DotFN(0))
		
		If FnParamChecker(strGUID, "strGUID") = 1 Then
            call fnConst_VarFail("strGUID", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnGetARPUninstallString = 1
		End If
							'zmiana z true na 0 poniewaz wszystkie funkcje jako succes wychodza z 0 wiec powstawal blad logiczny
		If fnIs64BitOS(True) = 0 Then 
			fnAddLog("FUNCTION INFO: Operating system is 64-bit - going to next step.") 
			
			If FnParamChecker(strAppBit, "strAppBit") = 1 Then
                call fnConst_VarFail("strAppBit", strCurrentFNName)
				
				strLogTAB = strLogTABBak
				fnFinalize(1)
				fnGetARPUninstallString = 1
			End If	
			
			If strAppBit = "x64" Then 
				RegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID & "\UninstallString"
				fnAddLog("FUNCTION INFO: Application is 64-bit - registry key set to " & RegKey & ".") 
			ElseIf strAppBit = "x86" Then
				RegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID & "\UninstallString"
				fnAddLog("FUNCTION INFO: Application is 86-bit - registry key set to " & RegKey & ".") 
			Else	
				RegKey = NULL
			End If
		Else
			RegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strGUID & "\UninstallString"
			fnAddLog("FUNCTION INFO: Operating system is 86-bit - registry key set to " & RegKey & ".") 
		End If
		
		If IsNull(RegKey) = False Then
			ARPUninstallV = fnReadKey(RegKey)

			fnAddLog("FUNCTION INFO: " & strGUID & " uninstall string is - " & ARPUninstallV & ".")
		ElseIf IsNull(RegKey) = True Then
			ARPUninstallV = NULL
			fnAddLog("FUNCTION ERROR: variable strAppBit must be set to 'x64' or 'x86'.")
		End If  
							'zmieniony sposob zwracania wartosci'
		strLogTAB = strLogTABBak
		fnGetARPUninstallString = ARPUninstallV
End Function
'-------------------------------------------------------------------------------------------
'! Uninstall application using ARP uninstall string. 
'!
'! @param  strARPUString       String from fnGetARPUninstallString function (patch to .exe file).
'! @param  strExeParams       	Additional uninstall parameters.
'! @param  strName           	Package strName according naming convention.
'!
'! @return intRVal
Function fnARPUninstall(strARPUString, strExeParams, strName)
	Dim strEXEFileName, strEXEPath, A, X
    Dim strCurrentFNName                   : strCurrentFNName = "fnARPUninstall"
	Dim number							: number = 0
	Dim intRVal							: intRVal = 0
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strARPUString - " & strARPUString & ", strEXEParams - " & strEXEParams & ", strName - " & strName & DotFN(0))
		
		If FnParamChecker(strARPUString, "strARPUString") = 1 Then
            call fnConst_VarFail("strARPUString", strCurrentFNName) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnARPUninstall = 1
		End If
		
		If FnParamChecker(strEXEParams, "strEXEParams") = 1 Then
			     fnAddLog(VarsFNErr(0) & "strEXEParams" & VarsFNErr(1))
			         
                    strEXEParams = " "
        
			     fnAddLog(SetVarFN(0) & "strEXEParams" & SetVarFN(1) & "empty string" & DotFN(0))
		End If
		
		If FnParamChecker(strName, "strName") = 1 Then
			     fnAddLog(VarsFNErr(0) & "strName" & VarsFNErr(1))
			         
                    strName = strPKGName
        
			     fnAddLog(SetVarFN(0) & "strName" & SetVarFN(1) & strName & DotFN(0))
		End If
		
		
	'----- Getting arrEXE path and arrEXE filename from UninstallString:
		A = Split(strARPUString, "\")
		strEXEFileName = A(UBound(A))
		
		For Each X in A
			If number < UBound(A) Then
				If number = 0 Then
					strEXEPath = X
				Else
					strEXEPath = strEXEPath & "\" & X
				End If
			End If
			
			number = number + 1
		Next
		
		fnAddLog("FUNCTION INFO: EXEFullPath divided into: strEXEPath - " & strEXEPath & ", strEXEFileName - " & strEXEFileName & ".")
		
		
	'----- Getting rid of quotation marks from arrEXE path and arrEXE filename:
		If InStr(strEXEPath, chr(34)) Then
			strEXEPath = Replace(strEXEPath, chr(34), "")
		End If
		
		If InStr(strEXEFileName, chr(34)) Then
			strEXEFileName = Replace(strEXEFileName, chr(34), "")
		End If
		
		
	'----- Running fnUninstallEXE:
		intRVal = fnUninstallEXE(strEXEFileName, strExeParams, strEXEPath, strSuccessTable, strName)
		
		fnAddLog("FUNCTION INFO: ending fnARPUninstall.")
		
		strLogTAB = strLogTABBak
		fnARPUninstall = intRVal
End Function
'-------------------------------------------------------------------------------------------
'! Advanced arp uninstall function, split patch from parameters and uninstall. 
'!
'! @param  strARPUString       String from fnGetARPUninstallString function (patch to .exe file).
'! @param  strExeSeparator     Char or string - will be used to separate patch from ARP uninstall string.
'! @param  strParamsSeparator  Char or string - will be used to separate parameters from ARP uninstall string.
'! @param  strOptionalParams  	Additional uninstall parameters.
'! @param  strName  			Package strName according naming convention.
'!
'! @return intRVal
Function fnSplitARPStringAndUninstall(strARPUString, strExeSeparator, strParamsSeparator, strOptionalParams, strName)
    Dim ARPStringArray, EXEFullPath, strEXEParams, strEXEFileName, strEXEPath, A, X
    Dim strCurrentFNName                   : strCurrentFNName = "fnSplitARPStringAndUninstall"
    Dim number							: number = 0
	Dim intRVal							: intRVal = 0
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strARPUString - " & strARPUString & ", strExeSeparator - " & strExeSeparator & ", strParamsSeparator - " & strParamsSeparator & ", strOptionalParams - " & strOptionalParams & ", strName - " & strName & DotFN(0))
		
		If FnParamChecker(strARPUString, "strARPUString") = 1 Then
            call fnConst_VarFail("strARPUString", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnSplitARPStringAndUninstall = 1
		End If
		
		If FnParamChecker(strExeSeparator, "strExeSeparator") = 1 Then
			     fnAddLog(VarsFNErr(0) & "strExeSeparator" & VarsFNErr(1))
			         
                    strExeSeparator = ".exe"
        
			     fnAddLog(SetVarFN(0) & "strExeSeparator" & SetVarFN(1) & strExeSeparator & DotFN(0))
		End If
		
		If FnParamChecker(strParamsSeparator, "strParamsSeparator") = 1 Then
			     fnAddLog(VarsFNErr(0) & "strParamsSeparator" & VarsFNErr(1))
			         
                    strParamsSeparator = "-"
        
			     fnAddLog(SetVarFN(0) & "strParamsSeparator" & SetVarFN(1) & strParamsSeparator & DotFN(0))
		End If
		
		If FnParamChecker(strOptionalParams, "strOptionalParams") = 1 Then
			     fnAddLog(VarsFNErr(0) & "strOptionalParams" & VarsFNErr(1))
			         
                    strOptionalParams = " "
        
			     fnAddLog(SetVarFN(0) & "strOptionalParams" & SetVarFN(1) & "empty string" & DotFN(0))
		End If
		
		If FnParamChecker(strName, "strName") = 1 Then
			     fnAddLog(VarsFNErr(0) & "strName" & VarsFNErr(1))
			         
                    strName = strPKGName
        
			     fnAddLog(SetVarFN(0) & "strName" & SetVarFN(1) & strName & DotFN(0))
		End If
		
		
	'----- Getting arrEXE full path and arrEXE parameters from UninstallString:
		ARPStringArray = Split(strARPUString, strExeSeparator)
		EXEFullPath = ARPStringArray(0) & strExeSeparator
        strEXEParams = ARPStringArray(1)
		
            If Left(EXEFullPath, 1) = chr(34) Then		'Getting rid of intFirst quotation mark from arrEXE full path.
                EXEFullPath = Mid(EXEFullPath, 2)
            End If

            If Left(strEXEParams, 1) = chr(34) Then		'Getting rid of intFirst quotation marks from arrEXE parameters.
                strEXEParams = Mid(strEXEParams, 2)
            End If

            strEXEParams = LTrim(strEXEParams)				'Getting rid of leading spaces from arrEXE parameters.
		
		fnAddLog("FUNCTION INFO: strARPUString divided into: EXEFullPath - " & EXEFullPath & ", strEXEParams - " & strEXEParams & ".")
		
		
	'----- Getting arrEXE path and arrEXE filename from EXEFullPath:
		A = Split(EXEFullPath, "\")
		strEXEFileName = A(UBound(A))
		
		For Each X in A
			If number < UBound(A) Then
				If number = 0 Then
					strEXEPath = X
				Else
					strEXEPath = strEXEPath & "\" & X
				End If
			End If
			
			number = number + 1
		Next
		
		fnAddLog("FUNCTION INFO: EXEFullPath divided into: strEXEPath - " & strEXEPath & ", strEXEFileName - " & strEXEFileName & ".")
		
		
	'----- Running fnUninstallEXE:
		If NOT strOptionalParams = "" Then
			strEXEParams = strEXEParams & " " & strOptionalParams
			fnAddLog("FUNCTION INFO: strEXEParams supplemented by strOptionalParams, new strEXEParams value - " & strEXEParams & ".")
		End If
		
		intRVal = fnUninstallEXE(strEXEFileName, strEXEParams, strEXEPath, strSuccessTable, strName)
		
		fnAddLog("FUNCTION INFO: ending fnSplitARPStringAndUninstall.")
		
		strLogTAB = strLogTABBak
		fnSplitARPStringAndUninstall = intRVal
End Function