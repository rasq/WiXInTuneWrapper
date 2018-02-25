'****************************************************************************************************************************************************
' FUNCTIONS: 
'	fnInstallMSI(strMSIFile, strMSTFile, strMSPFile, strAdditionalParams, strVbsDirectory, strPKGName) 
'	fnTestingLoudInstallMSI(strMSIFile, strMSTFile, strMSPFile, strAdditionalParams, strVbsDirectory, strPKGName) - only for testing vendor or package if we need loud installation.
'	fnUninstallMSI(strGUID, strPKGName, strAdditionalParams) 
'	fnIsIbmMsiPKGInstalled(strGUID, strVersion, PKGNameV) 
'   fnSearchForGUID(strRegKeyName)
'	fnIsMSIInstalled(strGUID, strVersion)  
'	fnIsInstalled64(strGUID, strVersion)
'	fnIsInstalled(strGUID, strVersion)
'   fnIsIsntalledDisplayName(strRegKeyName) 
'	
' FUNCTION EXAMPLE USAGE (IN and OUT):
' 
' fnIsIbmMsiPKGInstalled(strGUID, strVersion, PKGNameV) 
'	Return 0 - installed, 1 - non IBM pkg, 2 - not installed
' 
' fnIsInstalled64(strGUID, version) 
'	 true, false, 1 - fail
' 
' fnIsInstalled(strGUID, version) 
'	Return true, false, 1 - fail 	
'	
' Attention : use NULL where parameters not required. For example if there's no arrMSP and no additional parameters, and package is in the same folder like .vbs 
'  then example syntax should look like this :
' retCode = fnInstallMSI("vendor_package.msi","vendor_name_version_account_W7x64_EN_B1.mst",NULL,NULL,NULL)
'
'
' if package is in subfolder we put as strVbsDirectory : "Subfolder\" 
'
'
' fnSearchForGUID("Mozilla Firefox")
'	Return:	Array[0]="Mozilla Firefox 52.1.1 ESR (x64 de)"
'			Array[1]="Mozilla Firefox 52.3.0 ESR (x86 en-US)"
' 
'	Possible return values: array	-	success, one or more arrGUIDs have been found
'							0		-	success, no arrGUIDs have been found
'							1		-	fail
'***************************************************************************************************************************************************
'! Install specified *.msi file with or without *.mst file and patch. Patch can be installed without *.msi and *.mst. 
'!
'! @param  strMSIFile          arrMSI file strName.
'! @param  strMSTFile          arrMST file strName.
'! @param  strMSPFile          arrMSP file strName.
'! @param  strAdditionalParams Additional cmd parameters to use with msiexec durring installation.
'! @param  strVbsDirectory        strVbsDirectory (path) where msi/mst/msp file are placed.
'! @param  strPKGName             Package strName (according naming convention), will be used to create or check TAGfile.
'!
'! @return intRVal
Function fnInstallMSI(strMSIFile, strMSTFile, strMSPFile, strAdditionalParams, strVbsDirectory, strPKGName)
	Dim intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnInstallMSI"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strMSIFile - " & strMSIFile & ", strMSTFile - " & strMSTFile & ", strMSPFile - " & strMSPFile & ", strAdditionalParams - " & strAdditionalParams & ", strVbsDirectory - " & strVbsDirectory & ", strPKGName - " & strPKGName & DotFN(0)) 
    
		fnAddLog("FUNCTION INFO: going to silent arrMSI installation - InstallMSICore.")	
		
		intRVal = InstallMSICore(strMSIFile, strMSTFile, strMSPFile, strAdditionalParams, strVbsDirectory, strPKGName, False)
		
			call Const_FnEndrVal(strCurrentFNName, intRVal) 
			
			strLogTAB = strLogTABBak
			fnInstallMSI = intRVal
End Function 
'-------------------------------------------------------------------------------------------
'! Install specified *.msi file with or without *.mst file and patch. Patch can be installed without *.msi and *.mst. This is DEBUG function, only for testing.
'!
'! @param  strMSIFile          arrMSI file strName.
'! @param  strMSTFile          arrMST file strName.
'! @param  strMSPFile          arrMSP file strName.
'! @param  strAdditionalParams Additional cmd parameters to use with msiexec durring installation.
'! @param  strVbsDirectory        strVbsDirectory (path) where msi/mst/msp file are placed.
'! @param  strPKGName             Package strName (according naming convention), will be used to create or check TAGfile.
'!
'! @return intRVal
Function fnTestingLoudInstallMSI(strMSIFile, strMSTFile, strMSPFile, strAdditionalParams, strVbsDirectory, strPKGName)
	Dim intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnTestingLoudInstallMSI"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strMSIFile - " & strMSIFile & ", strMSTFile - " & strMSTFile & ", strMSPFile - " & strMSPFile & ", strAdditionalParams - " & strAdditionalParams & ", strVbsDirectory - " & strVbsDirectory & ", strPKGName - " & strPKGName & DotFN(0)) 
    
		fnAddLog("FUNCTION INFO: going to loud arrMSI installation - InstallMSICore.")	
		
		MsgBox "You are starting fnTestingLoudInstallMSI, please change this function to fnInstallMSI for production version. This function can be used only for tests.", vbCritical 
		
		intRVal = InstallMSICore(strMSIFile, strMSTFile, strMSPFile, strAdditionalParams, strVbsDirectory, strPKGName, True)
		
			call Const_FnEndrVal(strCurrentFNName, intRVal)
			
			strLogTAB = strLogTABBak
			fnTestingLoudInstallMSI = intRVal
End Function 
'-------------------------------------------------------------------------------------------
'! Install specified *.msi file with or without *.mst file and patch. Patch can be installed without *.msi and *.mst. 
'!
'! @param  strMSIFile          arrMSI file strName.
'! @param  strMSTFile          arrMST file strName.
'! @param  strMSPFile          arrMSP file strName.
'! @param  strAdditionalParams Additional cmd parameters to use with msiexec durring installation.
'! @param  strVbsDirectory        strVbsDirectory (path) where msi/mst/msp file are placed.
'! @param  strPKGName             Package strName (according naming convention), will be used to create or check TAGfile.
'! @param  Loud             If application is installing thrue loud or standard arrMSI install function.
'!
'! @return intRVal
Function InstallMSICore(strMSIFile, strMSTFile, strMSPFile, strAdditionalParams, strVbsDirectory, strPKGName, Loud)
    Dim strMsiInstallCommand, boolOnlyMSP, arrMSTFileArray
    Dim strCurrentFNName                   : strCurrentFNName = "InstallMSICore"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		boolOnlyMSP = false
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strMSIFile - " & strMSIFile & ", strMSTFile - " & strMSTFile & ", strMSPFile - " & strMSPFile & ", strAdditionalParams - " & strAdditionalParams & ", strVbsDirectory - " & strVbsDirectory & ", strPKGName - " & strPKGName & DotFN(0)) 
		
		If FnParamChecker(strMSIFile, "strMSIFile") = 1 Then
			fnAddLog("FUNCTION ERROR: variable strMSIFile cannot be NULL.")
			fnAddLog("FUNCTION INFO: setting strMSIFile as empty string.") 
			
			strMSIFile = ""
			
			boolOnlyMSP = true
			
				If FnParamChecker(strMSPFile, "strMSPFile") = 1 Then
                    call fnConst_VarFail("strMSPFile", strCurrentFNName)
					strMSPFile = ""
					
					strLogTAB = strLogTABBak
					fnFinalize(1)
					InstallMSICore = 1
				End If	
		End If
				
		If FnParamChecker(strMSTFile, "strMSTFile") = 1 Then
            fnAddLog(VarsFNErr(0) & "strMSTFile" & VarsFNErr(1))
			         
                strMSTFile = ""
        
            fnAddLog(SetVarFN(0) & "strMSTFile" & SetVarFN(1) & "empty string" & DotFN(0))
		End If
				
		If FnParamChecker(strMSPFile, "strMSPFile") = 1 Then
			if boolOnlyMSP = false Then
                fnAddLog(VarsFNErr(0) & "strMSPFile" & VarsFNErr(1))

                    strMSPFile = ""

                fnAddLog(SetVarFN(0) & "strMSPFile" & SetVarFN(1) & "empty string" & DotFN(0))
			Else	
				strLogTAB = strLogTABBak
				fnFinalize(1)
				InstallMSICore = 1
			End If
		End If	
		
		If FnParamChecker(strAdditionalParams, "strAdditionalParams") = 1 Then
            fnAddLog(VarsFNErr(0) & "strAdditionalParams" & VarsFNErr(1))

                strAdditionalParams = ""

            fnAddLog(SetVarFN(0) & "strAdditionalParams" & SetVarFN(1) & "empty string" & DotFN(0))
		End If
		
		If FnParamChecker(strVbsDirectory, "strVbsDirectory") = 1 Then
			fnAddLog(VarsFNErr(0) & "strVbsDirectory" & VarsFNErr(1))
			         
                strVbsDirectory = strScriptPath
        
            fnAddLog(SetVarFN(0) & "strVbsDirectory" & SetVarFN(1) & strVbsDirectory & DotFN(0))
		End If
		
		If FnParamChecker(strPKGName, "strPKGName") = 1 Then
			fnAddLog(VarsFNErr(0) & "strPKGName" & VarsFNErr(1))
			         
                strPKGName = strPKGName
        
            fnAddLog(SetVarFN(0) & "strPKGName" & SetVarFN(1) & strPKGName & DotFN(0))
		End If
		
		If fnIsExistD(strVbsDirectory) = 0 Then
			strMsiInstallCommand = "msiexec.exe /i " & chr(34) & strVbsDirectory 
		
			if boolOnlyMSP = false Then
				if strMSIFile <> "" Then
					If fnIsExist(strMSIFile, strVbsDirectory) = 0 Then
						strMsiInstallCommand = strMsiInstallCommand & "\" & strMSIFile & Chr(34) ' CHANGE : Additional quotes added to fix
					Else
						fnAddLog("FUNCTION ERROR: strMSIFile - " & strMSIFile & " doesn't exist. ")
						fnAddLog("FUNCTION INFO: ending InstallMSICore - fail.") 
						
						strLogTAB = strLogTABBak
						fnFinalize(1)
					End If
				End If
				
				If strMSTFile <> "" Then
					if InStr(strMSTFile, ";") <> 0 then									' check if there is more than one arrMST file
						Dim i, MST_array
						
						fnAddLog("FUNCTION INFO: There is more than one arrMST file.")
						MST_array = Split(strMSTFile, ";", -1, 1)							' find everything arrMST files in strMSTFile string and write they into MST_array
						For i = 0 to UBound(MST_array)
							If fnIsExist(MST_array(i), strVbsDirectory) <> 0 Then
								fnAddLog("FUNCTION ERROR: strMSTFile - " & MST_array(i) & " doesn't exist. ")
								fnAddLog("FUNCTION INFO: ending InstallMSICore - fail.")
								fnFinalize(1)
							end if
                        
                            if i = 0 Then
                                arrMSTFileArray = strVbsDirectory & "\" & MST_array(i)
                            Else
                                arrMSTFileArray = arrMSTFileArray & ";" & strVbsDirectory & "\" & MST_array(i)
                            End If
						Next
						
						strMsiInstallCommand = strMsiInstallCommand & " TRANSFORMS=" & chr(34) & arrMSTFileArray & Chr(34)
					else
						If fnIsExist(strMSTFile, strVbsDirectory) = 0 Then
							strMsiInstallCommand = strMsiInstallCommand & " TRANSFORMS=" & chr(34) & strVbsDirectory & "\" & strMSTFile & Chr(34) 
						Else
							fnAddLog("FUNCTION ERROR: strMSTFile - " & strMSTFile & " doesn't exist. ")
							fnAddLog("FUNCTION INFO: ending InstallMSICore - fail.") 
							
							strLogTAB = strLogTABBak
							fnFinalize(1)
						End If
					End If
				End If
				
				If strMSPFile <> "" Then
					If fnIsExist(strMSPFile, strVbsDirectory) = 0 Then
						strMsiInstallCommand = strMsiInstallCommand & " PATCH=" & chr(34) & strVbsDirectory & "\" & strMSPFile & chr(34)
					Else
						fnAddLog("FUNCTION ERROR: strMSPFile - " & strMSPFile & " doesn't exist. ")
						fnAddLog("FUNCTION INFO: ending InstallMSICore - fail.") 
						
						strLogTAB = strLogTABBak
						fnFinalize(1)
					End If
				End If
			Else 'install only arrMSP
				If strMSPFile <> "" Then
					If fnIsExist(strMSPFile, strVbsDirectory) = 0 Then
						'strMsiInstallCommand = strMsiInstallCommand & " PATCH=" & chr(34) & strVbsDirectory & "\" & strMSPFile & chr(34)
                        strMsiInstallCommand = "msiexec.exe /update " & chr(34) & strVbsDirectory & "\" & strMSPFile & chr(34)
					Else
						fnAddLog("FUNCTION ERROR: strMSPFile - " & strMSPFile & " doesn't exist. ")
						fnAddLog("FUNCTION INFO: ending InstallMSICore - fail.") 
						
						strLogTAB = strLogTABBak
						fnFinalize(1)
					End If
				End If
			End If
			
			
				If Loud = false Then
					strMsiInstallCommand = strMsiInstallCommand & " /l*v+ " & chr(34) & strLogFolder & "\" & strPKGName & "\" & strPKGName & "_Install.log" & chr(34) & " /qn REBOOT=ReallySuppress " & strAdditionalParams ' CHANGE : strPKGName instead of Appname  
				Else
					strMsiInstallCommand = strMsiInstallCommand & " /l*v+ " & chr(34) & strLogFolder & "\" & strPKGName & "\" & strPKGName & "_Install.log" & chr(34) & " " & strAdditionalParams
				End If
				fnAddLog("INFO: strMsiInstallCommand: " & strMsiInstallCommand)
				intRVal = objShell.Run(strMsiInstallCommand, 0, TRUE)
				call Const_FnEndrVal(strCurrentFNName, intRVal)	
				
				if intRVal <> 0 And intRVal <> 3010 Then
					fnAddLog("FUNCTION INFO: ending InstallMSICore - fail.") 
					
					strLogTAB = strLogTABBak
					fnFinalize(1)		
					InstallMSICore = 1
				Else
					If intRVal = 3010 then boolRebootContainer = True
					
					If fnIsExistTag(strPKGName) = 1 Then
						call fnAddTag(strPKGName)	
					End If
					
					'fnAddLog("FUNCTION INFO: self checking InstallMSICore - if arrMSI is properly installed.") 
					'intRVal = fnIsIbmMsiPKGInstalled(strGUID, strVersion, strPKGName) 
					'if intRVal <> 0 Then
					'	fnAddLog("FUNCTION INFO: InstallMSICore self check - fail. arrMSI is not properly installed. Please try again or check your arrMSI/arrMST/Wrap.")
					'	fnAddLog("FUNCTION INFO: ending InstallMSICore - fail.") 
					'	strLogTAB = strLogTABBak
					'	fnFinalize(1)		
					'	InstallMSICore = 1
					'	Exit Function
					'End If
				
                    WScript.Sleep 5000
                
					fnAddLog("FUNCTION INFO: ending InstallMSICore - success.")

					strLogTAB = strLogTABBak	
				
					InstallMSICore = intRVal
					Exit Function
				End If
		Else 
			fnAddLog("FUNCTION INFO: strVbsDirectory- " & strVbsDirectory & " don't exist. Exiting function.")
			fnAddLog("FUNCTION INFO: ending InstallMSICore - fail.") 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)		
			InstallMSICore = 1
		End If
End Function 
'-------------------------------------------------------------------------------------------
'! Buildin function for fnIsIbmMsiPKGInstalled. Please use fnIsIbmMsiPKGInstalled instead.
'!
'! @param  strGUID             strGUID or registry strName in \CurrentVersion\Uninstall.
'! @param  strVersion          Application strVersion from ARP.
'!
'! @return true,false,1
Function fnIsMSIInstalled(strGUID, strVersion) 
    Dim strCurrentFNName                   : strCurrentFNName = "fnIsMSIInstalled"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strGUID - " & strGUID & ", strVersion - " & strVersion & DotFN(0)) 	
		
		If FnParamChecker(strGUID, "strGUID") = 1 Then
            call fnConst_VarFail("strGUID", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnIsMSIInstalled = 1
		End If
				
		If FnParamChecker(strVersion, "strVersion") = 1 Then
            fnAddLog(VarsFNErr(0) & "strVersion" & VarsFNErr(1))

                strVersion = ""

            fnAddLog(SetVarFN(0) & "strVersion" & SetVarFN(1) & "empty string" & DotFN(0))
		End If
		
		If fnIsInstalled64(strGUID, strVersion) = false Then
			fnAddLog("FUNCTION INFO: fnIsInstalled64 = false, going next step.")
				If fnIsInstalled(strGUID, strVersion) = false Then
					fnAddLog("FUNCTION INFO: fnIsInstalled = false, going next step.")
					fnAddLog("FUNCTION INFO: ending fnIsMSIInstalled, strGUID - " & strGUID & " strVersion - " & strVersion & " not present.") 

					strLogTAB = strLogTABBak					
					fnIsMSIInstalled = 1
					Exit Function
				Else
					fnAddLog("FUNCTION INFO: fnIsInstalled("& strGUID & ", " & strVersion & ") = true. Exiting function.")
					fnAddLog("FUNCTION INFO: ending fnIsMSIInstalled - success.") 	
					
					strLogTAB = strLogTABBak
					fnIsMSIInstalled = 0
					Exit Function
				End If
		Else
			fnAddLog("FUNCTION INFO: fnIsInstalled64("& strGUID & ", " & strVersion & ") = true. Exiting function.")
			fnAddLog("FUNCTION INFO: ending fnIsMSIInstalled - success.") 	
			
			strLogTAB = strLogTABBak
			fnIsMSIInstalled = 0
			Exit Function
		End If
		
		fnAddLog("FUNCTION INFO: ending fnIsMSIInstalled, strGUID - " & strGUID & " strVersion - " & strVersion & " not present.") 

		strLogTAB = strLogTABBak		
		fnIsMSIInstalled = 1
End Function
'-------------------------------------------------------------------------------------------
'! Search for application arrGUIDs (both x86 and x64), that contains specific string in its strName and return an array of found arrGUIDs.
'!
'! @param	strRegKeyName		strPKGName of the registry key. Can be strGUID or part of strGUID (in case the application has similar strGUID with different beginning/ending, depending on its version)
'!
'! @return array,0,1
Function fnSearchForGUID(strRegKeyName)
	Const HKEY_LOCAL_MACHINE = &H80000002
	Dim strSubKey, arrSubKeys, arrSubKeys32, arrSubKeys64, arrGUIDs
	Dim i : i = -1
	Dim j : j = 0
	Dim k
	Dim strCurrentFNName					: strCurrentFNName = "fnSearchForGUID"
	Dim strFail							: strFail = 0
	Dim intRVal							: intRVal = 0
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKeyName" & DotFN(3) & strRegKeyName & DotFN(0))
		
		If FnParamChecker(strRegKeyName, "strRegKeyName") = 1 Then
			call fnConst_VarFail("strRegKeyName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnSearchForGUID = 1
		End If
		
		
		If fnIs64BitOS(True) = 0 Then		'Searching for arrGUIDs in two locations for 64-bit OS
			objReg.EnumKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall", arrSubKeys32
			objReg.EnumKey HKEY_LOCAL_MACHINE, "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall", arrSubKeys64
			arrSubKeys = joinArray(arrSubKeys32, arrSubKeys64)
		Else							'Searching for arrGUIDs in one location for 32-bit OS
			objReg.EnumKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall", arrSubKeys
		End If
		
		
		For Each strSubKey In arrSubKeys	'Determining number of elements in array
			If InStr(strSubKey, Keyname) <> 0 Then
				i = i + 1
			End If
		Next
		
		ReDim arrGUIDs(i)		'Declaring dynamic-array variable
		
		For Each strSubKey In arrSubKeys	'Adding found arrGUIDs (that meets the criteria) to array
			If InStr(strSubKey, Keyname) <> 0 Then
				fnAddLog(PreTxtFN(0) & "String" & DotFN(3) & strRegKeyName & " is included in strGUID" & DotFN(3) & strSubKey & DotFN(0))
				
				Err.Clear
				arrGUIDs(j) = strSubKey
				
				If Err.Number = 0 Then
					fnAddLog(PreTxtFN(0) & "strGUID" & DotFN(3) & strSubKey & " has been successfully added to array.")
					j = j + 1
				Else
					fnAddLog(PreTxtFN(1) & "strGUID" & DotFN(3) & strSubKey & " has NOT been added to array.")
					strFail = strFail + 1
				End If
				
			End If
		Next
		
		
		fnAddLog(PreTxtFN(0) & "Array elements:")		'Self-checking array elements
		For k = 0 to i
			fnAddLog(PreTxtFN(0) & "arrGUIDs[" & k & "]: " & arrGUIDs(k))
		Next
		
		
		If strFail <> 0 Then	'Finalizing
			fnAddLog(EndingFN(2) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnSearchForGUID = 1
		Else
			If UBound(arrGUIDs) >= 0 Then
				fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0))
				fnSearchForGUID = arrGUIDs
			Else
				fnAddLog(EndingFN(0) & strCurrentFNName & DotFN(2) & " No matching arrGUIDs have been found.")
				fnSearchForGUID = intRVal
			End If
		End If
		
		strLogTAB = strLogTABBak
End Function
'-------------------------------------------------------------------------------------------
'! Test if our package/application is installed and if this is IBM package or Vendor application. Can be used for msi or exe based installators. 
'!
'! @param  strGUID             strGUID or registry strName in \CurrentVersion\Uninstall.
'! @param  strVersion          Application strVersion from ARP.
'! @param  PKGNameV         Package strName according naming convention.
'!
'! @return intRVal
Function fnIsIbmMsiPKGInstalled(strGUID, strVersion, PKGNameV) 
    Dim strCurrentFNName                   : strCurrentFNName = "fnIsIbmMsiPKGInstalled"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strGUID - " & strGUID & ", strVersion - " & strVersion & ", PKGNameV - " & PKGNameV & DotFN(0)) 	
		
		If FnParamChecker(strGUID, "strGUID") = 1 Then
            call fnConst_VarFail("strGUID", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnIsIbmMsiPKGInstalled = 1
		End If
				
		If FnParamChecker(strVersion, "strVersion") = 1 Then
            fnAddLog(VarsFNErr(0) & "strVersion" & VarsFNErr(1))

                strVersion = ""

            fnAddLog(SetVarFN(0) & "strVersion" & SetVarFN(1) & "empty string" & DotFN(0))
		End If
				
		If FnParamChecker(PKGNameV, "PKGNameV") = 1 Then
            fnAddLog(VarsFNErr(0) & "PKGNameV" & VarsFNErr(1))
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnIsIbmMsiPKGInstalled = 1
		End If
		
		If fnIsExistTag(PKGNameV) = 0 Then
			fnAddLog("FUNCTION INFO: TAG for PKG - " & PKGNameV & " exist, going to check if " & strGUID & " is installed.")
				if fnIsMSIInstalled(strGUID, strVersion) = 0 Then
					fnAddLog("FUNCTION INFO: IBM PKG - " & PKGNameV & " is installed.")
					fnAddLog("FUNCTION INFO: ending fnIsIbmMsiPKGInstalled - success.") 
					
					strLogTAB = strLogTABBak
					fnIsIbmMsiPKGInstalled = 0 
					Exit Function
				Else
					fnAddLog("FUNCTION INFO: IBM PKG - " & PKGNameV & " not installed. There is TAG on workstation.")
                        
                        fnAddLog("FUNCTION INFO: trying to clean TAG file.")
                        call fnRemoveTag(PKGNameV) 
                    
					fnAddLog("FUNCTION INFO: ending fnIsIbmMsiPKGInstalled.") 
					
					strLogTAB = strLogTABBak
					fnIsIbmMsiPKGInstalled = 2 
					Exit Function
				End If
		Else
			fnAddLog("FUNCTION INFO: TAG for PKG - " & PKGNameV & " don't exist, going to check if " & strGUID & " is installed.")
				if fnIsMSIInstalled(strGUID, strVersion) = 0 Then
					fnAddLog("FUNCTION INFO: strGUID - " & strGUID & " is installed. This is not an IBM PKG.")
					fnAddLog("FUNCTION INFO: ending fnIsIbmMsiPKGInstalled.")
					
					strLogTAB = strLogTABBak
					fnIsIbmMsiPKGInstalled = 1
					Exit Function
				Else
					fnAddLog("FUNCTION INFO: IBM PKG - " & PKGNameV & " not installed.")
					fnAddLog("FUNCTION INFO: ending fnIsIbmMsiPKGInstalled.")
					
					strLogTAB = strLogTABBak
					fnIsIbmMsiPKGInstalled = 2
					Exit Function 
				End If
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Uninstall msi based application. 
'!
'! @param  strGUID             strGUID or registry strName in \CurrentVersion\Uninstall.
'! @param  strAdditionalParams Additional cmd parameters to use with msiexec durring uninstallation.
'! @param  strVersion          Application strVersion from ARP.
'!
'! @return intRVal
Function fnUninstallMSI(strGUID, strPKGName, strAdditionalParams) 
	Dim strSUninstallCommand, intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnUninstallMSI"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator	
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strGUID - " & strGUID & ", strPKGName - " & strPKGName & DotFN(0)) 	
        
		
		If FnParamChecker(strGUID, "strGUID") = 1 Then
            call fnConst_VarFail("strGUID", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnUninstallMSI = 1
		End If
				
		If FnParamChecker(strPKGName, "strPKGName") = 1 Then
			fnAddLog(VarsFNErr(0) & "strPKGName" & VarsFNErr(1))
			         
                strPKGName = strPKGName
        
            fnAddLog(SetVarFN(0) & "strPKGName" & SetVarFN(1) & strPKGName & DotFN(0))
		End If
				
		If FnParamChecker(strAdditionalParams, "strAdditionalParams") = 1 Then
            fnAddLog(VarsFNErr(0) & "strAdditionalParams" & VarsFNErr(1))

                strAdditionalParams = ""

            fnAddLog(SetVarFN(0) & "strAdditionalParams" & SetVarFN(1) & "empty string" & DotFN(0))
		End If

		if isInstalled64(guid, version) = false And isInstalled(guid, version) = false Then
			fnAddLog("FUNCTION INFO: guid " & guid & " absent, going to next step.")
            fnAddLog("FUNCTION INFO: ending fnUninstallMSI - success.")
			
			strLogTAB = strLogTABBak
			fnUninstallMSI = 0
			Exit Function
		Else
			strSUninstallCommand = "MsiExec.exe /X " & chr(34) & strGUID & chr(34) & strAdditionalParams & " REBOOT=ReallySuppress /l*v+ " & chr(34) & strLogFolder & "\" & strPKGName & "\" & strPKGName & "_Uninstall.log" & chr(34) & " /qn"
																																			'AO - added REBOOT=ReallySuppress in case vendor arrMSI uninstall causes machine to restart
			fnAddLog("INFO: strSUninstallCommand: " & strSUninstallCommand)
			intRVal = objShell.Run(strSUninstallCommand, 0, TRUE)
			call Const_FnEndrVal(strCurrentFNName, intRVal)
					
			if intRVal <> 0 And intRVal <> 3010 Then
				fnAddLog("FUNCTION INFO: ending fnUninstallMSI - fail.") 
				
				strLogTAB = strLogTABBak
				fnFinalize(1)		
				fnUninstallMSI = 1
			Else
				If intRVal = 3010 then boolRebootContainer = True
			
				If fnIsExistTag(strPKGName) = 0 Then
					call fnRemoveTag(strPKGName)  
				End If
				
                WScript.Sleep 5000
				fnAddLog("FUNCTION INFO: ending fnUninstallMSI - success.") 

				strLogTAB = strLogTABBak				
				fnUninstallMSI = intRVal
				Exit Function
			End If
		End if
End Function
'-------------------------------------------------------------------------------------------
'! Buildin function for fnIsIbmMsiPKGInstalled. Please use fnIsIbmMsiPKGInstalled instead.
'!
'! @param  strGUID             strGUID or registry strName in \CurrentVersion\Uninstall.
'! @param  strVersion          Application strVersion from ARP.
'!
'! @return true,false,1
Function fnIsInstalled(strGUID, strVersion)
    Dim strCurrentFNName                   : strCurrentFNName = "fnIsInstalled"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator	
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strGUID - " & strGUID & ", strVersion - " & strVersion & DotFN(0)) 
		
				
		If FnParamChecker(strVersion, "strVersion") = 1 Then
            fnAddLog(VarsFNErr(0) & "strVersion" & VarsFNErr(1))

                strVersion = ""

            fnAddLog(SetVarFN(0) & "strVersion" & SetVarFN(1) & "empty string" & DotFN(0))
		End If
		
		
		If FnParamChecker(strGUID, "strGUID") = 1 Then
            call fnConst_VarFail("strGUID", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnIsInstalled = 1
		End If
				
		Dim strValue, strKeyPath, arrSubKeys, subkey, objReg
		Const HKEY_LOCAL_MACHINE = &H80000002
		Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
			strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" 
			objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys, subkey, strKeyPathV, strDisplayVersion
			isInstalled = false
			
		For Each subkey In arrSubKeys
			If (subkey = strGUID) Then
				If strVersion <> "" Then
					strKeyPathV = "HKEY_LOCAL_MACHINE\" & strKeyPath & strGUID & "\DisplayVersion"
					if fnSimpleKeyExist(strKeyPathV) = 0 Then
						strDisplayVersion = fnReadKey(strKeyPathV) 
						
						If strDisplayVersion = strVersion Then
							fnAddLog("FUNCTION INFO: ending, strGUID - " & strGUID & " exist, strVersion - " & strVersion & " present and confirmed.")
							fnAddLog("FUNCTION INFO: ending fnIsInstalled - success.")
						
							strLogTAB = strLogTABBak
							isInstalled = true
							Exit For
						Else
							fnAddLog("FUNCTION INFO: ending, strGUID - " & strGUID & " exist, strVersion - " & strVersion & " present but different from ARP DisplayVersion - " & strDisplayVersion & ".")
							fnAddLog("FUNCTION INFO: ending fnIsInstalled - fail.")
						
							strLogTAB = strLogTABBak
							isInstalled = false
							Exit For
						End If
					Else 
						fnAddLog("FUNCTION INFO: ending, strGUID - " & strGUID & " exist, strVersion - " & strVersion & " not present.")
						fnAddLog("FUNCTION INFO: ending fnIsInstalled - success.")
					
						strLogTAB = strLogTABBak
						isInstalled = true
						Exit For
					End If
				Else
					fnAddLog("FUNCTION INFO: ending, strGUID - " & strGUID & " exist, strVersion not checked.")
					fnAddLog("FUNCTION INFO: ending fnIsInstalled - success.")
				
					strLogTAB = strLogTABBak
					isInstalled = true
					Exit For
				End If
			End If
		Next
End Function
'-------------------------------------------------------------------------------------------
'! Buildin function for fnIsIbmMsiPKGInstalled. Please use fnIsIbmMsiPKGInstalled instead.
'!
'! @param  strGUID             strGUID or registry strName in \CurrentVersion\Uninstall.
'! @param  strVersion          Application strVersion from ARP.
'!
'! @return true,false,1
Function fnIsInstalled64(strGUID, strVersion)
	Dim strKeyPathV, strDisplayVersion
    Dim strCurrentFNName                   : strCurrentFNName = "fnIsInstalled64"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator		
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strGUID - " & strGUID & ", strVersion - " & strVersion & DotFN(0)) 	
		
				
		If FnParamChecker(strVersion, "strVersion") = 1 Then
            fnAddLog(VarsFNErr(0) & "strVersion" & VarsFNErr(1))

                strVersion = ""

            fnAddLog(SetVarFN(0) & "strVersion" & SetVarFN(1) & "empty string" & DotFN(0))
		End If
		
		
		If FnParamChecker(strGUID, "strGUID") = 1 Then
            call fnConst_VarFail("strGUID", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnIsInstalled64 = 1
		End If
		
		
		
		if fnIs64BitOS(false) = 1 Then
			fnAddLog("FUNCTION INFO: ending fnIsInstalled64 - this is 32bit OS.")
		
			strLogTAB = strLogTABBak
			isInstalled64 = false                      
			Exit Function      
		Else
			Dim strValue, strKeyPath, arrSubKeys, subkey, objReg1, objCtx, objLocator, objServices
			Const HKEY_LOCAL_MACHINE = &H80000002
			Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet") 
				objCtx.Add "__ProviderArchitecture", 64 
			Set objLocator = CreateObject("Wbemscripting.SWbemLocator") 
			Set objServices = objLocator.ConnectServer("","root\default","","",,,,objCtx) 
			Set objReg1 = objServices.Get("StdRegProv")
				strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" 
				objReg1.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys, subkey
				isInstalled64 = false
				
			For Each subkey In arrSubKeys
				If (subkey = strGUID) Then
					If strVersion <> "" Then
						strKeyPathV = "HKEY_LOCAL_MACHINE\" & strKeyPath & strGUID & "\DisplayVersion"
						if fnSimpleKeyExist(strKeyPathV) = 0 Then
							strDisplayVersion = fnReadKey(strKeyPathV) 
							
							if strDisplayVersion = strVersion Then
								fnAddLog("FUNCTION INFO: ending, strGUID - " & strGUID & " exist, strVersion - " & strVersion & " present and confirmed.")
								fnAddLog("FUNCTION INFO: ending isInstalled64 - success.")
							
								strLogTAB = strLogTABBak
								isInstalled64 = true
								Exit For
							Else
								fnAddLog("FUNCTION INFO: ending, strGUID - " & strGUID & " exist, strVersion - " & strVersion & " present but different from ARP DisplayVersion - " & strDisplayVersion & ".")
								fnAddLog("FUNCTION INFO: ending isInstalled64 - fail.")
							
								strLogTAB = strLogTABBak
								isInstalled64 = false
								Exit For
							End If
						Else 
							fnAddLog("FUNCTION INFO: ending, strGUID - " & strGUID & " exist, strVersion - " & strVersion & " not present.")
							fnAddLog("FUNCTION INFO: ending isInstalled64 - success.")
						
							strLogTAB = strLogTABBak
							isInstalled64 = true
							Exit For
						End If
					Else
						fnAddLog("FUNCTION INFO: ending, strGUID - " & strGUID & " exist, strVersion not checked.")
						fnAddLog("FUNCTION INFO: ending isInstalled64 - success.")
					
						strLogTAB = strLogTABBak
						isInstalled64 = true
						Exit For
					End If     
				End If
			Next
			' ADDED: \\WOW6432NODE searching support
			
			strKeyPath = "SOFTWARE\WOW6432NODE\Microsoft\Windows\CurrentVersion\Uninstall\" 
			objReg1.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys, subkey
				
			For Each subkey In arrSubKeys
				If (subkey = strGUID) Then
					If strVersion <> "" Then
                        strKeyPathV = "HKEY_LOCAL_MACHINE\" & strKeyPath & strGUID & "\DisplayVersion"
						if fnSimpleKeyExist(strKeyPathV) = 0 Then
							strDisplayVersion = fnReadKey(strKeyPathV) 
							
							if strDisplayVersion = strVersion Then
								fnAddLog("FUNCTION INFO: ending, strGUID - " & strGUID & " exist, strVersion - " & strVersion & " present and confirmed.")
								fnAddLog("FUNCTION INFO: ending isInstalled64 - success.")
							
								strLogTAB = strLogTABBak
								isInstalled64 = true
								Exit For
							Else
								fnAddLog("FUNCTION INFO: ending, strGUID - " & strGUID & " exist, strVersion - " & strVersion & " present but different from ARP DisplayVersion - " & strDisplayVersion & ".")
								fnAddLog("FUNCTION INFO: ending isInstalled64 - fail.")
							
								strLogTAB = strLogTABBak
								isInstalled64 = false
								Exit For
							End If
						Else 
							fnAddLog("FUNCTION INFO: ending, strGUID - " & strGUID & " exist, strVersion - " & strVersion & " not present.")
							fnAddLog("FUNCTION INFO: ending isInstalled64 - success.")
						
							strLogTAB = strLogTABBak
							isInstalled64 = true
							Exit For
						End If
					Else
						fnAddLog("FUNCTION INFO: ending, strGUID - " & strGUID & " exist, strVersion not checked.")
						fnAddLog("FUNCTION INFO: ending isInstalled64 - success.")
					
						strLogTAB = strLogTABBak
						isInstalled64 = true
						Exit For
					End If      
				End If
			Next
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Test if application is installed, not using strGUID but it's display strName.
'!
'! @param  strRegKeyName          DisplayName we are searching for.
'!
'! @return 0,1
Function fnIsIsntalledDisplayName(strRegKeyName) 
	Const HKEY_LOCAL_MACHINE =	&H80000002	
	Dim strSubKey, arrSubKeys, arrSubKeys32, arrSubKeys64
    Dim strCurrentFNName                   : strCurrentFNName = "fnIsIsntalledDisplayName"
	Dim intRVal 							: intRVal = 1
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator		
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strRegKeyName - " & strRegKeyName & DotFN(0)) 	
		
		If FnParamChecker(strRegKeyName, "strRegKeyName") = 1 Then
            call fnConst_VarFail("strRegKeyName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnIsIsntalledDisplayName = 1
		End If
		
		If fnIs64BitOS(True) = 0 Then
			objReg.EnumKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall", arrSubKeys32
			objReg.EnumKey HKEY_LOCAL_MACHINE, "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall", arrSubKeys64
			
			On Error resume next
			
			For Each strSubKey In arrSubKeys32			
				If InStr(objShell.RegRead ("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & strSubKey &"\DisplayName"),Keyname) <> 0 Then	
					if err.number = 0 Then
						fnIsIsntalledDisplayName = 0
						intRVal = 0
					End If 
					err.clear
				End If
			Next

			For Each strSubKey In arrSubKeys64
				If InStr(objShell.RegRead ("HKEY_LOCAL_MACHINE\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" & strSubKey&"\DisplayName"),Keyname) <> 0 Then	
					if err.number = 0 Then
						fnIsIsntalledDisplayName = 0
						intRVal = 0
					End If 
					err.clear
				End If
			Next

		Else
			objReg.EnumKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall", arrSubKeys
			
			For Each strSubKey In arrSubKeys
				If InStr(objShell.RegRead ("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & strSubKey & "\DisplayName"),Keyname) <> 0 Then
					if err.number = 0 Then
						fnIsIsntalledDisplayName = 0
						intRVal = 0
					End If 
					err.clear
				End If
			Next	
		End If		

		If intRVal <> 0 Then
			fnAddLog("FUNCTION INFO: ending fnIsIsntalledDisplayName - fail.") 
			strLogTAB = strLogTABBak			
			fnIsIsntalledDisplayName = 1
			Exit Function
		Else
			fnAddLog("FUNCTION INFO: ending fnIsIsntalledDisplayName - success.") 
			strLogTAB = strLogTABBak			
			fnIsIsntalledDisplayName = 0
			Exit Function
		End If
		
End Function