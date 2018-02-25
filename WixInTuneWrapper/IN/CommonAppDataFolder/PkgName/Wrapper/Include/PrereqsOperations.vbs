'****************************************************************************************************************************************************
' FUNCTIONS:
'	CheckJavaVersion(javaVersion, bitness, yourUpdateVer, forceUninstall)
'	CheckFlashVersion(flashVersion, PKGVersion)
' 
' FUNCTION EXAMPLE USAGE (IN and OUT):
' 	CheckJavaVersion("8", "32", "45", all)
' 		javaVersion - "8", "7", "6"
' 		bitness - "32", "64"
' 		yourUpdateVer - 0-131
' 		forceUninstall - all, new, old, false
' 						 all - uninstall ALL found Java versions
' 						 new - uninstall ALL found Java versions that are NEWER than yourUpdateVer
' 						 old - uninstall ALL found Java versions that are OLDER than yourUpdateVer
' 						 false - do not uninstall ANY Java version
' 	Return: 1 - fail,
' 			0 - no bigger version on machine, 
' 			arrMSI exit code - from uninstallation if forced, 
' 			3 - is newer version on machine if uninstall not forced 
' 
' 	CheckFlashVersion("plugin", "18.0.0.194")
' 		flashVersion - "plugin", "activex", "pepper"
' 		PKGVersion - version of Flash Player we are trying to install (usually should be used "strAppNumber" variable)
' 	Return: 1 - fail,
' 			0 - no Flash Player found,
' 			2 - is the same version on machine,
' 			3 - is older version on machine,
' 			4 - is newer version on machine
'***************************************************************************************************************************************************
'! Check system for installed Java version (older and newer), then perform proper defined action (e.g. uninstall ALL found Java versions).
'!
'! @param  javaVersion		Package Java major version (6, 7, 8).
'! @param  bitness			Package Java bitness (32, 64).
'! @param  yourUpdateVer	Package Java minor version (e.g.: 70, 101 etc.).
'! @param  forceUninstall	Force Java uninstall or only check for installed Java version (all - uninstall ALL versions, new - uninstall NEWER versions, old - uninstall OLDER versions, false - do not uninstall).
'!
'! @return 1 - fail, 0 - no bigger version on machine, arrMSI exit code - from uninstallation if forced, 3 - is newer version on machine if uninstall not forced
Function CheckJavaVersion(javaVersion, bitness, yourUpdateVer, forceUninstall)
	Dim guidsP_1, guidsP_2, guidsP_3, guidsK, guids, update, strKeyPathx64, strKeyPathx86
	Dim strCurrentFNName					: strCurrentFNName = "CheckJavaVersion"
	Dim intRVal							: intRVal = 0
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "javaVersion - " & javaVersion & ", bitness - " & bitness & ", yourUpdateVer - " & yourUpdateVer & ", forceUninstall - " & forceUninstall & DotFN(0))
		
		If FnParamChecker(javaVersion, "javaVersion") = 1 Then
			call fnConst_VarFail("javaVersion", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			CheckJavaVersion = 1
		ElseIf NOT javaVersion = "6" AND NOT javaVersion = "7" AND NOT javaVersion = "8" Then
			fnAddLog(VarsFNErr(0) & "javaVersion is out of range" & DotFN(0))
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			CheckJavaVersion = 1
		End If
		
		If FnParamChecker(bitness, "bitness") = 1 Then
			fnAddLog(VarsFNErr(0) & "bitness" & VarsFNErr(1))
			bitness = "32"
			fnAddLog(SetVarFN(0) & "bitness" & SetVarFN(1) & bitness & DotFN(0))
		ElseIf bitness <> "32" AND bitness <> "64" Then
			fnAddLog(VarsFNErr(0) & "bitness is out of range" & DotFN(0))
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			CheckJavaVersion = 1
		End If
		
		If FnParamChecker(yourUpdateVer, "yourUpdateVer") = 1 Then
			fnAddLog(VarsFNErr(0) & "yourUpdateVer" & VarsFNErr(2))
			fnAddLog(PreTxtFN(0) & "checking for ALL installed Java versions" & DotFN(0))
		End If
		
		If FnParamChecker(forceUninstall, "forceUninstall") = 1 Then
			fnAddLog(VarsFNErr(0) & "forceUninstall" & VarsFNErr(1))
			forceUninstall = false
			fnAddLog(SetVarFN(0) & "forceUninstall" & SetVarFN(1) & forceUninstall & DotFN(0))
		End If
		
		
		If javaVersion = "6" OR javaVersion = "7" Then
			guidsK = "FF}"
		ElseIf javaVersion = "8" Then
			guidsK = "F0}"
		End If
		
		guidsP_1 = "{26A24AE4-039D-4CA4-87B4-2F8" & bitness		'arrGUIDs beginning for Java versions: 6.0 (up to Update 95), 7.0 (up to Update 51), 8.0 (up to Update 92)
		guidsP_2 = "{26A24AE4-039D-4CA4-87B4-2F0" & bitness		'arrGUIDs beginning for Java versions: 7.0 (from Update 60 to Update 99)
		guidsP_3 = "{26A24AE4-039D-4CA4-87B4-2F" & bitness		'arrGUIDs beginning for Java versions: 6.0 (from Update 101), 7.0 (from Update 101), 8.0 (from Update 101)
		
		strKeyPathx64 = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
		strKeyPathx86 = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
		
		
'Check for Java versions 1.6:
	If javaVersion = "6" Then
		For update = 0 to 131
			If update < 10 Then
				guids = guidsP_1 & "1600" & update & guidsK
			ElseIf update >= 10 AND update < 100 Then
				guids = guidsP_1 & "160" & update & guidsK
			Else
				guids = guidsP_3 & "160" & update & guidsK
			End If
			
			If fnSimpleKeyExist(strKeyPathx64 & guids & "\" ) = 0 OR fnSimpleKeyExist(strKeyPathx86 & guids & "\" ) = 0 Then
				fnAddLog(PreTxtFN(0) & "Java 1" & DotFN(0) & javaVersion & DotFN(0) & update & " detected. Closing processes may be required" & DotFN(0))
				
				If forceUninstall = "all" OR (forceUninstall = "new" AND yourUpdateVer < update) OR (forceUninstall = "old" AND yourUpdateVer > update) Then
					intRVal = fnUninstallMSI(guids, guids, NULL)
					Call Const_FnEndrVal(strCurrentFNName, intRVal)
					
					strLogTAB = strLogTABBak
					CheckJavaVersion = intRVal
					
				ElseIf forceUninstall = false Then
					intRVal = 3
					Call Const_FnEndrVal(strCurrentFNName, intRVal)
					
					strLogTAB = strLogTABBak
					CheckJavaVersion = intRVal
					Exit Function
					
				End If
				
			End If
			
		Next
		
		
'Check for Java versions 1.7:
	ElseIf javaVersion = "7" Then
		For update = 0 to 121
			If update < 10 Then
				guids = guidsP_1 & "1700" & update & guidsK
			ElseIf update >= 10 AND update < 60 Then
				guids = guidsP_1 & "170" & update & guidsK
			ElseIf update >= 60 AND update < 100 Then
				guids = guidsP_2 & "170" & update & guidsK
			Else
				guids = guidsP_3 & "170" & update & guidsK
			End If
			
			If fnSimpleKeyExist(strKeyPathx64 & guids & "\" ) = 0 OR fnSimpleKeyExist(strKeyPathx86 & guids & "\" ) = 0 Then
				fnAddLog(PreTxtFN(0) & "Java 1" & DotFN(0) & javaVersion & DotFN(0) & update & " detected. Closing processes may be required" & DotFN(0))
				
				If forceUninstall = "all" OR (forceUninstall = "new" AND yourUpdateVer < update) OR (forceUninstall = "old" AND yourUpdateVer > update) Then
					intRVal = fnUninstallMSI(guids, guids, NULL)
					Call Const_FnEndrVal(strCurrentFNName, intRVal)
					
					strLogTAB = strLogTABBak
					CheckJavaVersion = intRVal
					
				ElseIf forceUninstall = false Then
					intRVal = 3
					Call Const_FnEndrVal(strCurrentFNName, intRVal)
					
					strLogTAB = strLogTABBak
					CheckJavaVersion = intRVal
					Exit Function
					
				End If
				
			End If
			
		Next
		
		
'Check for Java versions 1.8:
	ElseIf javaVersion = 8 Then
		For update = 0 to 112
			If update < 10 Then
				guids = guidsP_1 & "1800" & update & guidsK
			ElseIf update >= 10 AND update < 100 Then
				guids = guidsP_1 & "180" & update & guidsK
			Else
				guids = guidsP_3 & "180" & update & guidsK
			End If
			
			If fnSimpleKeyExist(strKeyPathx64 & guids & "\" ) = 0 OR fnSimpleKeyExist(strKeyPathx86 & guids & "\" ) = 0 Then
				fnAddLog(PreTxtFN(0) & "Java 1" & DotFN(0) & javaVersion & DotFN(0) & update & " detected. Closing processes may be required" & DotFN(0))
				
				If forceUninstall = "all" OR (forceUninstall = "new" AND yourUpdateVer < update) OR (forceUninstall = "old" AND yourUpdateVer > update) Then
					intRVal = fnUninstallMSI(guids, guids, NULL)
					Call Const_FnEndrVal(strCurrentFNName, intRVal)
					
					strLogTAB = strLogTABBak
					CheckJavaVersion = intRVal
					
				ElseIf forceUninstall = false Then
					intRVal = 3
					Call Const_FnEndrVal(strCurrentFNName, intRVal)
					
					strLogTAB = strLogTABBak
					CheckJavaVersion = intRVal
					Exit Function
					
				End If
				
			End If
			
		Next
		
	End If
	
	
	Call Const_FnEndrVal(strCurrentFNName, intRVal)
	
	strLogTAB = strLogTABBak
	CheckJavaVersion = intRVal
End Function
'-------------------------------------------------------------------------------------------
'! Check system for installed Flash Player version (older and newer).
'!
'! @param  flashVersion		Flash Player version type (plugin, activex, pepper)
'! @param  PKGVersion		Version of Flash Player we are trying to install (usually should be used "strAppNumber" variable)
'!
'! @return 1 - fail, 0 - no Flash Player found, 2 - is the same version on machine, 3 - is older version on machine, 4 - is newer version on machine
Function CheckFlashVersion(flashVersion, PKGVersion)
	Dim valueInstalled, valuePKG, a1, a2, a3, a4, b1, b2, b3, b4, strKeyPath, regValue
	Dim strCurrentFNName					: strCurrentFNName = "CheckFlashVersion"
	Dim intRVal							: intRVal = 0
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "flashVersion - " & flashVersion & ", PKGVersion - " & PKGVersion & DotFN(0))
		
		If FnParamChecker(flashVersion, "flashVersion") = 1 Then
			fnAddLog(VarsFNErr(0) & "flashVersion" & VarsFNErr(1))
			flashVersion = "plugin"
			fnAddLog(SetVarFN(0) & "flashVersion" & SetVarFN(1) & flashVersion & DotFN(0))
		ElseIf NOT LCase(flashVersion) = "plugin" AND NOT LCase(flashVersion) = "activex" AND NOT LCase(flashVersion) = "pepper" Then
			fnAddLog(VarsFNErr(0) & "flashVersion is out of range" & DotFN(0))
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			CheckFlashVersion = 1
		End If
		
		If FnParamChecker(PKGVersion, "PKGVersion") = 1 Then
			call fnConst_VarFail("PKGVersion", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			CheckFlashVersion = 1
		End If
		
		
			If LCase(flashVersion) = "plugin" Then
				strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Macromedia\FlashPlayerPlugin\Version"
			ElseIf LCase(flashVersion) = "activex" Then
				strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Macromedia\FlashPlayerActiveX\Version"
			ElseIf LCase(flashVersion) = "pepper" Then
				strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Macromedia\FlashPlayerPepper\Version"
			End If
		
		
	'Check Flash Player version:
		regValue = fnReadKey(strKeyPath)
		
		If fnIsNull(regValue) = True OR isEmpty(regValue) = True OR regValue = "" Then	'Flash Player NOT found
			fnAddLog(PreTxtFN(0) & "No version of Flash Player " & flashVersion & " has been found" & DotFN(0))
			fnAddLog(EndingFN(0) & strCurrentFNName & EndingFN(1) & intRVal & DotFN(0))
			
			strLogTAB = strLogTABBak
			CheckFlashVersion = intRVal
			Exit Function
		Else																			'Flash Player found
		
		'Splitting Flash Player versions to 4 parts:
			valueInstalled = Split(regValue, ".")
				a1 = CInt(valueInstalled(0))
				a2 = CInt(valueInstalled(1))
				a3 = CLng(valueInstalled(2))
				a4 = CLng(valueInstalled(3))
			
			valuePKG = Split(PKGVersion, ".")
				b1 = CInt(valuePKG(0))
				b2 = CInt(valuePKG(1))
				b3 = CLng(valuePKG(2))
				b4 = CLng(valuePKG(3))
			
		'Comparing Flash Player versions:
			If a1 = b1 Then
				If a2 = b2 Then
					If a3 = b3 Then
						If a4 = b4 Then
							intRVal = 2
						ElseIf a4 < b4 Then
							intRVal = 3
						ElseIf a4 > b4 Then
							intRVal = 4
						End If
					ElseIf a3 < b3 Then
						intRVal = 3
					ElseIf a3 > b3 Then
						intRVal = 4
					End If
				ElseIf a2 < b2 Then
					intRVal = 3
				ElseIf a2 > b2 Then
					intRVal = 4
				End If
			ElseIf a1 < b1 Then
				intRVal = 3
			ElseIf a1 > b1 Then
				intRVal = 4
			End If
		
		End If
		
		
		If intRVal = 2 Then
			fnAddLog(PreTxtFN(0) & "Flash Player " & flashVersion & " - version: " & regValue & " has been found. It is the same version as this package" & DotFN(0))
		ElseIf intRVal = 3 Then
			fnAddLog(PreTxtFN(0) & "Flash Player " & flashVersion & " - version: " & regValue & " has been found. It is older version than this package" & DotFN(0))
		ElseIf intRVal = 4 Then
			fnAddLog(PreTxtFN(0) & "Flash Player " & flashVersion & " - version: " & regValue & " has been found. It is newer version than this package" & DotFN(0))
		End If
		
		
		fnAddLog(EndingFN(0) & strCurrentFNName & EndingFN(1) & intRVal & DotFN(0))
		
		strLogTAB = strLogTABBak
		CheckFlashVersion = intRVal
End Function
'-------------------------------------------------------------------------------------------