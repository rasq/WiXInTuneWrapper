'****************************************************************************************************************************************************	
' FUNCTIONS:  
'	fnReadIni( strMyFilePath, strMySection, strMyKey )
'
'	fnDirectoryPath (strDirectoryVar)
'			strDirectoryVar - e.g.: %windir%, %windir%\someDirectory
'
' FUNCTION EXAMPLE USAGE (IN and OUT):
'	fnReadIni( strMyFilePath, strMySection, strMyKey )
' 			[string] value for the specified key in the specified section, if key without value it return space
'
' 	fnDirectoryPath (strDirectoryVar)
'			[string] with path from variable
'***************************************************************************************************************************************************
Dim strPKGNameVar 					: strPKGNameVar 						= fnReadIni(strIniConfFilePath, "Main", "strPKGNameVar")
Dim strFriendlyPKGNameVar 			: strFriendlyPKGNameVar 				= fnReadIni(strIniConfFilePath, "Main", "strFriendlyPKGNameVar")

Dim arrGUIDsVar 					: arrGUIDsVar 							= fnReadIni(strIniConfFilePath, "Main", "arrGUIDsVar")
									  If arrGUIDsVar <> "N/A" Then arrGUIDsVar = Split(arrGUIDsVar, ",")
									
Dim arrMSIVar 						: arrMSIVar 							= fnReadIni(strIniConfFilePath, "Main", "arrMSIVar")
									  If arrMSIVar <> "N/A" Then arrMSIVar = Split(arrMSIVar, ",")
									  
Dim arrMSTVar 						: arrMSTVar 							= fnReadIni(strIniConfFilePath, "Main", "arrMSTVar")
									  If arrMSTVar <> "N/A" Then arrMSTVar = Split(arrMSTVar, ",")
									  
Dim arrMSPVar 						: arrMSPVar 							= fnReadIni(strIniConfFilePath, "Main", "arrMSPVar")
									  If arrMSPVar <> "N/A" Then arrMSPVar = Split(arrMSPVar, ",")
									  
Dim arrEXEVar 						: arrEXEVar 							= fnReadIni(strIniConfFilePath, "Main", "arrEXEVar")
									  If arrEXEVar <> "N/A" Then arrEXEVar = Split(arrEXEVar, ",")
									  
Dim arrExVBSVar 					: arrExVBSVar 							= fnReadIni(strIniConfFilePath, "Main", "arrExVBSVar")
									  If arrExVBSVar <> "N/A" Then arrExVBSVar = Split(arrExVBSVar, ",")
									  
Dim arrCABsVar 						: arrCABsVar 							= fnReadIni(strIniConfFilePath, "Main", "arrCABs")
									  If arrCABsVar <> "N/A" Then arrCABsVar = Split(arrCABsVar, ",")
									  
Dim strProcessessToKillVar 			: strProcessessToKillVar 				= fnReadIni(strIniConfFilePath, "Main", "strProcessessToKillVar")
Dim strFriendlyProcessNameVar 		: strFriendlyProcessNameVar 			= fnReadIni(strIniConfFilePath, "Main", "strFriendlyProcessNameVar")
Dim strLanguagesToUseVar 			: strLanguagesToUseVar 					= fnReadIni(strIniConfFilePath, "Main", "strLanguagesToUseVar")

Dim intProcessHTAWindowTimeVar 		: intProcessHTAWindowTimeVar 			= fnReadIni(strIniConfFilePath, "Main", "intProcessHTAWindowTimeVar")
									  intProcessHTAWindowTimeVar = CInt(intProcessHTAWindowTimeVar)
									  
Dim strAppNameVar 					: strAppNameVar 						= fnReadIni(strIniConfFilePath, "Main", "strAppNameVar")
Dim strAppNumberVar 				: strAppNumberVar 						= fnReadIni(strIniConfFilePath, "Main", "strAppNumberVar")
Dim strPkgGUIDVar 					: strPkgGUIDVar 						= fnReadIni(strIniConfFilePath, "Main", "strPkgGUIDVar")
Dim strLanguageCodeVar 				: strLanguageCodeVar 					= fnReadIni(strIniConfFilePath, "Main", "strLanguageCodeVar")

Dim boolRebootNeededVar 			: boolRebootNeededVar 					= fnReadIni(strIniConfFilePath, "Main", "boolRebootNeededVar")
									  boolRebootNeededVar = Cbool(boolRebootNeededVar)
									  
Dim boolLogOffNeededVar 			: boolLogOffNeededVar 					= fnReadIni(strIniConfFilePath, "Main", "boolLogOffNeededVar")
									  boolLogOffNeededVar = Cbool(boolLogOffNeededVar)
									  
Dim intInstallationTimeVar 			: intInstallationTimeVar 				= fnReadIni(strIniConfFilePath, "Main", "intInstallationTimeVar")
									  intInstallationTimeVar = CInt(intInstallationTimeVar)
'--------------------------------------------------------------------------------------------
'! Converts path from ini property (if system variable is used) to string.
'!
'! @param  strDirectoryVar     System variable path.
'!
'! @return string
Function fnDirectoryPath (strDirectoryVar)
	Dim strDirVar, intFirst, intLast, intLengthOfArray
	
		if strDirectoryVar = "NONE" Then
			fnDirectoryPath = "NONE"
			Exit Function 
		End If

		strDirVar = Split(strDirectoryVar, "%")
		
		intFirst = LBound(strDirVar)
		intLast = UBound(strDirVar)

		intLengthOfArray = intLast - 1
		
		if strDirVar(intLast) = "" Then
			fnDirectoryPath = fnOSVariableTranslation("%" & strDirVar(intLast-1) & "%", false)
			Exit Function 
		Else 
			fnDirectoryPath = fnOSVariableTranslation("%" & strDirVar(intLast-1) & "%", false) & strDirVar(intLast)
			Exit Function 
		End If
End Function
'--------------------------------------------------------------------------------------------
'! Reads and return variable from .ini file.
'!
'! @param  strMyFilePath       Full path and file Name of the INI file.
'! @param  strMySection        The section in the INI file to be searched.
'! @param  strMyKey            The ini key whose value is to be returned.
'!
'! @return 1,string
Public Function fnReadIni (strMyFilePath, strMySection, strMyKey)
    Const intForReading   = 1
    Const intForWriting   = 2
    Const intForAppending = 8

    Dim intEqualPos, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection
	
	strFilePath = Trim(strMyFilePath)
    strSection  = Trim(strMySection)
    strKey      = Trim(strMyKey)
	
    If objFSO.FileExists(strFilePath) Then
        Set objIniFile = objFSO.OpenTextFile(strFilePath, intForReading, False)
      
		Do While objIniFile.AtEndOfStream = False
            strLine = Trim(objIniFile.ReadLine)
			
            If LCase(strLine) = "[" & LCase(strSection) & "]" Then
                strLine = Trim(objIniFile.ReadLine)

                Do While Left(strLine, 1 ) <> "["
                    intEqualPos = InStr(1, strLine, "=", 1)
                    If intEqualPos > 0 Then
                        strLeftString = Trim(Left(strLine, intEqualPos - 1))
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            fnReadIni = Trim(Mid(strLine, intEqualPos + 1))
                            If fnReadIni = "" Then
                                fnReadIni = " "
                            End If
                            Exit Do
                        End If
                    End If

                    If objIniFile.AtEndOfStream Then Exit Do

                    strLine = Trim(objIniFile.ReadLine)
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
			fnReadIni = 0
    End If
End Function
'--------------------------------------------------------------------------------------------
	