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
Dim AccountNameVar 					: AccountNameVar 						= fnReadIni(strIniFilePath, "Main", "AccountNameVar")
	
Dim HowManyStrikesVar				: HowManyStrikesVar 					= fnReadIni(strIniFilePath, "Main", "HowManyStrikesVar")
									  HowManyStrikesVar = CInt(HowManyStrikesVar)
	
Dim TitleVar						: TitleVar 								= fnReadIni(strIniFilePath, "Main", "TitleVar")
Dim	intTimeVar							: intTimeVar 								= fnReadIni(strIniFilePath, "Main", "intTimeVar")
									  intTimeVar = CInt(intTimeVar)
	
Dim WhatInstallsVar					: WhatInstallsVar 						= fnReadIni(strIniFilePath, "Main", "WhatInstallsVar")
Dim IsForceKillingVar				: IsForceKillingVar 					= fnReadIni(strIniFilePath, "Main", "IsForceKillingVar")
									  IsForceKillingVar = Cbool(IsForceKillingVar)
	
Dim IsKillingWindow					: IsKillingWindow 						= fnReadIni(strIniFilePath, "Main", "IsKillingWindow")
									  IsKillingWindow = Cbool(IsKillingWindow)
	
Dim IsUninstallKillingWindow		: IsUninstallKillingWindow 				= fnReadIni(strIniFilePath, "Main", "IsUninstallKillingWindow")
									  IsUninstallKillingWindow = Cbool(IsUninstallKillingWindow)
	
Dim IsLangWindow					: IsLangWindow 							= fnReadIni(strIniFilePath, "Main", "IsLangWindow")
									  IsLangWindow = Cbool(IsLangWindow)

Dim strProcessesToCloseVar				: strProcessesToCloseVar 					= fnReadIni(strIniFilePath, "Main", "strProcessesToCloseVar")
Dim IsRestartMessageVar				: IsRestartMessageVar 					= fnReadIni(strIniFilePath, "Main", "IsRestartMessageVar")
									  IsRestartMessageVar = Cbool(IsRestartMessageVar)
	
Dim IsFailFinalizeMsg				: IsFailFinalizeMsg 					= fnReadIni(strIniFilePath, "Main", "IsFailFinalizeMsg")
									  IsFailFinalizeMsg = Cbool(IsFailFinalizeMsg)
	
Dim no3010							: no3010 								= fnReadIni(strIniFilePath, "Main", "no3010")
									  no3010 = Cbool(no3010)

Dim PBMode 		                    : PBMode 			                    = fnReadIni(strIniFilePath, "Main", "PBMode")
Dim InstallProgressBarTitle 		: InstallProgressBarTitle 			    = fnReadIni(strIniFilePath, "Main", "InstallProgressBarTitle")
Dim UninstallProgressBarTitle 	    : UninstallProgressBarTitle  			= fnReadIni(strIniFilePath, "Main", "UninstallProgressBarTitle")

Dim InstallProgressBarCaption 		: InstallProgressBarCaption 			= fnReadIni(strIniFilePath, "Main", "InstallProgressBarCaption")
Dim UninstallProgressBarCaption 	: UninstallProgressBarCaption  			= fnReadIni(strIniFilePath, "Main", "UninstallProgressBarCaption")
Dim ProgressBarActive 				: ProgressBarActive  					= fnReadIni(strIniFilePath, "Main", "ProgressBarActive")
									  ProgressBarActive = Cbool(ProgressBarActive)

Dim UninstallProgressBarActive 		: UninstallProgressBarActive  			= fnReadIni(strIniFilePath, "Main", "UninstallProgressBarActive")
									  UninstallProgressBarActive = Cbool(UninstallProgressBarActive)
	
Dim PrecheckIsInstalled 			: PrecheckIsInstalled  					= fnReadIni(strIniFilePath, "Main", "PrecheckIsInstalled")
									  PrecheckIsInstalled = Cbool(PrecheckIsInstalled)

Dim StrikeWindowIsPersistent		: StrikeWindowIsPersistent				= fnReadIni(strIniFilePath, "Main", "StrikeWindowIsPersistent")									  
									  StrikeWindowIsPersistent = Cbool(StrikeWindowIsPersistent)
		
Dim StrikeWindowActive 				: StrikeWindowActive  					= fnReadIni(strIniFilePath, "Main", "StrikeWindowActive")
									  StrikeWindowActive = Cbool(StrikeWindowActive)
	
Dim StrikeWindowActiveUninstall		: StrikeWindowActiveUninstall 			= fnReadIni(strIniFilePath, "Main", "StrikeWindowActiveUninstall")
									  StrikeWindowActiveUninstall = Cbool(StrikeWindowActiveUninstall)
									  
Dim blnFthStrikeActive				: blnFthStrikeActive					= fnReadIni(strIniFilePath, "Main", "blnFthStrikeActive")
									  blnFthStrikeActive = Cbool(blnFthStrikeActive)
									  
Dim selfCopy						: selfCopy								= fnReadIni(strIniFilePath, "Main", "selfCopy")
									  selfCopy = Cbool(selfCopy)					  
	
Dim OneFileSetup 					: OneFileSetup  						= fnReadIni(strIniFilePath, "Main", "OneFileSetup")
									  OneFileSetup = Cbool(OneFileSetup)
	
Dim TAGType							: TAGType  								= fnReadIni(strIniFilePath, "Main", "TAGType")
Dim strLOGType							: strLOGType  								= fnReadIni(strIniFilePath, "Main", "strLOGType")
Dim SummaryWindow					: SummaryWindow  						= fnReadIni(strIniFilePath, "Main", "SummaryWindow")
									  SummaryWindow = Cbool(SummaryWindow)

Dim PreLogDir						: PreLogDir  							= fnReadIni(strIniFilePath, "Main", "PreLogDir")
									  PreLogDir = fnDirectoryPath(PreLogDir)
	
Dim PreTAGDir						: PreTAGDir  							= fnReadIni(strIniFilePath, "Main", "PreTAGDir")
									  PreTAGDir = fnDirectoryPath(PreTAGDir)

Dim forCitrix 					    : forCitrix  						    = fnReadIni(strIniFilePath, "Main", "forCitrix")
									  forCitrix = Cbool(forCitrix)

Dim TagBitness 					    : TagBitness  						    = fnReadIni(strIniFilePath, "Main", "TagBitness")
									  TagBitness = Cbool(TagBitness)
									  
Dim EntryRequestURL 		        : EntryRequestURL 			            = fnReadIni(strIniFilePath, "Main", "EntryRequestURL")
			



Dim G_MessageInstallRequired1 		: G_MessageInstallRequired1 			= fnReadIni(strIniFilePath, "Main", "G_MessageInstallRequired1")
Dim G_MessageInstallRequired2 		: G_MessageInstallRequired2 			= fnReadIni(strIniFilePath, "Main", "G_MessageInstallRequired2")
Dim G_MessageInstallRequired3 		: G_MessageInstallRequired3 			= fnReadIni(strIniFilePath, "Main", "G_MessageInstallRequired3")
Dim G_MessageInstallRequired4 		: G_MessageInstallRequired4 			= fnReadIni(strIniFilePath, "Main", "G_MessageInstallRequired4")
Dim G_MessageInstallRequired5 		: G_MessageInstallRequired5 			= fnReadIni(strIniFilePath, "Main", "G_MessageInstallRequired5")
Dim G_MessageInstallRequired6 		: G_MessageInstallRequired6 			= fnReadIni(strIniFilePath, "Main", "G_MessageInstallRequired6")
Dim G_MessageInstallRequired7 		: G_MessageInstallRequired7 			= fnReadIni(strIniFilePath, "Main", "G_MessageInstallRequired7")

Dim G_MessageInstallOptional1 		: G_MessageInstallOptional1 			= fnReadIni(strIniFilePath, "Main", "G_MessageInstallOptional1")
Dim G_MessageInstallOptional2 		: G_MessageInstallOptional2 			= fnReadIni(strIniFilePath, "Main", "G_MessageInstallOptional2")
Dim G_MessageInstallOptional3 		: G_MessageInstallOptional3 			= fnReadIni(strIniFilePath, "Main", "G_MessageInstallOptional3")
Dim G_MessageInstallOptional4 		: G_MessageInstallOptional4 			= fnReadIni(strIniFilePath, "Main", "G_MessageInstallOptional4")
Dim G_MessageInstallOptional5 		: G_MessageInstallOptional5 			= fnReadIni(strIniFilePath, "Main", "G_MessageInstallOptional5")
Dim G_MessageInstallOptional6 		: G_MessageInstallOptional6 			= fnReadIni(strIniFilePath, "Main", "G_MessageInstallOptional6")
Dim G_MessageInstallOptional7 		: G_MessageInstallOptional7 			= fnReadIni(strIniFilePath, "Main", "G_MessageInstallOptional7")
Dim G_MessageInstallOptional8 		: G_MessageInstallOptional8 			= fnReadIni(strIniFilePath, "Main", "G_MessageInstallOptional8")

Dim G_MessageInstallIfRebootRequired : G_MessageInstallIfRebootRequired 	= fnReadIni(strIniFilePath, "Main", "G_MessageInstallIfRebootRequired")
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
			fnFinalize(1)
			fnReadIni = 1
    End If
End Function
'--------------------------------------------------------------------------------------------