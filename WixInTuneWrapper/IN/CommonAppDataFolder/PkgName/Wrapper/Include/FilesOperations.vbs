'****************************************************************************************************************************************************
' FUNCTIONS: 
'	fnCopyFileF(strFileName, strFromDirectory, strToDirectory)
'   fnCopyFileToAllUsersProfiles(strFileName, strFromDirectory, strToDirectory)
'   fnCopyNewFile(DateTime, strFromDirectory, strToDirectory)
'	fnMoveFile(strFileName, strFromDirectory, strToDirectory) 
'	fnRemoveFile(strFileName, strVbsDirectory) 
'	fnIsExist(strFileName, strVbsDirectory) 
'	fnRenameFile(strFileName, strNewFileName, strVbsDirectory) 
'	fnReplaceStringInFile(strFileName, strVbsDirectory, strSString, strNString, boolLogAll) strSString - old string, strNString - new string
'	fnAddStringToFile(strFileName, strVbsDirectory, strSString) 
'	fnAddTXTFile(strFileName, strVbsDirectory) 
'	fnSerchFile(strFileName-full or with *, strDirectoryToStart) - for now only full Name
'   fnSearchStringLineInFile(strFileName, strVbsDirectory, strSString)
'	
' FUNCTION EXAMPLE USAGE (IN and OUT):
' 
'***************************************************************************************************************************************************
'! Copies file from one location to another. 
'!
'! @param  strFileName         File Name with extension.
'! @param  strFromDirectory    Source from where the file are to be copied to destination.
'! @param  strToDirectory      Destination where the file from source are to be copied.
'!
'! @return 0,1
Function fnCopyFileF(strFileName, strFromDirectory, strToDirectory)
	Dim intRVal, copyPath
    Dim strCurrentFNName                   : strCurrentFNName = "fnCopyFileF"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strFileName - " & strFileName & ", strFromDirectory - " & strFromDirectory & ", strToDirectory - " & strToDirectory & DotFN(0)) 
		
		If FnParamChecker(strFileName, "strFileName") = 1 Then
			call fnConst_VarFail("strFileName", strCurrentFNName) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnCopyFileF = 1
		End If
				
		If FnParamChecker(strFromDirectory, "strFromDirectory") = 1 Then
			call fnConst_VarFail("strFromDirectory", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnCopyFileF = 1
		ElseIf strFromDirectory = "\" Then
			strFromDirectory = objFSO.GetParentFolderName(WScript.ScriptFullName)
			fnAddLog(SetVarFN(0) & "strFromDirectory" & SetVarFN(1) & strFromDirectory & DotFN(0))
		End If
				
		If FnParamChecker(strToDirectory, "strToDirectory") = 1 Then
			call fnConst_VarFail("strToDirectory", strCurrentFNName) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnCopyFileF = 1
		End If

            if fnLastCharChecker(strFromDirectory, "\") = 0 Then
                strFromDirectory = Left(strFromDirectory, Len(strFromDirectory) - 1)
            End If
            
            if fnLastCharChecker(strToDirectory, "\") = 0 Then
                strToDirectory = Left(strToDirectory, Len(strToDirectory) - 1)
            End If
    
		If fnIsExistD(strFromDirectory) = 0 Then	
			fnAddLog(PreTxtFN(0) & "strFromDirectory - " & strFromDirectory & SetVarFN(2))
			If fnIsExistD(strToDirectory) = 0 Then	
				fnAddLog(PreTxtFN(0) & "strToDirectory - " & strToDirectory & SetVarFN(2))
				If fnIsExist(strFileName, strFromDirectory) = 0 Then	
					copyPath = strFromDirectory & "\" & strFileName
					fnAddLog(PreTxtFN(0) & "copyPath - " & copyPath & DotFN(0))
					intRVal = objFSO.CopyFile(copyPath, strToDirectory & "\", true)
						intRVal = fnIsExist(strFileName, strToDirectory) 
                    call Const_FnEndrVal(strCurrentFNName, intRVal)
					If intRVal = 1 Then 
						fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
						
						strLogTAB = strLogTABBak
						fnFinalize(1)
					End If
					fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0)) 
					
					strLogTAB = strLogTABBak
					fnCopyFileF = intRVal
				Else
					fnAddLog(PreTxtFN(1) & "strFileName - " & strFromDirectory & "\" & strFileName & SetVarFN(3))
					fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
					
					strLogTAB = strLogTABBak
					fnFinalize(1)
				End If
			Else
				if fnCreateDirectory(strToDirectory, true) = 0 Then
					fnAddLog(PreTxtFN(0) & " strToDirectory - " & strToDirectory & SetVarFN(4))
					If fnIsExist(strFileName, strFromDirectory)  = 0 Then		
						copyPath = strFromDirectory & "\" & strFileName
						fnAddLog(PreTxtFN(0) & " copyPath - " & copyPath & DotFN(0))
						intRVal = objFSO.CopyFile(copyPath, strToDirectory & "\", true)
							intRVal = fnIsExist(strFileName, strToDirectory) 
                        call Const_FnEndrVal(strCurrentFNName, intRVal)
						If intRVal = 1 Then 
							fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
							
							strLogTAB = strLogTABBak
							fnFinalize(1)
						End If
						fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0))

						strLogTAB = strLogTABBak					
						fnCopyFileF = intRVal
					Else
						fnAddLog(PreTxtFN(1) & "strFileName - " & strFromDirectory & "\" & strFileName & SetVarFN(3))
						fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
						
						strLogTAB = strLogTABBak
						fnFinalize(1)
					End If
				Else
					fnAddLog(PreTxtFN(1) & "strToDirectory - " & strToDirectory & SetVarFN(3))
					fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
					
					strLogTAB = strLogTABBak
					fnFinalize(1)
				End If
			End If
		Else
			fnAddLog(PreTxtFN(1) & "strFromDirectory - " & strFromDirectory & SetVarFN(3))
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Copies file from one location to all user profiles (including "Default") home directories.
'!
'! @param  strFileName			File Name with extension.
'! @param  strFromDirectory	Source from where the file are to be copied to destination.
'! @param  strToDirectory		Destination where the file from source are to be copied.
'!
'! @return intRVal
Function fnCopyFileToAllUsersProfiles(strFileName, strFromDirectory, strToDirectory)
	Const HKEY_LOCAL_MACHINE = &H80000002
	Dim strSubKey, arrSubKeys, arrProfilePath, strValue, strUserName
	Dim strKeyPath						: strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	Dim strDefault						: strDefault = strRootDrive & "\Users\Default"
	Dim intRVal							: intRVal = 0
	Dim strCurrentFNName					: strCurrentFNName = "fnCopyFileToAllUsersProfiles"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strFileName - " & strFileName & ", strFromDirectory - " & strFromDirectory & ", strToDirectory - " & strToDirectory & DotFN(0))
		
		If FnParamChecker(strFileName, "strFileName") = 1 Then
			Call fnConst_VarFail("strFileName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnCopyFileToAllUsersProfiles = 1
		End If
		
		If FnParamChecker(strFromDirectory, "strFromDirectory") = 1 Then
			Call fnConst_VarFail("strFromDirectory", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnCopyFileToAllUsersProfiles = 1
		ElseIf strFromDirectory = "\" Then
			strFromDirectory = objFSO.GetParentFolderName(WScript.ScriptFullName)
			fnAddLog(SetVarFN(0) & "strFromDirectory" & SetVarFN(1) & strFromDirectory & DotFN(0))
		End If
		
		If FnParamChecker(strToDirectory, "strToDirectory") = 1 OR strToDirectory = "\" Then
			strToDirectory = ""
			fnAddLog(SetVarFN(0) & "strToDirectory" & SetVarFN(1) & "empty string" & DotFN(0))
		End If
		
		
	'----- Getting user profile home directory path:
		objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
		
		For Each strSubKey In arrSubKeys
			objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, strKeyPath & "\" & strSubKey, "ProfileImagePath", strValue
			
			arrProfilePath = Split(strValue, "\")					'Splitting user profile path to get user Name, which will be stored in the intLast array element.
			strUserName = arrProfilePath(UBound(arrProfilePath))	'Getting user Name (intLast array element) from user profile path.
		
		
	'----- Copying file to user profile home directory (except for System and Service profiles):
			If NOT strUserName = "systemprofile" AND NOT strUserName = "LocalService" AND NOT strUserName = "NetworkService" Then
				intRVal = fnCopyFileF(strFileName, strFromDirectory, strValue & strToDirectory)
			End If
		
		Next
		
		
	'----- Copying file to "Default" user profile home directory:
		intRVal = fnCopyFileF(strFileName, strFromDirectory, strDefault & strToDirectory)
		
		
		fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0))
		
		strLogTAB = strLogTABBak
		fnCopyFileToAllUsersProfiles = intRVal
End Function
'-------------------------------------------------------------------------------------------
'! Copies file 
'!
'! @param  DateTime			.
'! @param  strFromDirectory	Source from where the file are to be copied to destination.
'! @param  strToDirectory		Destination where the file from source are to be copied.
'!
'! @return intRVal
Function fnCopyNewFile(DateTime, strFromDirectory, strToDirectory)
	Dim intRVal, copyPath
    Dim strCurrentFNName                   : strCurrentFNName = "fnCopyNewFile"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
	Dim objFolder, colFiles, objFile
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "DateTime - " & DateTime & ", strFromDirectory - " & strFromDirectory & ", strToDirectory - " & strToDirectory & DotFN(0)) 
		
		If FnParamChecker(DateTime, "DateTime") = 1 Then
			call fnConst_VarFail("DateTime", strCurrentFNName) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnCopyNewFile = 1
		End If
				
		If FnParamChecker(strFromDirectory, "strFromDirectory") = 1 Then
			call fnConst_VarFail("strFromDirectory", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnCopyNewFile = 1
		ElseIf strFromDirectory = "\" Then
			strFromDirectory = objFSO.GetParentFolderName(WScript.ScriptFullName)
			fnAddLog(SetVarFN(0) & "strFromDirectory" & SetVarFN(1) & strFromDirectory & DotFN(0))
		End If
				
		If FnParamChecker(strToDirectory, "strToDirectory") = 1 Then
			call fnConst_VarFail("strToDirectory", strCurrentFNName) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnCopyNewFile = 1
		End If

			
		Set objFolder = objFSO.GetFolder(strFromDirectory)
		Set colFiles = objFolder.Files

		For Each objFile in colFiles
			If DateDiff("s", DateTime, objFile.DateLastModified) > 0 Then
				intRVal = fnCopyFileF(objFile.Name, strFromDirectory, strToDirectory)
			End If
		Next
		
		fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0)) 
					
		strLogTAB = strLogTABBak
		fnCopyNewFile = intRVal
End Function
'-------------------------------------------------------------------------------------------
'! Moves file from one location to another. 
'!
'! @param  strFileName         File Name with extension.
'! @param  strFromDirectory    Source from where the file are to be moved to destination.
'! @param  strToDirectory      Destination where the file from source are to be moved.
'!
'! @return 0,1
Function fnMoveFile(strFileName, strFromDirectory, strToDirectory)
	Dim intRVal 
    Dim strCurrentFNName                   : strCurrentFNName = "fnMoveFile"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strFileName - " & strFileName & ", strFromDirectory - " & strFromDirectory & ", strToDirectory - " & strToDirectory & DotFN(0)) 
		
		If FnParamChecker(strFileName, "strFileName") = 1 Then
			call fnConst_VarFail("strFileName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnMoveFile = 1
		End If
				
		If FnParamChecker(strFromDirectory, "strFromDirectory") = 1 Then
			call fnConst_VarFail("strFromDirectory", strCurrentFNName)  
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnMoveFile = 1
		End If
				
		If FnParamChecker(strToDirectory, "strToDirectory") = 1 Then
			call fnConst_VarFail("strToDirectory", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnMoveFile = 1
		End If
    
            if fnLastCharChecker(strFromDirectory, "\") = 0 Then
                strFromDirectory = Left(strFromDirectory, Len(strFromDirectory) - 1)
            End If
            
            if fnLastCharChecker(strToDirectory, "\") = 1 Then
                strToDirectory = strToDirectory & "\"
            End If

		If fnIsExistD(strFromDirectory) = 0 Then	
			fnAddLog(PreTxtFN(0) & "strFromDirectory - " & strFromDirectory & SetVarFN(2))
			If fnIsExistD(strToDirectory) = 0 Then	
				fnAddLog(PreTxtFN(0) & "strToDirectory - " & strToDirectory & SetVarFN(2))
				If fnIsExist(strFileName, strFromDirectory) = 0 Then	
					intRVal = objFSO.fnMoveFile(strFromDirectory & "\" & strFileName, strToDirectory)
						intRVal = fnIsExist(strFileName, strToDirectory) 
						If intRVal = 0 Then
							intRVal = fnIsExist(strFileName, strFromDirectory) 
								if intRVal = 0 Then
									intRVal = 1
								Else
									intRVal = 0
								End If
						End If
                    call Const_FnEndrVal(strCurrentFNName, intRVal)
					If intRVal = 1 Then 
						fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))

						strLogTAB = strLogTABBak						
						fnFinalize(1)
					End If
					fnAddLog(ndingFN(0) & strCurrentFNName & SucFailFN(0)) 
					
					strLogTAB = strLogTABBak
					fnMoveFile = intRVal
				Else
					fnAddLog(PreTxtFN(1) & "strFileName - " & strFromDirectory & "\" & strFileName & SetVarFN(3))
					fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
					
					strLogTAB = strLogTABBak
					fnFinalize(1)
				End If
			Else
				if fnCreateDirectory(strToDirectory, true) = 0 Then
					fnAddLog(PreTxtFN(0) & "strToDirectory - " & strToDirectory & CreationFN(3))
					If fnIsExist(strFileName, strFromDirectory) = 0 Then		
						intRVal = objFSO.fnMoveFile(strFromDirectory & "\" & strFileName, strToDirectory)
							intRVal = fnIsExist(strFileName, strToDirectory) 
							If intRVal = 0 Then
								intRVal = fnIsExist(strFileName, strFromDirectory) 
									if intRVal = 0 Then
										intRVal = 1
									Else
										intRVal = 0
									End If
							End If
                        call Const_FnEndrVal(strCurrentFNName, intRVal)
						If intRVal = 1 Then 
							fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
							
							strLogTAB = strLogTABBak
							fnFinalize(1)
						End If
						fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0))

						strLogTAB = strLogTABBak						
						fnMoveFile = intRVal
					Else
						fnAddLog(PreTxtFN(1) & "strFileName - " & strFromDirectory & "\" & strFileName & SetVarFN(3))
						fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
						
						strLogTAB = strLogTABBak
						fnFinalize(1)
					End If
				Else
					fnAddLog(PreTxtFN(1) & "strToDirectory - " & strToDirectory & SetVarFN(3))
					fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
					
					strLogTAB = strLogTABBak
					fnFinalize(1)
				End If
			End If
		Else
			fnAddLog(PreTxtFN(1) & "strFromDirectory - " & strFromDirectory & SetVarFN(3))
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Removes file from pointed location. 
'!
'! @param  strFileName         File Name with extension.
'! @param  strVbsDirectory    	strVbsDirectory where the file is stored.
'!
'! @return 0,1
Function fnRemoveFile(strFileName, strVbsDirectory) 
	Dim intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnRemoveFile"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strFileName - " & strFileName & ", strVbsDirectory - " & strVbsDirectory & DotFN(0))
		
		If FnParamChecker(strFileName, "strFileName") = 1 Then
			call fnConst_VarFail("strFileName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRemoveFile = 1
		End If
				
		If FnParamChecker(strVbsDirectory, "strVbsDirectory") = 1 Then
			call fnConst_VarFail("strVbsDirectory", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRemoveFile = 1
		End If
    
            if fnLastCharChecker(strVbsDirectory, "\") = 0 Then
                strVbsDirectory = Left(strVbsDirectory, Len(strVbsDirectory) - 1)
            End If

		If fnIsExistD(strVbsDirectory) = 0 Then	
			fnAddLog(PreTxtFN(0) & "strVbsDirectory - " & strVbsDirectory & SetVarFN(2))
			If fnIsExist(strFileName, strVbsDirectory) = 0 Then
				fnAddLog(PreTxtFN(0) & "strFileName - " & strVbsDirectory & "\" & strFileName & SetVarFN(2))	
				intRVal = objFSO.DeleteFile(strVbsDirectory & "\" & strFileName, true)
						intRVal = fnIsExist(strFileName, strVbsDirectory) 
						if intRVal = 1 Then 
							intRVal = 0 
						Else 
							intRVal = 1
						End If
                call Const_FnEndrVal(strCurrentFNName, intRVal)
					If intRVal = 1 Then 
						fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
						
						strLogTAB = strLogTABBak
						fnFinalize(1)
					End If

                    If fnIsNull(intRVal) OR intRVal = 0 OR isEmpty(intRVal) Then
                        intRVal = 0
                    End IF
            
					fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0)) 
					
					strLogTAB = strLogTABBak
					fnRemoveFile = intRVal
			Else
				fnAddLog(PreTxtFN(0) & "strFileName - " & strVbsDirectory & "\" & strFileName & SetVarFN(3))	
				fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0)) 
				
				strLogTAB = strLogTABBak
                intRVal = 0 
				fnRemoveFile = intRVal
			End If
		Else
			fnAddLog(PreTxtFN(1) & "strVbsDirectory - " & strVbsDirectory & SetVarFN(3))
			fnAddLog(EndingFN(0) & strCurrentFNName & DotFN(2)) 
			
			strLogTAB = strLogTABBak
            fnRemoveFile = 0 'wyjscie bez faila bo skoro nie ma katalogu to pliku tez, wiec w sumie sukcess :D
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Checking if file is present in pointed location. 
'!
'! @param  strFileName         File Name with extension.
'! @param  strVbsDirectory    	strVbsDirectory where the file is stored.
'!
'! @return 0,1
Function fnIsExist(strFileName, strVbsDirectory)  
	Dim intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnIsExist"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strFileName - " & strFileName & ", strVbsDirectory - " & strVbsDirectory & DotFN(0))
		
		If FnParamChecker(strFileName, "strFileName") = 1 Then
			call fnConst_VarFail("strFileName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnIsExist = 1
		End If
				
		If FnParamChecker(strVbsDirectory, "strVbsDirectory") = 1 Then
			call fnConst_VarFail("strVbsDirectory", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnIsExist = 1
		End If
		
		If strVbsDirectory = "\" Then
			strVbsDirectory = strScriptPath
		End If
    
            if fnLastCharChecker(strVbsDirectory, "\") = 0 Then
                strVbsDirectory = Left(strVbsDirectory, Len(strVbsDirectory) - 1)
            End If

		If objFSO.FileExists(strVbsDirectory & "\" & strFileName) = true Then
			fnAddLog(PreTxtFN(0) & "strFileName - " & strVbsDirectory & "\" & strFileName & SetVarFN(2)) 
			fnAddLog(EndingFN(0) & strCurrentFNName & DotFN(0))

			strLogTAB = strLogTABBak			
			fnIsExist = 0
		Else
			fnAddLog(PreTxtFN(0) & "strFileName - " & strVbsDirectory & "\" & strFileName & SetVarFN(3)) 
			fnAddLog(EndingFN(0) & strCurrentFNName & DotFN(0)) 
			
			strLogTAB = strLogTABBak
			fnIsExist = 1
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Checking if file is present in pointed location. 
'!
'! @param  strFileName         Original file Name with extension.
'! @param  strNewFileName      New file Name with extension.
'! @param  strVbsDirectory    	strVbsDirectory where the file is stored.
'!
'! @return 0,1
Function fnRenameFile(strFileName, strNewFileName, strVbsDirectory) 
	Dim intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnRenameFile"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strFileName - " & strFileName & ", strVbsDirectory - " & strVbsDirectory & ", strNewFileName - " & strNewFileName & DotFN(0))
		
		If FnParamChecker(strFileName, "strFileName") = 1 Then
			call fnConst_VarFail("strFileName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRenameFile = 1
		End If
		
		If FnParamChecker(strNewFileName, "strNewFileName") = 1 Then
			call fnConst_VarFail("strNewFileName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRenameFile = 1
		End If
				
		If FnParamChecker(strVbsDirectory, "strVbsDirectory") = 1 Then
			call fnConst_VarFail("strVbsDirectory", strCurrentFNName) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRenameFile = 1
		End If
		
		If strVbsDirectory = "\" Then
			strVbsDirectory = strScriptPath
		End If
    
            if fnLastCharChecker(strVbsDirectory, "\") = 0 Then
                strVbsDirectory = Left(strVbsDirectory, Len(strVbsDirectory) - 1)
            End If

		If fnIsExistD(strVbsDirectory) = 0 Then
			If (objFSO.FileExists(strVbsDirectory & "\" & strFileName)) Then
				fnAddLog(PreTxtFN(0) & strVbsDirectory & "\" & strFileName & CreationFN(6))
					If (objFSO.FileExists(strVbsDirectory & "\" & strNewFileName)) Then
						fnAddLog(PreTxtFN(0) & CreationFN(8) & strVbsDirectory & "\" & strNewFileName & " to " & strVbsDirectory & "\" & strNewFileName & "_backup.") 
						objFSO.fnMoveFile strVbsDirectory & "\" & strNewFileName, strVbsDirectory & "\" & strNewFileName & "_backup"
					End If
				fnAddLog(PreTxtFN(0) & "renaming " & strVbsDirectory & "\" & strFileName & " to " & strVbsDirectory & "\" & strNewFileName & DotFN(0)) 
				objFSO.fnMoveFile strVbsDirectory & "\" & strFileName, strVbsDirectory & "\" & strNewFileName
			Else
				fnAddLog(PreTxtFN(0) & "strFileName - " & strVbsDirectory & "\" & strFileName & SetVarFN(3)) 
				fnAddLog(EndingFN(2) & strCurrentFNName & DotFN(0)) 
				
				strLogTAB = strLogTABBak
				fnRenameFile = 1
				Exit Function
			End If
		Else
			fnAddLog(PreTxtFN(2) & "strVbsDirectory: " & strVbsDirectory & SetVarFN(3))
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRenameFile = 1
		End If
		
		fnAddLog(SelfCheckFN)
		
		If (objFSO.FileExists(strVbsDirectory & "\" & strNewFileName)) Then
            fnAddLog(PreTxtFN(0) & " strFileName - " & strVbsDirectory & "\" & strFileName & SetVarFN(2)) 
            fnAddLog(EndingFN(0) & strCurrentFNName & DotFN(0))

			strLogTAB = strLogTABBak			
			fnRenameFile = 0
			Exit Function
		Else
            fnAddLog(PreTxtFN(0) & " strFileName - " & strVbsDirectory & "\" & strFileName & SetVarFN(3)) 
            fnAddLog(EndingFN(0) & strCurrentFNName & DotFN(0)) 
			
			strLogTAB = strLogTABBak
			fnRenameFile = 1
			Exit Function
		End If
		
		
		strLogTAB = strLogTABBak
End Function
'-------------------------------------------------------------------------------------------
'! Find and replace specified string in text file. 
'!
'! @param  strFileName         File Name with extension.
'! @param  strSString      	Existing string in text file to replace by new string.
'! @param  strNString      	New string to replace existing old string in file.
'! @param  strVbsDirectory    	strVbsDirectory where the file is stored.
'! @param  boolLogAll    	If TRUE text file will be added into vbs log.
'!
'! @return 0,1
Function fnReplaceStringInFile(strFileName, strVbsDirectory, strSString, strNString, boolLogAll) 
	Const intForReading = 1, intForWriting = 2, intForAppending = 8
	Dim MyFile, strAllText
    Dim strCurrentFNName                   : strCurrentFNName = "fnReplaceStringInFile"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strFileName - " & strFileName & ", strVbsDirectory - " & strVbsDirectory & ", strSString - " & strSString & ", strNString - " & strNString & DotFN(0))
		
		If FnParamChecker(strFileName, "strFileName") = 1 Then
			call fnConst_VarFail("strFileName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnReplaceStringInFile = 1
		End If
				
		If FnParamChecker(strVbsDirectory, "strVbsDirectory") = 1 Then
			call fnConst_VarFail("strVbsDirectory", strCurrentFNName) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnReplaceStringInFile = 1
		End If
		
		If FnParamChecker(strSString, "strSString") = 1 Then
			call fnConst_VarFail("strSString", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnReplaceStringInFile = 1
		End If
		
		If FnParamChecker(strNString, "strNString") = 1 Then
			call fnConst_VarFail("strNString", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnReplaceStringInFile = 1
		End If
		
		If FnParamChecker(boolLogAll, "boolLogAll") = 1 Then
			call fnConst_VarFail("boolLogAll", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnReplaceStringInFile = 1
		End If
		
		If strVbsDirectory = "\" Then
			strVbsDirectory = strScriptPath
		End If
    
            if fnLastCharChecker(strVbsDirectory, "\") = 0 Then
                strVbsDirectory = Left(strVbsDirectory, Len(strVbsDirectory) - 1)
            End If

		If fnIsExistD(strVbsDirectory) = 0 Then
			If fnIsExist(strFileName, strVbsDirectory) = 0 Then
				Set MyFile = objFSO.OpenTextFile(strVbsDirectory & "\" & strFileName, intForReading)

				strText = MyFile.ReadAll
				MyFile.Close
				strAllText = Replace(strText, strSString, strNString)

				Set MyFile = objFSO.OpenTextFile(strVbsDirectory & "\" & strFileName, intForWriting)
				MyFile.WriteLine strAllText
				MyFile.Close
            
                    if boolLogAll = true Then
                        fnAddLog(strAllText) 
                    End If
            
			        fnAddLog(PreTxtFN(0) & " strSString - " & strSString & EditFN(0) & "strNString - " & strNString & EditFN(1) & strVbsDirectory & "\" & strFileName & DotFN(0)) 
			        fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0)) 
				
				strLogTAB = strLogTABBak
				fnReplaceStringInFile = 0
				Exit Function
			Else
			    fnAddLog(PreTxtFN(0) & " File - " & strVbsDirectory & "\" & strFileName & SetVarFN(3)) 
			    fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
				
				strLogTAB = strLogTABBak
				fnFinalize(1)
				fnReplaceStringInFile = 1
			End If
		Else
			fnAddLog(PreTxtFN(0) & " strVbsDirectory - " & strVbsDirectory & SetVarFN(3)) 
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnReplaceStringInFile = 1
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Write string in any text file. 
'!
'! @param  strFileName         File Name with extension.
'! @param  strSString      	String to write in file.
'! @param  strVbsDirectory    	strVbsDirectory where the file is stored.
'!
'! @return 0,1
Function fnAddStringToFile(strFileName, strVbsDirectory, strSString) 	
	Const intForReading = 1, intForWriting = 2, intForAppending = 8
	Dim intRVal, MyFile
    Dim strCurrentFNName                   : strCurrentFNName = "fnAddStringToFile"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strFileName - " & strFileName & ", strVbsDirectory - " & strVbsDirectory & ", strSString - " & strSString & DotFN(0))
		
		If FnParamChecker(strFileName, "strFileName") = 1 Then
			call fnConst_VarFail("strFileName", strCurrentFNName) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddStringToFile = 1
		End If
				
		If FnParamChecker(strVbsDirectory, "strVbsDirectory") = 1 Then
			call fnConst_VarFail("strVbsDirectory", strCurrentFNName) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddStringToFile = 1
		End If
		
		If FnParamChecker(strSString, "strSString") = 1 Then
			call fnConst_VarFail("strSString", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddStringToFile = 1
		End If
		
		If strVbsDirectory = "\" Then
			strVbsDirectory = strScriptPath
		End If
    
            if fnLastCharChecker(strVbsDirectory, "\") = 0 Then
                strVbsDirectory = Left(strVbsDirectory, Len(strVbsDirectory) - 1)
            End If

		If fnIsExistD(strVbsDirectory) = 0 Then
			If fnIsExist(strFileName, strVbsDirectory) = 0 Then
				Set MyFile = objFSO.OpenTextFile(strVbsDirectory & "\" & strFileName, intForAppending, TristateUseDefault)
				MyFile.WriteLine(strSString)
				MyFile.Close
					fnAddLog(PreTxtFN(0) & "strSString - " & strSString & " added to file - " &  strVbsDirectory  & "\" &  strFileName & DotFN(0)) 
					fnAddLog(EndingFN(0) & strCurrentFNName & DotFN(0)) 
					
					strLogTAB = strLogTABBak
					fnAddStringToFile = 0
				Exit Function
			Else
				fnAddLog(reTxtFN(0) & "File - " & strVbsDirectory  & "\" &  strFileName & SetVarFN(3)) 
				fnAddLog(PreTxtFN(0) & CreationFN(4) & CreationFN(5) & strCurrentFNName & DotFN(2)) 
				
				If fnAddTXTFile(strFileName, strVbsDirectory) = 0 Then
					Set MyFile = objFSO.OpenTextFile(strVbsDirectory & "\" & strFileName, intForAppending, TristateUseDefault)
					MyFile.WriteLine(strSString)
					MyFile.Close
						fnAddLog(PreTxtFN(0) & "strSString - " & strSString & EditFN(2) & strVbsDirectory & "\" & strFileName & DotFN(0)) 
						fnAddLog(EndingFN(0) & strCurrentFNName & DotFN(0)) 
						
						strLogTAB = strLogTABBak
						fnAddStringToFile = 0
					Exit Function
				Else
			        fnAddLog(PreTxtFN(0) & " strVbsDirectory - " & strVbsDirectory  & "\" &  strFileName & SetVarFN(3)) 
			        fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 

					strLogTAB = strLogTABBak					
					fnFinalize(1)
					fnAddStringToFile = 1
				End If
			End If
		Else
			fnAddLog(PreTxtFN(0) & " strVbsDirectory - " & strVbsDirectory & SetVarFN(3)) 
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddStringToFile = 1
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Create a new text file. 
'!
'! @param  strFileName         File Name with extension.
'! @param  strVbsDirectory    	strVbsDirectory where the file will be stored.
'!
'! @return 0,1
Function fnAddTXTFile(strFileName, strVbsDirectory) 
	Dim intRVal, MyFile
    Dim strCurrentFNName                   : strCurrentFNName = "fnAddTXTFile"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName& StartingFN(1) & "strFileName - " & strFileName & ", strVbsDirectory - " & strVbsDirectory & DotFN(0))
		
		If FnParamChecker(strFileName, "strFileName") = 1 Then
			call fnConst_VarFail("strFileName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnSerchFile = 1
		End If
				
		If FnParamChecker(strVbsDirectory, "strVbsDirectory") = 1 Then
			call fnConst_VarFail("strVbsDirectory", strCurrentFNName)  
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnSerchFile = 1
		End If
		
		If strVbsDirectory = "\" Then
			strVbsDirectory = strScriptPath
		End If
    
            if fnLastCharChecker(strVbsDirectory, "\") = 0 Then
                strVbsDirectory = Left(strVbsDirectory, Len(strVbsDirectory) - 1)
            End If

		if fnIsExistD(strVbsDirectory) = 0 Then
			if fnIsExist(strFileName, strVbsDirectory) = 1 Then
				Set MyFile = objFSO.CreateTextFile(strVbsDirectory & "\" & strFileName, True)
				MyFile.WriteLine("")
				MyFile.Close
				
				if fnIsExist(strFileName, strVbsDirectory) = 0 Then
					fnAddLog(PreTxtFN(0) & " strFileName - " & strVbsDirectory & "\" & strFileName & CreationFN(0)) 
			        fnAddLog(EndingFN(0) & strCurrentFNName& DotFN(0)) 
					
					strLogTAB = strLogTABBak
					fnAddTXTFile = 0
					Exit Function
				Else
					fnAddLog(PreTxtFN(0) & " strFileName - " & strVbsDirectory & "\" & strFileName & CreationFN(1)) 
			        fnAddLog(EndingFN(0) & strCurrentFNName& SucFailFN(1))
					
					strLogTAB = strLogTABBak
					fnFinalize(1)
					fnAddTXTFile = 1	
				End if
			Else
				fnAddLog(PreTxtFN(0) & " strFileName - " & strVbsDirectory & "\" & strFileName & CreationFN(2))
			    fnAddLog(EndingFN(0) & strCurrentFNName& DotFN(0)) 
				
				strLogTAB = strLogTABBak
				fnAddTXTFile = 0
				Exit Function
			End If
		Else
			fnAddLog(PreTxtFN(0) & " strVbsDirectory - " & strVbsDirectory & SetVarFN(3)) 
			fnAddLog(EndingFN(0) & strCurrentFNName& SucFailFN(1)) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddTXTFile = 1	
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Search for file or files. 
'!
'! @param  strFileName         File Name.
'! @param  strVbsDirectory    	strVbsDirectory where serching will start.
'!
'! @return 0,1
Function fnSerchFile(strFileName, strDirectoryToStart)
	Dim intRVal, folder, file
    Dim strCurrentFNName                   : strCurrentFNName = "fnSerchFile"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strFileName - " & strFileName & ", strDirectoryToStart - " & strDirectoryToStart & DotFN(0))
		
		If FnParamChecker(strFileName, "strFileName") = 1 Then
			call fnConst_VarFail("strFileName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnSerchFile = 1
		End If
				
		If FnParamChecker(strDirectoryToStart, "strDirectoryToStart") = 1 Then
			call fnConst_VarFail("strDirectoryToStart", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnSerchFile = 1
		End If
		
		If strDirectoryToStart = "\" Then
			strDirectoryToStart = strScriptPath
		End If
		
		if fnIsExistD(strDirectoryToStart) = 0 Then
			Set folder = objFSO.GetFolder(strDirectoryToStart)  

			' Loop over all files in the folder until the strFileName is found
			For each file In folder.Files    
				If instr(file.Name, strFileName) = 1 Then
					fnAddLog(PreTxtFN(0) & " file: " & file.Name & SetVarFN(5)) 
					
					strLogTAB = strLogTABBak
					fnSerchFile = 0
					Exit Function
				End If
			Next
			
			fnAddLog(PreTxtFN(0) & " file: " & file.Name & SetVarFN(3)) 
			
			strLogTAB = strLogTABBak
			fnSerchFile = 1
			Exit Function
		Else
			fnAddLog(PreTxtFN(0) & " strDirectoryToStart: " & strDirectoryToStart & SetVarFN(3)) 
			
			strLogTAB = strLogTABBak
			fnSerchFile = 1
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Search for a string in any text file. The document is being searched line by line, every line is compared to the string being searched for.
'!
'! @param  strFileName			File Name with extension.
'! @param  strVbsDirectory		strVbsDirectory where the file is stored.
'! @param  strSString			String being searched for in file.
'!
'! @return 0,1
Function fnSearchStringLineInFile(strFileName, strVbsDirectory, strSString)
	Const intForReading = 1
	Dim MyFile, strText
	Dim intRVal							: intRVal = 1
	Dim strCurrentFNName					: strCurrentFNName = "fnSearchStringLineInFile"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strFileName - " & strFileName & ", strVbsDirectory - " & strVbsDirectory & ", strSString - " & strSString & DotFN(0))
		
		If FnParamChecker(strFileName, "strFileName") = 1 Then
			Call fnConst_VarFail("strFileName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnSearchStringLineInFile = 1
		End If
		
		If FnParamChecker(strVbsDirectory, "strVbsDirectory") = 1 Then
			Call fnConst_VarFail("strVbsDirectory", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnSearchStringLineInFile = 1
		End If
		
		If FnParamChecker(strSString, "strSString") = 1 Then
			Call fnConst_VarFail("strSString", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnSearchStringLineInFile = 1
		End If
		
		
		If strVbsDirectory = "\" Then
			strVbsDirectory = strScriptPath
		End If
		
			If fnLastCharChecker(strVbsDirectory, "\") = 0 Then
				strVbsDirectory = Left(strVbsDirectory, Len(strVbsDirectory) - 1)
			End If
		
		
		If fnIsExistD(strVbsDirectory) = 0 Then
			If fnIsExist(strFileName, strVbsDirectory) = 0 Then
				Set MyFile = objFSO.OpenTextFile(strVbsDirectory & "\" & strFileName, intForReading)	'Opening file to read its content.
				
				Do Until MyFile.AtEndOfStream
					strText = MyFile.ReadLine		'Reading file content line by line.
					
				'Checking, if line exists in file:
					If strText = strSString Then
						fnAddLog(PreTxtFN(0) & "File - " & strVbsDirectory & "\" & strFileName & " contains line with strSString - " & strSString & DotFN(1))
						intRVal = 0
						Exit Do		'There's no point to keep searching the file, when specific line has been found.
					End If
					
				Loop
				
				MyFile.Close
				Set MyFile = Nothing
				
				If intRVal = 1 Then
					fnAddLog(PreTxtFN(0) & "File - " & strVbsDirectory & "\" & strFileName & " doesn't contain line with strSString - " & strSString & DotFN(1))
				End If
				
			Else
				fnAddLog(PreTxtFN(1) & "File - " & strVbsDirectory & "\" & strFileName & SetVarFN(3))
				fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
				
				strLogTAB = strLogTABBak
				fnFinalize(1)
				fnSearchStringLineInFile = 1
			End If
		Else
			fnAddLog(PreTxtFN(1) & "strVbsDirectory - " & strVbsDirectory & "\" & strFileName & SetVarFN(3))
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnSearchStringLineInFile = 1
		End If
		
		
		fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0))
		
		strLogTAB = strLogTABBak
		fnSearchStringLineInFile = intRVal
End Function
'-------------------------------------------------------------------------------------------