'****************************************************************************************************************************************************
' FUNCTIONS: 
'	fnCopyDirectory(strFromDirectory, strToDirectory) 
'	fnCreateDirectory(strPathToCreate, boolIsLog)  
'	fnMoveDirectory(strFromDirectory, strToDirectory) 
'	fnRemoveDirectory(strVbsDirectory, boolOnlyIfEmpty)		boolOnlyIfEmpty:
'											 	- True -> this will remove directory only when it is empty
'												- False -> this will remove all folders and files inside directory and the directory itself
'	fnIsExistD(strVbsDirectory) 
'	fnRenameDirectory(strDirName, strNewDirName) 
'	fnCurrentDir()
'	fnRootDrive()
' 	fnSubstituteSelf()
'	fnCopyFoldersToAllUsersProfiles(strFromDirectory, strToDirectory)
'   fnRemoveDirectoryFromAllUsersProfiles(strDirectoryV, boolOnlyIfEmpty)
'	
' FUNCTION EXAMPLE USAGE (IN and OUT):
' 
' fnRemoveDirectory(strVbsDirectory, boolOnlyIfEmpty)
'	Return: 0 - success (directory removed), 1 - fail (directory not removed), -1 - success/fail (directory not removed due to it's not empty)
'***************************************************************************************************************************************************
'! Check if objFolder exist. 
'!
'! @param  strVbsDirectory        strVbsDirectory strPath.
'!
'! @return 0,1
Function fnIsExistD(strDirectoryV) 
    Dim strVbsDirectory
    Dim strCurrentFNName                   : strCurrentFNName = "fnIsExistD"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strDirectoryV - " & strDirectoryV & DotFN(0)) 
		
		If FnParamChecker(strDirectoryV, "strDirectoryV") = 1 Then
			call fnConst_VarFail("strDirectoryV", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnIsExistD = 1
        Else
                strVbsDirectory = strDirectoryV
		End If
    
            if fnLastCharChecker(strVbsDirectory, "\") = 1 Then
                strVbsDirectory = strVbsDirectory & "\"
            End IF
		
		If objFSO.FolderExists(strVbsDirectory) = true Then
			fnAddLog(PreTxtFN(0) & "strVbsDirectory - " & strVbsDirectory & CreationFN(6)) 
			fnAddLog(EndingFN(0) & strCurrentFNName & DotFN(0)) 
			
			strLogTAB = strLogTABBak
			fnIsExistD = 0
			Exit Function
		Else
			fnAddLog(PreTxtFN(0) & "strVbsDirectory - " & strVbsDirectory & SetVarFN(3)) 
			fnAddLog(EndingFN(0) & strCurrentFNName & DotFN(0)) 
			
			strLogTAB = strLogTABBak
			fnIsExistD = 1
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Copy objFolder from one directory to another. 
'!
'! @param  strFromDirectory    strVbsDirectory to copy (full strPath).
'! @param  strToDirectory      Target directory strPath.
'!
'! @return 0,1
Function fnCopyDirectory(strFromDirectory, strToDirectory) 
	Dim intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnCopyDirectory"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strFromDirectory - " & strFromDirectory & "strToDirectory - " & strToDirectory & DotFN(0)) 
		
		If FnParamChecker(strFromDirectory, "strFromDirectory") = 1 Then
			call fnConst_VarFail("strFromDirectory", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnCopyDirectory = 1
		End If
		
		If FnParamChecker(strToDirectory, "strToDirectory") = 1 Then
			call fnConst_VarFail("strToDirectory", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnCopyDirectory = 1
		End If
    
            if fnLastCharChecker(strFromDirectory, "\") = 0 Then
                strFromDirectory = Left(strFromDirectory, Len(strFromDirectory) - 1)
            End If
            
            if fnLastCharChecker(strToDirectory, "\") = 1 Then
                strToDirectory = strToDirectory & "\"
            End If
    
		If fnIsExistD(strFromDirectory) = 0 Then	
			If fnIsExistD(strToDirectory) = 0 Then	
				intRVal = objFSO.CopyFolder(strFromDirectory, strToDirectory, true)
                call Const_FnEndrVal(strCurrentFNName, intRVal)
            
				If intRVal = 1 Then 
					fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
					
					strLogTAB = strLogTABBak
					fnFinalize(1)
				End If
				
				fnAddLog(SelfCheckFN)
				if fnIsExistD(strToDirectory) <> 0 Then
					fnAddLog(SelfCheckFN) 
					fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
					
					strLogTAB = strLogTABBak
                    intRVal = 1
					fnFinalize(intRVal)
				End If
			
                If fnIsNull(intRVal) OR intRVal = 0 Then
                    intRVal = 0
                End IF
            
                fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0)) 
				
				strLogTAB = strLogTABBak
				fnCopyDirectory = intRVal
				Exit Function
			Else
				fnAddLog(PreTxtFN(1) & "strToDirectory - " & strToDirectory & SetVarFN(3)) 
				fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
				
				strLogTAB = strLogTABBak
				fnFinalize(1)
			End If
		Else
            fnAddLog(PreTxtFN(1) & "strFromDirectory - " & strFromDirectory & SetVarFN(3))
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))

			strLogTAB = strLogTABBak		
			fnFinalize(1)
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Create all missing folders in pointed strPath. 
'!
'! @param  strPathToCreate     Full strPath to create.
'! @param  boolIsLog            If - log actions from this function.
'!
'! @return 0,1
Function fnCreateDirectory(strPathToCreate, boolIsLog)  
	Dim intRVal, arrA, strPath
    Dim strCurrentFNName                   : strCurrentFNName = "fnCreateDirectory"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		If fnIsNull(boolIsLog) = True or boolIsLog = "" or isEmpty(boolIsLog) = True Then
			boolIsLog = true
		End If
		
		if boolIsLog = true Then
		    fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strPathToCreate - " & strPathToCreate & DotFN(0)) 
		End If
		
		If FnParamChecker(strPathToCreate, "strPathToCreate") = 1 Then
			if boolIsLog = true Then
			     call fnConst_VarFail("strPathToCreate", strCurrentFNName)
			End If
			
			strLogTAB = strLogTABBak		
			fnFinalize(1)
			fnCreateDirectory = 1
		End If
    
    
            if fnLastCharChecker(strPathToCreate, "\") = 1 Then
                strPathToCreate = strPathToCreate & "\"
            End If
        
		
		arrA = split(strPathToCreate, "\")
		strPath = ""
			For Each dir In arrA
                If dir <> "" Then
				    If strPath <> "" Then strPath = strPath & "\"
				    strPath = strPath & dir
				    If objFSO.FolderExists(strPath) = False Then objFSO.CreateFolder(strPath)
                End If
			Next
		
		
		If objFSO.FolderExists(strPathToCreate) = True Then
			if boolIsLog = true Then
                fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0))  
			End If
			
			strLogTAB = strLogTABBak
			fnCreateDirectory = 0
			Exit Function
		Else
			if boolIsLog = true Then
				fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
			End If

			strLogTAB = strLogTABBak		
			fnFinalize(1)
			fnCreateDirectory = 1
		End If			
End Function
'-------------------------------------------------------------------------------------------
'! Move objFolder from one directory to another. 
'!
'! @param  strFromDirectory    strVbsDirectory to move (full strPath).
'! @param  strToDirectory      Target directory strPath.
'!
'! @return 0,1
Function fnMoveDirectory(strFromDirectory, strToDirectory) 
	Dim intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnMoveDirectory"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strFromDirectory - " & strFromDirectory & "strToDirectory - " & strToDirectory & DotFN(0)) 
		
		If FnParamChecker(strFromDirectory, "strFromDirectory") = 1 Then
            call fnConst_VarFail("strFromDirectory", strCurrentFNName) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnMoveDirectory = 1
		End If
		
		If FnParamChecker(strToDirectory, "strToDirectory") = 1 Then
            call fnConst_VarFail("strToDirectory", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnMoveDirectory = 1
		End If
            
            if fnLastCharChecker(strFromDirectory, "\") = 0 Then
                strFromDirectory = Left(strFromDirectory, Len(strFromDirectory) - 1)
            End If
            
            if fnLastCharChecker(strToDirectory, "\") = 1 Then
                strToDirectory = strToDirectory & "\"
            End If
            
		If fnIsExistD(strFromDirectory) = 0 Then	
			If fnIsExistD(strToDirectory) = 0 Then	
				intRVal = objFSO.MoveFolder(strFromDirectory, strToDirectory)
                call Const_FnEndrVal(strCurrentFNName, intRVal)
				If intRVal = 1 Then 
					fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
					
					strLogTAB = strLogTABBak
					fnFinalize(1)
				End If
				
				
				
				fnAddLog(SelfCheckFN) 'added, check if working properly
				if fnIsExistD(strToDirectory) <> 0 Then
					fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))  
					
					strLogTAB = strLogTABBak
					fnFinalize(1)
				End If
				
				
						
                fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0)) 
				
				strLogTAB = strLogTABBak
				fnMoveDirectory = intRVal
				Exit Function
			Else
                fnAddLog(PreTxtFN(1) & "strToDirectory - " & strToDirectory & SetVarFN(3))
				fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
				
				strLogTAB = strLogTABBak
				fnFinalize(1)
			End If
		Else
            fnAddLog(PreTxtFN(1) & "strFromDirectory - " & strFromDirectory & SetVarFN(3))
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Delete intLast objFolder from full strPath. 
'!
'! @param  strVbsDirectory        Full strPath to objFolder.
'! @param  boolOnlyIfEmpty      If - remove recursively content of pointed objFolder.
'!
'! @return -1,0,1
Function fnRemoveDirectory(strDirectoryV, boolOnlyIfEmpty)
    Dim objFolder, strName, strVbsDirectory
    Dim strCurrentFNName                   : strCurrentFNName = "fnRemoveDirectory"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strDirectoryV - " & strDirectoryV & ", boolOnlyIfEmpty - " & boolOnlyIfEmpty & DotFN(0)) 
		
		If FnParamChecker(strDirectoryV, "strDirectoryV") = 1 Then
            call fnConst_VarFail("strDirectoryV", strCurrentFNName) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRemoveDirectory = 1
        Else
                strVbsDirectory = strDirectoryV
		End If
		
		If FnParamChecker(boolOnlyIfEmpty, "boolOnlyIfEmpty") = 1 Then
			fnAddLog(VarsFNWar(0) & "boolOnlyIfEmpty" & VarsFNWar(1) & "true" & DotFN(0))
                
			boolOnlyIfEmpty = true
		End If
            
            if fnLastCharChecker(strVbsDirectory, "\") = 0 Then
                strVbsDirectory = Left(strVbsDirectory, Len(strVbsDirectory) - 1)
            End If
		
		If fnIsExistD(strVbsDirectory) = 0 Then
			Set objFolder = objFSO.GetFolder(strVbsDirectory)
			
			If boolOnlyIfEmpty = true Then
				If objFolder.Files.Count = 0 And objFolder.SubFolders.Count = 0 Then
					Err.Clear
					objFSO.DeleteFolder strVbsDirectory, True
					If Err.Number <> 0 Then
                        fnAddLog(PreTxtFN(2) & DirsTxtFN(3) & strVbsDirectory & DotFN(3) & Err.Description & DotFN(0))
						fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))
						strLogTAB = strLogTABBak
						fnRemoveDirectory = 1
						Exit Function
					Else
                        fnAddLog(PreTxtFN(0) & DirsTxtFN(2) & strVbsDirectory & DotFN(0))
                        fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0)) 
						strLogTAB = strLogTABBak
						fnRemoveDirectory = 0
						Exit Function
					End If
				Else
					fnAddLog("FUNCTION INFO: ending fnRemoveDirectory with intRVal = -1. strVbsDirectory is not empty.")
					strLogTAB = strLogTABBak
					fnRemoveDirectory = -1
					Exit Function
				End If
				
			Else
				
			'Deleting all files in directory:
				For Each f In objFolder.Files
					On Error Resume Next
					strName = f.Name
					f.Delete True
					If Err Then
						fnAddLog(PreTxtFN(2) & DirsTxtFN(3) & Name & DotFN(3) & Err.Description & DotFN(0))
					Else
						fnAddLog(reTxtFN(0) & DirsTxtFN(2) & Name & DotFN(0))
					End If
					On Error GoTo 0
				Next
				
			'Deleting all subfolders in directory:
				For Each f In objFolder.SubFolders
					On Error Resume Next
					strName = f.Name
					f.Delete True
					If Err Then
						fnAddLog(PreTxtFN(2) & DirsTxtFN(3) & Name & DotFN(3) & Err.Description & DotFN(0))
					Else
						fnAddLog(PreTxtFN(0) & DirsTxtFN(2) & Name & DotFN(0))
					End If
					On Error GoTo 0
				Next
				
			'Deleting directory:
				Err.Clear
				objFSO.DeleteFolder strVbsDirectory, True
				If Err.Number <> 0 Then
                    fnAddLog(EndingFN(2) & strCurrentFNName & EndingFN(1) & Err.Number & DotFN(0))
			        fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
                        
					strLogTAB = strLogTABBak
					fnRemoveDirectory = 1
					Exit Function
				Else
					fnAddLog(PreTxtFN(0) & DirsTxtFN(2) & strVbsDirectory & DotFN(0))
                    fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0)) 
                        
					strLogTAB = strLogTABBak
					fnRemoveDirectory = 0
					Exit Function
				End If
				
			End If
		Else
			fnAddLog("FUNCTION INFO: strVbsDirectory - " & strVbsDirectory & " doesn't exist. Exiting function, nothing to do.")
			fnAddLog(EndingFN(0) & strCurrentFNName & DotFN(0)) 
			
			strLogTAB = strLogTABBak
			fnRemoveDirectory = 0
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Change Name of intLast objFolder from full strPath. 
'!
'! @param  strDirName          Full strPath to objFolder.
'! @param  strNewDirName       Full strPath to objFolder with new Name.
'!
'! @return 0,1
Function fnRenameDirectory(strDirName, strNewDirName) 
	Dim intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnRenameDirectory"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strDirName - " & strDirName & "strNewDirName - " & strNewDirName & DotFN(0)) 
		
		If FnParamChecker(strDirName, "strDirName") = 1 Then
            call fnConst_VarFail("strDirName", strCurrentFNName) 

			strLogTAB = strLogTABBak			
			fnFinalize(1)
			fnRenameDirectory = 1
		End If
		
		If FnParamChecker(strNewDirName, "strNewDirName") = 1 Then
            call fnConst_VarFail("strNewDirName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRenameDirectory = 1
		End If
            
            if fnLastCharChecker(strDirName, "\") = 0 Then
                strDirName = Left(strDirName, Len(strDirName) - 1)
            End If
            
            if fnLastCharChecker(strNewDirName, "\") = 0 Then
                strNewDirName = Left(strNewDirName, Len(strNewDirName) - 1)
            End If
            
		If fnIsExistD(strDirName) = 0 Then		
			intRVal = objFSO.MoveFolder(strDirName, strNewDirName)
            call Const_FnEndrVal(strCurrentFNName, intRVal)
			If intRVal = 1 Then 
				fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
				
				strLogTAB = strLogTABBak
				fnFinalize(1)
			End If
				
					fnAddLog(SelfCheckFN) 'added, check if working properly
					if fnIsExistD(strNewDirName) <> 0 Then
						fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1))  
						
						strLogTAB = strLogTABBak
						fnFinalize(1)
					End If
					
            fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0)) 
			
			strLogTAB = strLogTABBak
			fnRenameDirectory = intRVal
		Else
			fnAddLog(PreTxtFN(1) & "strDirName - " & strDirName & SetVarFN(3))
			fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(1)) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Check and provide strPath to current directory (wher is located script) 
'!
'! @return strPath
Function fnCurrentDir()
	Dim strCDir
    Dim strCurrentFNName                   : strCurrentFNName = "fnCurrentDir"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2)) 
            
			strCDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
			 
        fnAddLog(PreTxtFN(0) & DirsTxtFN(1) & strCDir & DotFN(0))
        fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0)) 
		
		strLogTAB = strLogTABBak
		fnCurrentDir = strCDir
End Function
'-------------------------------------------------------------------------------------------
'! Provide root drive letter
'!
'! @return drive
Function fnRootDrive()
	Dim strDrive
    Dim strCurrentFNName                   : strCurrentFNName = "fnRootDrive"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2)) 
		
			strDrive = objFSO.GetDriveName(strWindir)
			
		fnAddLog(PreTxtFN(0) & DirsTxtFN(0) & strDrive & DotFN(0)) 
        fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0)) 
		
		strLogTAB = strLogTABBak
		fnRootDrive = strDrive
End Function
'-------------------------------------------------------------------------------------------
'! It's copying setup files into IBMSRC objFolder and trigger substitute ARP entry, replacing uninstall string with our script.
'!
'! @return 0,1
Function fnSubstituteSelf()
	Dim intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnRenameDirectory"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
			intRVal = fnCreateDirectory(strTempForPKGSrc & "\" & strPKGName, False) 
			
			If fnIsExistD(strTempForPKGSrc & "\" & strPKGName) = 0 Then
				objFSO.CopyFolder strScriptDir , strTempForPKGSrc & "\" & strPKGName , True
				'intRVal = fnCopyDirectory(strScriptDir & "", strTempForPKGSrc & "\" & strPKGName & "\") 
			End If
	
		strLogTAB = strLogTABBak
		fnSubstituteSelf = 0
End Function
'-------------------------------------------------------------------------------------------
'! Copies folders from one location to all user profiles (including "Default") home directories.
'!
'! @param  strFromDirectory	Source from where the folders are to be copied to destination.
'! @param  strToDirectory		Destination where the folders from source are to be copied.
'!
'! @return intRVal
Function fnCopyFoldersToAllUsersProfiles(strFromDirectory, strToDirectory)
	Const HKEY_LOCAL_MACHINE = &H80000002
	Dim strSubKey, arrSubKeys, arrProfilePath, strValue, strUserName
	Dim strKeyPath						: strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	Dim strDefault						: strDefault = fnRootDrive & "\Users\Default"
	Dim intRVal							: intRVal = 0
	Dim strCurrentFNName					: strCurrentFNName = "fnCopyFoldersToAllUsersProfiles"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & ", strFromDirectory - " & strFromDirectory & ", strToDirectory - " & strToDirectory & DotFN(0))
		
		
		If FnParamChecker(strFromDirectory, "strFromDirectory") = 1 Then
			Call fnConst_VarFail("strFromDirectory", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnCopyFoldersToAllUsersProfiles = 1
		ElseIf strFromDirectory = "\" Then
			strFromDirectory = objFSO.GetParentFolderName(WScript.ScriptFullName)
			fnAddLog(SetVarFN(0) & "strFromDirectory" & SetVarFN(1) & strFromDirectory & DotFN(0))
		End If
		
		If FnParamChecker(strToDirectory, "strToDirectory") = 1 OR strToDirectory = "\" Then
			strToDirectory = ""
			fnAddLog(SetVarFN(0) & "strToDirectory" & SetVarFN(1) & "empty string" & DotFN(0))
		End If
		
	'----- Getting user profile home directory strPath:
		objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
		
		For Each strSubKey In arrSubKeys
			objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, strKeyPath & "\" & strSubKey, "ProfileImagePath", strValue
			
			arrProfilePath = Split(strValue, "\")					'Splitting user profile strPath to get user Name, which will be stored in the intLast array element.
			strUserName = arrProfilePath(UBound(arrProfilePath))	'Getting user Name (intLast array element) from user profile strPath.
		
	       '----- Copying file to user profile home directory (except for System and Service profiles):
			If NOT strUserName = "systemprofile" AND NOT strUserName = "LocalService" AND NOT strUserName = "NetworkService" Then
				intRVal = fnCopyDirectory(strFromDirectory, strValue & strToDirectory) 
				'fnCopyFileF(strFileName, strFromDirectory, strValue & strToDirectory)
			End If
		Next
		
	'----- Copying file to "Default" user profile home directory:
		If fnIsExistD(fnRootDrive & "\Users\Default\")  = 0 Then
			intRVal = fnCopyDirectory(strFromDirectory, fnRootDrive & "\Users\Default" & strToDirectory) 
		End If
		'fnCopyFileF(strFileName, strFromDirectory, strDefault & strToDirectory)
		
		fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0))
		
		strLogTAB = strLogTABBak
		fnCopyFoldersToAllUsersProfiles = intRVal
End Function
'-------------------------------------------------------------------------------------------'
'! Deletes objFolder from all user profiles (including "Default") home directories.
'!
'! @param  strDirectoryV		Path to objFolder.
'! @param  boolOnlyIfEmpty      If true, then remove recursively content of pointed objFolder.
'!
'! @return -1, 0, 1
Function fnRemoveDirectoryFromAllUsersProfiles(strDirectoryV, boolOnlyIfEmpty)
	Const HKEY_LOCAL_MACHINE = &H80000002
	Dim strSubKey, arrSubKeys, arrProfilePath, strValue, strUserName, strVbsDirectory
	Dim strKeyPath						: strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	Dim strDefault						: strDefault = fnRootDrive & "\Users\Default"
	Dim intRVal							: intRVal = 0
	Dim strCurrentFNName					: strCurrentFNName = "fnRemoveDirectoryFromAllUsersProfiles"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strDirectoryV - " & strDirectoryV & ", boolOnlyIfEmpty - " & boolOnlyIfEmpty & DotFN(0))
		
		If FnParamChecker(strDirectoryV, "strDirectoryV") = 1 Then
			Call fnConst_VarFail("strDirectoryV", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRemoveDirectoryFromAllUsersProfiles = 1
		Else
			strVbsDirectory = strDirectoryV
		End If
		
		If FnParamChecker(boolOnlyIfEmpty, "boolOnlyIfEmpty") = 1 Then
			fnAddLog(VarsFNWar(0) & "boolOnlyIfEmpty" & VarsFNWar(1) & "True" & DotFN(0))
			
			boolOnlyIfEmpty = True
		End If
		
			If fnLastCharChecker(strVbsDirectory, "\") = 0 Then
				strVbsDirectory = Left(strVbsDirectory, Len(strVbsDirectory) - 1)
			End If
		
		
	'----- Getting user profile home directory strPath:
		objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys
		
		For Each strSubKey In arrSubKeys
			objReg.GetExpandedStringValue HKEY_LOCAL_MACHINE, strKeyPath & "\" & strSubKey, "ProfileImagePath", strValue
			
			arrProfilePath = Split(strValue, "\")					'Splitting user profile strPath to get user Name, which will be stored in the intLast array element.
			strUserName = arrProfilePath(UBound(arrProfilePath))	'Getting user Name (intLast array element) from user profile strPath.
		
		
	'----- Deleting objFolder from user profile home directory (except for System and Service profiles):
			If NOT strUserName = "systemprofile" AND NOT strUserName = "LocalService" AND NOT strUserName = "NetworkService" Then
				intRVal = fnRemoveDirectory(strValue & strVbsDirectory, boolOnlyIfEmpty)
			End If
		
		Next
		
		
	'----- Deleting objFolder from "Default" user profile home directory:
		intRVal = fnRemoveDirectory(strDefault & strVbsDirectory, boolOnlyIfEmpty)
		
		
		fnAddLog(EndingFN(0) & strCurrentFNName & SucFailFN(0))
		
		strLogTAB = strLogTABBak
		fnRemoveDirectoryFromAllUsersProfiles = intRVal
End Function