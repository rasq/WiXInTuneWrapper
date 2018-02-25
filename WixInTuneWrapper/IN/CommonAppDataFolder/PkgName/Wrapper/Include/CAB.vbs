'****************************************************************************************************************************************************
' FUNCTIONS: 
'	fnCABFilesValidate() 'it will use global var arrCABs
'	fnCopyCABsToMachine() 'it will use global var arrCABs  
'	fnCheckIfAlredyUnpacked(strCABDirectory) 
'	fnUncompress(strCABDirectory)  'it will use global var arrCABs, and global for IBMTemp, strCABDirectory is for arrCABs folder 
'	fnRemoveCABsFromMachine()  'it will use global var arrCABs 
'	fnRemoveUncompressedFiles(strCABDirectory)
'	
' FUNCTION EXAMPLE USAGE (IN and OUT):
' 
'***************************************************************************************************************************************************
'! Check if all CAB files are in place before unpacking. 
'!
'! @return 0,1
Function fnCABFilesValidate()
    Dim strCurrentFNName                : strCurrentFNName = "fnCABFilesValidate"
	Dim CABALength 						: CABALength = UBound(arrCABs)
	Dim strLogTABBak					: strLogTABBak = strLogTAB
	Dim i 								: i = 0
		strLogTAB = strLogTAB & strLogSeparator
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
		
				For Each item In arrCABs
					If FnParamChecker(arrCABs(i), "arrCABs(" & i & ")") = 1 Then
                        call fnConst_VarFail("arrCABs(" & i & ")", strCurrentFNName) 
						
						strLogTAB = strLogTABBak
						fnFinalize(1)
						fnCABFilesValidate = 1
					End If
					
					i = i+1
				Next	
				
				i = 0
				For Each item In arrCABs
					If fnIsExist(item, "\") = 0 Then
						fnAddLog("FUNCTION INFO: arrCABs(" & i & ")- " & item & " exist. Going to next step.")
					Else
						fnAddLog("FUNCTION ERROR: arrCABs(" & i & ")- " & item & " don't exist. Exiting function.")
						fnAddLog("FUNCTION INFO: ending fnCABFilesValidate - fail.") 
						
						strLogTAB = strLogTABBak
						fnFinalize(1)				
						fnCABFilesValidate = 1
					End If
					
					i = i+1
				Next	
				
		fnAddLog("FUNCTION INFO: fnCABFilesValidate, all files exist.")
		fnAddLog("FUNCTION INFO: ending fnCABFilesValidate - success.") 
		
		strLogTAB = strLogTABBak
		fnCABFilesValidate = 0
End Function
'-------------------------------------------------------------------------------------------
'! Copy all CAB files to client IBMSRC directory. 
'!
'! @return 0,1
Function fnCopyCABsToMachine() 
    Dim strCurrentFNName                : strCurrentFNName = "fnCopyCABsToMachine"
	Dim strLogTABBak					: strLogTABBak = strLogTAB
	Dim i 								: i = 0
		strLogTAB = strLogTAB & strLogSeparator
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
		
				For Each item In arrCABs
					If FnParamChecker(arrCABs(i), "arrCABs(" & i & ")") = 1 Then
                        call fnConst_VarFail("arrCABs(" & i & ")", strCurrentFNName)
						
						strLogTAB = strLogTABBak
						fnFinalize(1)
						fnCopyCABsToMachine = 1
					End If
					
					i = i+1
				Next	
				
				i = 0
				For Each item In arrCABs
					If fnCopyFileF(arrCABs(i), "\", strTempForPKGSrc) = 0 Then
						fnAddLog("FUNCTION INFO: arrCABs(" & i & ")- " & arrCABs(i) & " successfully copied from currentDir to strTempForPKGSrc - " & strTempForPKGSrc & ". Going to next step.")
					Else
						fnAddLog("FUNCTION ERROR: arrCABs(" & i & ")- " & arrCABs(i) & " don't copied. Exiting function.")
						fnAddLog("FUNCTION INFO: ending fnCopyCABsToMachine - fail.") 
						
						strLogTAB = strLogTABBak
						fnFinalize(1)				
						fnCopyCABsToMachine = 1
					End If
					
					i = i+1
				Next	
		
		fnAddLog("FUNCTION INFO: fnCopyCABsToMachine, all files copied.")
		fnAddLog("FUNCTION INFO: ending fnCopyCABsToMachine - success.") 
		
		strLogTAB = strLogTABBak
		fnCopyCABsToMachine = 0
End Function 
'-------------------------------------------------------------------------------------------
'! Uncab all cab files into specified directory.
'!
'! @param  strCABDirectory        Destination directory.
'!
'! @return 0,1
Function fnUncompress(strCABDirectory) 
	Dim intRVal, strUncCMD
    Dim strCurrentFNName                : strCurrentFNName = "fnUncompress"
	Dim i 								: i = 0
	Dim strLogTABBak					: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strCABDirectory - " & strCABDirectory & DotFN(0)) 
		
				For Each item In arrCABs
					If FnParamChecker(arrCABs(i), "arrCABs(" & i & ")") = 1 Then
                        call fnConst_VarFail("arrCABs(" & i & ")", strCurrentFNName)
						
						strLogTAB = strLogTABBak
						fnFinalize(1)
						fnUncompress = 1
					End If
					
					i = i+1
				Next	
				
				i = 0
		
		If FnParamChecker(strCABDirectory, "strCABDirectory") = 1 Then
            call fnConst_VarFail("strCABDirectory", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnUncompress = 1
		ElseIf strCABDirectory = "\" Then
			strCABDirectory = objFSO.GetParentFolderName(WScript.ScriptFullName)
			fnAddLog("FUNCTION INFO: translation variable strCABDirectory \ to - " & strCABDirectory & ".")
		End If
		
		If fnIsExistD(strTempForPKGSrc) = 0 Then 
			if fnCheckIfAlredyUnpacked(strTempForPKGSrc) = 0 Then
				fnAddLog("FUNCTION INFO: fnUncompress, strTempForPKGSrc: " & strTempForPKGSrc & " exist, and arrCABs are decompressed, going to next step.")
				fnAddLog("FUNCTION INFO: ending fnUncompress - success.") 
				
				strLogTAB = strLogTABBak
				fnUncompress = 0
				Exit Function
			Else
				fnAddLog("FUNCTION INFO: fnUncompress, strTempForPKGSrc: " & strTempForPKGSrc & " exist, but arrCABs are not decompressed.")
			End If
		Else
			If fnCreateDirectory(strTempForPKGSrc, true) = 0 Then
				fnAddLog("FUNCTION INFO: fnUncompress, strTempForPKGSrc: " & strTempForPKGSrc & " created, going to next step.")
			Else
				fnAddLog("FUNCTION INFO: fnUncompress, strTempForPKGSrc: " & strTempForPKGSrc & " creation failed.")
				fnAddLog("FUNCTION INFO: ending fnUncompress - fail.") 
				
				strLogTAB = strLogTABBak
				fnFinalize(1)	
				fnUncompress = 1
			End If
		End If
		
		strUncCMD = Chr(34) & strCABDirectory & "\" & arrCABs(0) & Chr(34) & " /Y /L " & Chr(34) & strTempForPKGSrc & Chr(34)
		fnAddLog("FUNCTION INFO: running strUncCMD: " & strUncCMD & ".")
		intRVal = objShell.Run (strUncCMD, 0, true)		
		
			if intRVal = 0 Then
				fnAddLog("FUNCTION INFO: Uncompressed all files.")
				fnAddLog("FUNCTION INFO: ending fnUncompress - success.") 
				
				strLogTAB = strLogTABBak
				fnUncompress = 0
				Exit Function
			Else
				fnAddLog("FUNCTION INFO: cannot uncompress all files. intRVal = " & intRVal & ".")
				fnAddLog("FUNCTION INFO: ending fnUncompress - fail.") 
				
				strLogTAB = strLogTABBak
				fnFinalize(intRVal)	
				fnUncompress = intRVal
			End If
End Function 
'-------------------------------------------------------------------------------------------
'! Check if there is a unpacked directory.
'!
'! @param  strCABDirectory        Destination directory.
'!
'! @return 0,1
Function fnCheckIfAlredyUnpacked(strCABDirectory) 
	Dim objFolderA, intDatafolder
    Dim strCurrentFNName                : strCurrentFNName = "fnCheckIfAlredyUnpacked"
	Dim strLogTABBak					: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strCABDirectory - " & strCABDirectory & DotFN(0)) 
		
		If FnParamChecker(strCABDirectory, "strCABDirectory") = 1 Then
            call fnConst_VarFail("strCABDirectory", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnCheckIfAlredyUnpacked = 1
		End If
			
		If fnIsExistD(strCABDirectory) = 0 Then
			Set objFolderA = objFSO.GetFolder(strCABDirectory)
				intDatafolder = objFolderA.Size / 1024 / 1024
				fnAddLog("FUNCTION INFO: fnCheckIfAlredyUnpacked, directory: " & strCABDirectory & " exist. Current size: " & intDatafolder & ".")
		Else
			fnAddLog("FUNCTION INFO: fnCheckIfAlredyUnpacked, directory: " & strCABDirectory & " doesn't exist. Files are not unpacked.")
			
			strLogTAB = strLogTABBak
			fnCheckIfAlredyUnpacked = 1
			Exit Function
		End If
	  
	  
		if CStr(intDatafolder) = CStr(UncabbedDirSize) then
			fnAddLog("FUNCTION INFO: fnCheckIfAlredyUnpacked, files are in place.")
			
			strLogTAB = strLogTABBak
			fnCheckIfAlredyUnpacked = 0
			Exit Function
		Else 
			fnAddLog("FUNCTION INFO: strCABDirectory = " & intDatafolder & " <> " & UncabbedDirSize) 
			fnAddLog("FUNCTION INFO: fnCheckIfAlredyUnpacked, directory: " & strCABDirectory & " exist. But files are not unpacked.")
			
			strLogTAB = strLogTABBak
			fnCheckIfAlredyUnpacked = 1
			Exit Function
		End if
End Function
'-------------------------------------------------------------------------------------------
'! Remove all copied CAB files from client machine.
'!
'! @return 0,1
Function fnRemoveCABsFromMachine()
    Dim strCurrentFNName                : strCurrentFNName = "fnRemoveCABsFromMachine"
	Dim i 								: i = 0
	Dim boolWasFail 					: boolWasFail = false
	Dim strLogTABBak					: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
        fnAddLog(StartingFN(0) & strCurrentFNName & DotFN(2))
		
				For Each item In arrCABs
					If FnParamChecker(arrCABs(i), "arrCABs(" & i & ")") = 1 Then
                        call fnConst_VarFail("arrCABs(" & i & ")", strCurrentFNName)
						
						strLogTAB = strLogTABBak
						fnFinalize(1)
						fnRemoveCABsFromMachine = 1
					End If
					
					i = i+1
				Next	
				
				i = 0
		
			For Each item In arrCABs
				If fnRemoveFile(item, strTempForPKGSrc) = 0 Then
					fnAddLog("FUNCTION INFO: fnRemoveCABsFromMachine(" & i & ")- " & item & " successfully removed from strTempForPKGSrc. Going to next step.")
				Else
					fnAddLog("FUNCTION ERROR: fnRemoveCABsFromMachine(" & i & ")- " & item & " don't removed. Going to next step.")
					boolWasFail = true
				End If
					
				i = i+1
			Next	
				
			if boolWasFail = true Then
				fnAddLog("FUNCTION INFO: fnRemoveCABsFromMachine, not all files was removed from machine.")
				
				strLogTAB = strLogTABBak
				fnRemoveCABsFromMachine = 1
				Exit Function
			Else
			
			End If
End Function
'-------------------------------------------------------------------------------------------
'! Remove all unpacked files and directory from destination directory.
'!
'! @param  strCABDirectory        Destination directory.
'!
'! @return 0,1
Function fnRemoveUncompressedFiles(strCABDirectory)
    Dim strCurrentFNName               	: strCurrentFNName = "fnRemoveUncompressedFiles"
	Dim strLogTABBak					: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strCABDirectory - " & strCABDirectory & DotFN(0)) 
		
		If FnParamChecker(strCABDirectory, "strCABDirectory") = 1 Then
            call fnConst_VarFail("strCABDirectory", strCurrentFNName) 
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRemoveUncompressedFiles = 1
			Exit Function
		End If
		
		if fnRemoveDirectory(strCABDirectory, false) = 1 Then
            fnAddLog(EndingFN(2) & strCurrentFNName & SucFailFN(1)) 
			
			strLogTAB = strLogTABBak
			fnRemoveUncompressedFiles = 1
			Exit Function
		End If
		
		fnAddLog("FUNCTION INFO: ending fnRemoveUncompressedFiles - success.") 
		
		strLogTAB = strLogTABBak
		fnRemoveUncompressedFiles = 0
End Function