'****************************************************************************************************************************************************
' FUNCTIONS: 
'	fnAddTag(strLocalPKGName)   
'	fnRemoveTag(strLocalPKGName)  
'	fnIsExistTag(strLocalPKGName) 
'	
' FUNCTION EXAMPLE USAGE (IN and OUT):
' 
'***************************************************************************************************************************************************
'! Check if tag file is in place (in client TAG directory). 
'!
'! @param  strLocalPKGName         Tag file Name.
'!
'! @return 0,1
Function fnIsExistTag(strLocalPKGName) 
	Dim intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnIsExistTag"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strLocalPKGName - " & strLocalPKGName & ".") 
			
		If FnParamChecker(strLocalPKGName, "strLocalPKGName") = 1 Then
            call fnConst_VarFail("strLocalPKGName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnIsExistTag = 1
		End If
		
			if TAGType = "FILE" Then
				intRVal = fnIsExist(strLocalPKGName & ".tag", strTAGFolder)
			ElseIf TAGType = "REG" Then 
					'TODO: registry tag template checking.
			ElseIf TAGType = "NONE" Then
				fnAddLog(EndingFN(0) & strCurrentFNName & TAGTxtFN(3)) 
				
				strLogTAB = strLogTABBak
				fnIsExistTag = 1 
				Exit Function
			End If
			
		If intRVal = 0 Then
			fnAddLog(EndingFN(0) & strCurrentFNName & EndingFN(1) & intRVal & ", " & SucFailFN(0))
			
			strLogTAB = strLogTABBak
			fnIsExistTag = 0
			Exit Function
		ElseIf intRVal = 1 Then
			fnAddLog(EndingFN(0) & strCurrentFNName & EndingFN(1) & intRVal & ", " & SucFailFN(1))
			
			strLogTAB = strLogTABBak
			fnIsExistTag = 1
			Exit Function
		End If
End Function
'-------------------------------------------------------------------------------------------
'! Create tag file(in client TAG directory). 
'!
'! @param  strLocalPKGName         Tag file Name.
'!
'! @return 0,1
Function fnAddTag(strLocalPKGName)  
	Dim intRVal
    Dim strCurrentFNName                   : strCurrentFNName = "fnAddTag"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strLocalPKGName - " & strLocalPKGName & ".")
		
		If FnParamChecker(strLocalPKGName, "strLocalPKGName") = 1 Then
            call fnConst_VarFail("strLocalPKGName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnAddTag = 1
		End If	
		    
		
			If TAGType = "FILE" Then
				
				if objFSO.FolderExists(strTAGFolder) = False then
					fnCreateDirectory strTAGFolder,NULL
				End if
								
				intRVal = fnAddTXTFile(strLocalPKGName & ".tag", strTAGFolder) 
			ElseIf TAGType = "REG" Then 
					'TODO: registry tag template to add.
			ElseIf TAGType = "NONE" Then
				fnAddLog(EndingFN(0) & strCurrentFNName & TAGTxtFN(2)) 
				
				strLogTAB = strLogTABBak
				fnAddTag = 0
				Exit Function
			End If
			
        call Const_FnEndrVal(strCurrentFNName, intRVal)
		
		strLogTAB = strLogTABBak
		fnAddTag = intRVal
End Function
'-------------------------------------------------------------------------------------------
'! Remove tag file(from client TAG directory). 
'!
'! @param  strLocalPKGName         Tag file Name.
'!
'! @return 0,1
Function fnRemoveTag(strLocalPKGName)
	Dim intRVal, objFolder
    Dim strCurrentFNName                   : strCurrentFNName = "fnRemoveTag"
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "strLocalPKGName - " & strLocalPKGName & ".") 
		
		Set objFolder = objFSO.GetFolder(strTAGFolder)
		
		If FnParamChecker(strLocalPKGName, "strLocalPKGName") = 1 Then
            call fnConst_VarFail("strLocalPKGName", strCurrentFNName)
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			fnRemoveTag = 1
		End If
		
			if TAGType = "FILE" Then
				If fnIsExistTag(strLocalPKGName) = 0 Then
					intRVal = fnRemoveFile(strLocalPKGName & ".tag", strTAGFolder) 
						If intRVal = 0 Then
							If objFolder.Files.Count = 0 And objFolder.SubFolders.Count = 0 Then
								intRVal = fnRemoveDirectory(strTAGFolder, True)
							End If
						End If
				Else   
					fnAddLog(EndingFN(0) & strCurrentFNName & DotFN(2) & TAGTxtFN(1)) 
					
					strLogTAB = strLogTABBak
					fnRemoveTag = 0
					Exit Function
				End If
			ElseIf TAGType = "REG" Then 
				If fnIsExistTag(strLocalPKGName) = 0 Then
					'TODO: registry tag template remove.
				Else 
					fnAddLog(EndingFN(0) & strCurrentFNName & DotFN(2) & TAGTxtFN(1))
					
					strLogTAB = strLogTABBak
					fnRemoveTag = 0
					Exit Function
				End If
			ElseIf TAGType = "NONE" Then
				fnAddLog(EndingFN(0) & strCurrentFNName & DotFN(2) & TAGTxtFN(0))
				
				strLogTAB = strLogTABBak
				fnRemoveTag = 0
				Exit Function
			End If
			
        call Const_FnEndrVal(strCurrentFNName, intRVal)
		
		strLogTAB = strLogTABBak
		fnRemoveTag = intRVal
End Function