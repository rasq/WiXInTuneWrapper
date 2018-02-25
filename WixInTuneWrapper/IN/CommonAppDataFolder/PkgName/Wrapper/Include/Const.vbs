Dim Const_StrikeEC				  : Const_StrikeEC			= -1000
Dim Const_isServEC				  : Const_isServEC			= -1001
Dim Const_newVerEC				  : Const_newVerEC			= -1002
Dim Const_noCoreVerEC			  : Const_noCoreVerEC		= -1003
Dim Const_bAppInEC				  : Const_bAppInEC			= -1004
Dim Const_procRunningEC			  : Const_procRunningEC		= -1005
Dim Const_RebootPEC				  : Const_RebootPEC 		= -1006
Dim Const_notEHDDEC				  : Const_notEHDDEC			= -1007
Dim Const_ReqPortalNotEC	      : Const_ReqPortalNotEC	= -1008
Dim Const_VendorInstalled         : Const_VendorInstalled   = -1009
Dim Const_toRetryEC				  : Const_toRetryEC			= 1618


Dim PreTxtFN                      : PreTxtFN    = Array("FUNCTION INFO: ","FUNCTION ERROR: ","FUNCTION WARRNING: ","SETUP INFO: ","SETUP ERROR: ")
Dim StartingFN                    : StartingFN  = Array(PreTxtFN(0) & "starting "," function with parameters: ")
Dim EndingFN                      : EndingFN    = Array(PreTxtFN(0) & "ending "," function - with intRVal: ",PreTxtFN(1) & "ending ")
Dim SucFailFN                     : SucFailFN   = Array(" function - SUCCESS."," function - FAIL.")
Dim DotFN                         : DotFN       = Array("."," ."," function."," - ")
Dim VarsFNErr                     : VarsFNErr   = Array(PreTxtFN(1) & "variable "," cannot be NULL.", " is empty or null.")
Dim VarsFNWar                     : VarsFNWar   = Array(PreTxtFN(2) & "variable "," is NULL. Setting to ")
Dim ExitFN                        : ExitFN      = Array(PreTxtFN(0) & "variable ",VarsFNWar(1))
Dim SetupSucFN                    : SetupSucFN  = Array(PreTxtFN(3) & "installation success.",PreTxtFN(3) & "uninstallation success.")
Dim SetupFaiFN                    : SetupFaiFN  = Array(PreTxtFN(4) & "installation failed.",PreTxtFN(4) & "uninstallation failed.")
Dim StartLogFN                    : StartLogFN  = Array("=== Logging started: ",". Calling process: "," ===")
Dim StopLogFN                     : StopLogFN   = Array("=== Logging stopped: ",StartLogFN(1),StartLogFN(2))
Dim LineFN                        : LineFN      = "========================================================================================"
Dim TAGTxtFN                      : TAGTxtFN    = Array(" This account don't use TAG files, removing skipped."," TAG doesn't exist."," This account don't use TAG files, creation skipped."," This account don't use TAG files, checking skipped.")
Dim SetVarFN                      : SetVarFN    = Array(PreTxtFN(0) & "setting variable: "," , to: "," exist, going to next step."," doesn't exist."," created, going to next step."," found.")
Dim CreationFN                    : CreationFN  = Array(" created."," creation failed."," exist. Creation skipped."," created, going to next step."," creating file"," - calling: "," exist, starting renaming.","renaming existing ")
Dim EditFN                        : EditFN      = Array(" replaced string "," ,edited file - "," added to file - ")
Dim SelfCheckFN                   : SelfCheckFN = PreTxtFN(0) & "self check."
Dim DirsTxtFN                     : DirsTxtFN   = Array("current root drive - ","current directory - ","Deleted: ","Error deleting: ")

'***************************************LOGS TEMPLATES****************************************************
'
'
'
'
'
'
'
'
'
'
'
'
'
'***************************************LOGS TEMPLATES****************************************************


'! Generate log entry about missing variable value. To use when variable is essential for function.
'!
'! @param  strFunctionName     Function Name.
'! @param  VarName          Variable Name.
'!
'! @return N/A.
Function fnConst_VarFail(VarName, strFunctionName)
    If fnIsObject(objLog) Then 
        fnAddLog(VarsFNErr(0) & VarName & VarsFNErr(1))
        fnAddLog(EndingFN(0) & strFunctionName & SucFailFN(1))
    End If
End Function
'-------------------------------------------------------------------------------------------
'! Generate log entry about empty variable set to empty string. 
'!
'! @param  strFunctionName     Function Name.
'! @param  VarName          Variable Name.
'!
'! @return N/A.
Function Const_VarEmpty(VarName, strFunctionName)
    
End Function
'-------------------------------------------------------------------------------------------
'! Generate log entry about empty variable set to default value. 
'!
'! @param  strFunctionName     Function Name.
'! @param  VarName          Variable Name.
'! @param  VarValue         Variable new value.
'!
'! @return N/A.
Function Const_VarSetTo(VarName, VarValue, strFunctionName)
    
End Function
'-------------------------------------------------------------------------------------------
'! Generate log entry about intRVal returned by function.
'!
'! @param  strFunctionName     Function Name.
'! @param  intRVal             Function exit code.
'!
'! @return N/A.
Function Const_FnEndrVal(strFunctionName, intRVal)
    If fnIsObject(objLog) Then 
        fnAddLog(EndingFN(0) & strFunctionName & EndingFN(1) & intRVal & DotFN(0))
    End If
End Function
'-------------------------------------------------------------------------------------------
