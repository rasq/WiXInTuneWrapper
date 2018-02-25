'****************************************************************************************************************************************************	
' FUNCTIONS: 
'	fnPrereqsInstall()
'	fnPreInstall() 
'	fnPostInstall() 
'	fnPreUninstall() 
'	fnPostUninstall()  
'	Install()   
'	Uninstall()
'	fnInclude(strFunctionFile)
'	fnRollback()
' 
' 
' Custom ExitCodes to use:
' -1000 -> 5th, 3th strike exit. 										(Const_StrikeEC)
' -1001 -> isServer. 													(Const_isServEC)
' -1002 -> newer version installed. 									(Const_newVerEC)
' -1003 -> core app not installed (for patch installation). 			(Const_noCoreVerEC)
' -1004 -> app with more functionality is already installed. 			(Const_bAppInEC)
' -1005 -> process is running and cannot be terminated. 				(Const_procRunningEC)
' -1006 -> reboot pending. 												(Const_RebootPEC)
' -1007 -> not enough free disk space. 									(Const_notEHDDEC)
' -1008 -> fnRequestPortal files absent when needed. 						(Const_ReqPortalNotEC)
' -1009 -> vendor installed					                            (Const_VendorInstalled)
'  1618 -> info for SCCM to retry installation (msi error code). 		(Const_toRetryEC)
'
'***************************************************************************************************************************************************
Option Explicit 

' Account Declarations ----------------------------------------------------------------------
	Dim strTAGFolder, strTempForPKGSrc, strAccountName, intTimeVar
	
		' Chose per account settings----------------------------------------------------
				strAccountName = "GLOBAL" 'for purpose of client configuration include
		' Custom per account settings section end---------------------------------------
'--------------------------------------------------------------------------------------------
' Core Variable Declarations ----------------------------------------------------------------
		Dim objFSO
			Set objFSO					= CreateObject("Scripting.FileSystemObject")
			
		Dim strRootDrive, objWMIService, objLog, objFile, objReg, objVarClass, objShell, colProcessList, strScriptPath
		Dim boolSilentInstall				: boolSilentInstall = false
		Dim boolFSilentInstall				: boolFSilentInstall = false
		Dim boolRebootContainer 			: boolRebootContainer = false
		Dim strComputer						: strComputer = "."
		Dim strSelectedLang					: strSelectedLang = ""
		Dim strLogTAB						: strLogTAB = ""
		Dim strLogSeparator					: strLogSeparator = "------"
			Set objShell 				= CreateObject("WScript.Shell")
				strScriptPath 			= Wscript.ScriptFullName
			Set objFile 				= objFSO.GetFile(strScriptPath)
			Set objVarClass 			= GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2:Win32_Environment")
			Set objWMIService 			= GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
			Set objReg 					= GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
				strScriptPath 			= objFSO.GetParentFolderName(objFile) 	
					
		Dim strIniFilePath					: strIniFilePath = ""
			strIniFilePath 				= LCase(objFSO.GetParentFolderName(WScript.ScriptFullName)) & "\Include\AccountSpecific\" & strAccountName & ".ini"
			
		Dim strIniConfFilePath				: strIniConfFilePath = ""
			strIniConfFilePath 			= LCase(objFSO.GetParentFolderName(WScript.ScriptFullName)) & "\conf.ini"
'--------------------------------------------------------------------------------------------
' Includes ----------------------------------------------------------------------------------
	Dim arrIncludes : arrIncludes = Array("Include\Logs.vbs","Include\Const.vbs","Include\DirectoryOperations.vbs","Include\FilesOperations.vbs","Include\RegistryOperations.vbs","Include\OS.vbs","Include\MSI.vbs","Include\EXE.vbs","Include\CAB.vbs","Include\ExternalVBS.vbs","Include\StrikesMSGBoxes.vbs","Include\TAGOperations.vbs","Include\Processes.vbs","Include\Services.vbs","Include\ProgresBar.vbs","Include\Other.vbs","Include\TempOperations.vbs")
		
		For Each item In arrIncludes
			Call fnInclude(item)
		Next	
        
'--------------------------------------------------------------------------------------------
 
' Core Variable Declarations ----------------------------------------------------------------
		strRootDrive = fnOSVariableTranslation("%SYSTEMDRIVE%", false)
'--------------------------------------------------------------------------------------------

' PKG Variable Declarations -----------------------------------------------------------------
	Dim strPKGName                      : strPKGName 				= "PKGName" '<vendor>_<strAppName>_<appVersion>_<account>_<os>_<lang>_<build>
	Dim strFriendlyPKGName              : strFriendlyPKGName 		= "FriendlyPKGName" 'friendly Name for dialogue boxes, if PKG Name use some sort of ID, then it must be included
				'All arrMSI/arrMST/arrMSP and arrEXE/VBS files we will store in arrays. Future usage will be simpler because of one variable for each type. 
				'SAMPLE: fnInstallMSI(arrMSI(0), arrMST(0), arrMSP(0), strAdditionalParams, strVbsDirectory), fnInstallMSI(arrMSI(1), arrMST(1), arrMSP(1), strAdditionalParams, strVbsDirectory), etc.
	Dim arrGUIDs 						: arrGUIDs 					= Array("GUIDs01")	'additional arrGUIDs for uninstall - besides main appGUID. You can use them to 
																	'check if additional package products are installed, or to uninstall them.
	Dim arrMSI 							: arrMSI 					= Array("MSI01")
	Dim arrMST 							: arrMST 					= Array("MST01") 
	Dim arrMSP 							: arrMSP 					= Array("MsPFile00")
	Dim arrEXE 							: arrEXE 					= Array("EXEFile00")
	Dim arrExVBS 						: arrExVBS 					= Array("ExVBSFile00", "ExVBSFile01")
	Dim arrCABs							: arrCABs 					= Array("CABs01")							
	Dim strProcessessToKill 			: strProcessessToKill 		= "Process00,Process01,Process01" 	'processes to kill, list divided by "," without any spaces!!!
	Dim strFriendlyProcessName          : strFriendlyProcessName 	= "strFriendlyProcessName" 'friendly Name for dialogue boxes
	Dim strLanguagesToUse				: strLanguagesToUse 		= "English,German,Dutch,Japanese,French,Spanish,Italian,Turkish"		'Languages, list divided by "," without any spaces!!! Use like processes to kill - only if needed 
	Dim intProcessHTAWindowTime			: intProcessHTAWindowTime 	= 7200  ' In sec., 60 = 1min...., 7200 = 2h. If persistent window then in minutes.  
				  
	Dim strAppName 						: strAppName 				= "AppName"		'ARP Name for main application.
	Dim	strAppNumber 					: strAppNumber 				= "AppNumber"	'ARP version for main application.
	Dim	strPkgGUID 						: strPkgGUID 				= "PkgGUID"		'main GUID for package if you have multi product installation (or if you 
																	'will install one .msi package, this will be strGUID for this .msi).
	Dim strLanguageCode                 : strLanguageCode 			= "1033"     'Variable for language selection. Please use mst language codes. 
                                                                    'It will be used in language selection window and for proper language installation.	
	Dim boolRebootNeeded				: boolRebootNeeded 			= False 	'Use this variable only on account where information about reboot will be displayed before installation.
	Dim boolLogOffNeeded				: boolLogOffNeeded 			= False	'Use this variable only on account where information about log off will be displayed before installation.										
    Dim intInstallationTime             : intInstallationTime 		= 0      'Change only on client where information about estimated instalaltion time is displayed.
	
	Call fnInclude("Include\PKGConfig.vbs")' - ini to fill PKG Variable	
'-------------------------------------------------------------------------------------------
' Script Variable Declarations -------------------------------------------------------------
	Dim boolThisInstall, intRetCode, strLogName, intFuncRetCode
	
	Dim strScriptHostName 				: strScriptHostName 		= LCase(Wscript.ScriptName)
	Dim strScriptDir					: strScriptDir 				= LCase(objFSO.GetParentFolderName(WScript.ScriptFullName))
	Dim strWindir 						: strWindir 				= fnOSVariableTranslation("%WINDIR%", false)	
	Dim strFolder 						: strFolder 				= strWindir & chr(92) & "Sysnative"
	
	Dim strProgramFiles

		If fnIs64BitOS(false) = 0 Then
			strProgramFiles = fnOSVariableTranslation("%programfiles(x86)%", false) 	'%programfiles(x86)% - if 64bit OS
		Else
			strProgramFiles = fnOSVariableTranslation("%programfiles%", false) 			'%programfiles% - if 32bit OS
		End If
				
	Call fnInclude("Include\AccountSpecific.vbs")' - moved to per account include		
			
	Dim strLogFolder 					: strLogFolder = PreLogDir & "\Logs"		'Logs destination folder (only strProgramFiles can be changed)
	Dim UncabbedDirSize 				: UncabbedDirSize = 10715.3754110336 'size in MB for uncabbed directory
    Dim strSuccessTable                 : strSuccessTable = "0,3010"
		strTAGFolder = PreTAGDir & "\Logs"
				
		fnPreSetup()
'-------------------------------------------------------------------------------------------
' Packager Variable Declarations -----------------------------------------------------------
	Dim strMachineTmp 					: strMachineTmp = strWindir & "\PKGCache" 	'or rootdrive or other
		strTempForPKGSrc = strMachineTmp
		


		' Put here your variables-------------------------------------------------------
			'All custom variable names must be started with "Custom_" string in Name, e.g.: Dim Custom_TempRC
    
        
			
			
		' Custom variables section end--------------------------------------------------
    
    

        fnInitialize() 'First function to start. Please do not put any other code before it.
'-------------------------------------------------------------------------------------------






' Put here additional PKG actions ----------------------------------------------------------
'! Function for all actions regarding prerequisites. 
'!
'! @return 0,1
Function fnPrereqsInstall()
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		intFuncRetCode = 0
		
		fnAddLog(StartingFN(0) & "fnPrereqsInstall" & DotFN(2))  
			' Put here your code-------------------------------------
			
			
			
			
			' Custom code section end--------------------------------
		fnAddLog(EndingFN(0) & "fnPrereqsInstall" & DotFN(2)) 
		
		strLogTAB = strLogTABBak
		fnPrereqsInstall = intFuncRetCode
End Function
'-------------------------------------------------------------------------------------------
'! Function for all actions regarding preinstall actions like unpack archive before installation. 
'!
'! @return 0,1
Function fnPreInstall()
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		intFuncRetCode = 0
		 
		fnAddLog(StartingFN(0) & "fnPreInstall" & DotFN(2))  
			' Put here your code-------------------------------------
			
			
			
			
			' Custom code section end--------------------------------
		fnAddLog(EndingFN(0) & "fnPreInstall" & DotFN(2)) 
		
		strLogTAB = strLogTABBak
		fnPreInstall = intFuncRetCode
End Function
'-------------------------------------------------------------------------------------------
'! Function for all actions regarding postinstall actions like remove unpacked archives after installation. 
'!
'! @return 0,1
Function fnPostInstall()
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		intFuncRetCode = 0
		
		fnAddLog(StartingFN(0) & "fnPostInstall" & DotFN(2))
			' Put here your code-------------------------------------
			
			
			
			
			' Custom code section end--------------------------------
		ProcessesWindowRegClean(strPKGName) 				'Automatic registry cleanup after XStrike. 
		fnAddLog(EndingFN(0) & "fnPostInstall" & DotFN(2)) 
		
		strLogTAB = strLogTABBak
		fnPostInstall = intFuncRetCode
End Function
'-------------------------------------------------------------------------------------------
'! Function for all actions regarding preuninstall actions like backup files before installation. 
'!
'! @return 0,1
Function fnPreUninstall()
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		intFuncRetCode = 0 
		 
		fnAddLog(StartingFN(0) & "fnPreUninstall" & DotFN(2))
			' Put here your code-------------------------------------
			
			
			
			
			' Custom code section end--------------------------------
		fnAddLog(EndingFN(0) & "fnPreUninstall" & DotFN(2)) 
		
		strLogTAB = strLogTABBak
		fnPreUninstall = intFuncRetCode
End Function
'-------------------------------------------------------------------------------------------
'! Function for all actions regarding postuninstall actions like removing empty folders. 
'!
'! @return 0,1
Function fnPostUninstall()
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		intFuncRetCode = 0
		
		fnAddLog(StartingFN(0) & "fnPostUninstall" & DotFN(2))
			' Put here your code-------------------------------------
			
			
			
			
			' Custom code section end--------------------------------
		fnAddLog(EndingFN(0) & "fnPostUninstall" & DotFN(2)) 
		
		strLogTAB = strLogTABBak
		fnPostUninstall = intFuncRetCode
End Function
'-------------------------------------------------------------------------------------------
' Main funtions ----------------------------------------------------------------------------
'! Main installation function. 
'!
'! @return 0,1
Function Install()
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
		
		fnAddLog(StartingFN(0) & "Install" & DotFN(2))
			intFuncRetCode = fnPrereqsInstall()						'fnPrereqsInstall function call
            call fnRetCodeChecker(intFuncRetCode, "0,3010") 
			
			intFuncRetCode = fnPreInstall()							'fnPreInstall function call
            call fnRetCodeChecker(intFuncRetCode, "0,3010") 
			
			intFuncRetCode = 0
			' Put here your code, use intFuncRetCode for action exit codes-------------------------------------
			
			
			
			
			
			
			' Custom code section end--------------------------------
            call fnRetCodeChecker(intFuncRetCode, "0,3010") 
			
			intFuncRetCode = fnPostInstall()							'fnPostInstall function call
            call fnRetCodeChecker(intFuncRetCode, "0,3010")
		fnAddLog(EndingFN(0) & "Install" & DotFN(2))  
		
			strLogTAB = strLogTABBak
			fnFinalize(intFuncRetCode)								'fnFinalize function call - exiting script
End Function
'-------------------------------------------------------------------------------------------
'! Main uninstallation function. 
'!
'! @return 0,1
Function Uninstall()
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
			
		fnAddLog(StartingFN(0) & "Uninstall" & DotFN(2))
			intFuncRetCode = fnPreUninstall()						'fnPreUninstall function call
            call fnRetCodeChecker(intFuncRetCode, "0,3010") 
			
			intFuncRetCode = 0
			' Put here your code, use intFuncRetCode for action exit codes-------------------------------------
			
			
			
			
			' Custom code section end--------------------------------
            call fnRetCodeChecker(intFuncRetCode, "0,3010") 
			
			intFuncRetCode = fnPostUninstall()						'fnPostUninstall function call
            call fnRetCodeChecker(intFuncRetCode, "0,3010")
		fnAddLog(EndingFN(0) & "Uninstall" & DotFN(2))  
		
			strLogTAB = strLogTABBak
			fnFinalize(intFuncRetCode)								'fnFinalize function call - exiting script
End Function
'-------------------------------------------------------------------------------------------
'! Roolback function in case of failed installation. 
'!
'! @return N/A.
Function fnRollback()
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
			
		fnAddLog(StartingFN(0) & "fnRollback" & DotFN(2))
			' Put here your code, use intFuncRetCode for action exit codes-------------------------------------
		
		
		
			' Custom code section end--------------------------------
    
		          fnAddLog("SETUP INFO: removing TAG.") 
                  If fnIsExistTag(strPKGName) = 0 Then
					call fnRemoveTag(strPKGName)
				  End If
				  
				  'call Clear5thStr()
    
		fnAddLog(EndingFN(0) & "fnRollback" & DotFN(2)) 
		
			strLogTAB = strLogTABBak
End Function
' Template functions, not to edit ----------------------------------------------------------
'! Load all include files with main template. 
'!
'! @param  strFile          File Name (*.vbs) with full path.
'!
'! @return N/A.
Function fnInclude(strFile)
	Dim objTextFile    : Set objTextFile = objFSO.OpenTextFile(Replace(WScript.ScriptFullName,WScript.ScriptName,"") & strFile, 1, false)
	Dim myFunctionsStr : myFunctionsStr = objTextFile.ReadAll

		objTextFile.Close

	Set objTextFile = Nothing

		ExecuteGlobal myFunctionsStr
End Function
'-------------------------------------------------------------------------------------------