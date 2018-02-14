#cs ----------------------------------------------------------------------------

	Functions list:
		PreRemoveProducts()
		PreInstall()
		PostInstall()
		PreUninstall()
		PostUninstall()
		Uninstall()
		Install($instance = 0)
		Rollback()
		FiveStrike()
		Finalize($Last_RC)
		Initialize()
		DoAction()

#ce ----------------------------------------------------------------------------


#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Icon=Include\app.ico
#AutoIt3Wrapper_Res_Field=Internal Name|Wrapper for ULA package
#AutoIt3Wrapper_Res_Field=Title|Wrapper for ULA package
#AutoIt3Wrapper_Res_Field=Product Name|Wrapper for ULA package
#AutoIt3Wrapper_Res_Field=Product Version|0.9.0
#AutoIt3Wrapper_Res_Field=Build Number|20170707
#AutoIt3Wrapper_UseUpx=N
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

AutoItSetOption('TrayIconHide', '1')
AutoItSetOption('MustDeclareVars', '1')

#include <File.au3>
#include <WinAPIFiles.au3>

;~ Custom exit codes
Global Const $ExitCode_Five_Strike 					= -1000  
Global Const $ExitCode_Server 						= -1001  
Global Const $ExitCode_Newer_Version 				= -1002  
Global Const $ExitCode_CoreAppNotInstalled 			= -1003  
Global Const $ExitCode_PrereqNotInstalled	 		= -1004
Global Const $ExitCode_InvalidInstallationDetected 	= -1005
Global Const $ExitCode_NotEnoughFreeDiskSpace 		= -1006

Global $ExitCodeArray 					= StringSplit('-1000,-1001,-1002,-1003,-1004,-1005,-1006', ',')
Global $ExitCodeArrayMsgFail 		    = StringSplit('-1001,-1002,-1003,-1004,-1005,-1006', ',')

;~ Five strike
Global Const $Five_Strike_Registry_Root = 'HKEY_LOCAL_MACHINE\SOFTWARE\IBM\Packages'

;~ Environment variables
Global Const $ProgramFiles 				= EnvGet('ProgramFiles(x86)')					;C:\Program Files (x86)
Global Const $ProgramFilesX64 			= EnvGet('ProgramW6432')						;C:\Program Files
Global Const $C_Drive 					= StringLeft (@WindowsDir, 2)					;C:


;~ ######################################################################################
;~ # START Feel free to edit everything in here
;~ ######################################################################################
Global Const $IBM_AppTitle 				= ''											;~ Name provided by customer
Global Const $Friendly_AppTitle 		= ''											;~ Application name to be used in messages displayed to user
Global Const $Friendly_AppVersion 		= ''											;~ Application version to be used in messages displayed to user

Global Const $Substitute_ARP 			= True											;~ Hides original entry in Add/Remove Programs and adds one handled by script
Global Const $Use_Five_Strike 			= True											;~ Use 5-strike rule with defferal window
Global Const $Remove_Self 				= True											;~ Uninstall self (by GUID) before starting installation. Highly recommended, set to False only if there is a real need
Global Const $NeedReboot				= False											;~ Disables or enables Reboot warning for 5th Strike
Global Const $NeedLogOff				= False											;~ Disables or enables Log off warning for 5th Strike

Global Const $App_GUID 					= ''											;~ Application Product Code (main entry in registry)
Global Const $App_Version			 	= ''											;~ Application Version
Global Const $x64_product 				= False											;~ Application bitness

;~ Filenames and parameters. Do not include path in it
Global Const $MSI_Filename				= ''
Global Const $MST_Filename				= ''
Global Const $MSP_Filename				= ''
Global Const $MSIEXEC_Param				= ''
Global Const $EXE_Filename				= ''
Global Const $EXE_UPreParam				= ''
Global Const $EXE_PreParam				= ''
Global Const $EXE_Param					= ''
Global Const $EXE_Param_Uninstall		= ''

;~ Source folder of installer
Global Const $SRC_Dir					= @ScriptDir & '\_Source'						;~ Can be used as path for the files in script

Global Const $Process_To_Kill_On_Install= ''											;~ Semicolon (;) separated list
Global Const $Process_To_Kill_On_Remove = ''											;~ Semicolon (;) separated list
Global Const $Friendly_Process_Name 	= ''


Global Const $Products_To_Remove 		= ''											;~ Semicolon (;) separated list of products' GUIDs to be removed before installation. Append '_x64' after GUID referring to x64 product

;~ Message section
Local  $MSGStrike 						= ''											;~ DO NOT EDIT
Local  $Cancel_info 					= ''											;~ DO NOT EDIT
Global $HowManyStrikes 					= 5												;~ How many strikes will be used for 5ThStrike window (default = 5), can be change as client request.
Global $Message_Before					= ''
Global $Message_Before_Forced			= ''
	   $MSGStrike 						= RegRead($Five_Strike_Registry_Root & '\' & $IBM_Apptitle, 'strikecount')
	   $Cancel_info 					= $HowManyStrikes - $MSGStrike


Global $ConstInf01						= 'D "ULA IT is attempting to install the ' & $Friendly_AppTitle & ' ' & $Friendly_AppVersion & ' application.'
Global $ConstInf02						= 'Please close all ' & $Friendly_Process_Name & ' sessions. Selecting `Install Now` will close all ' & $Friendly_Process_Name & ' sessions AUTOMATICALLY.'
Global $ConstInf03						= @CRLF & @CRLF & 'Click `Cancel` to postpone this installation.' & @CRLF & @CRLF & 'You can postpone ' & $Cancel_info & ' more time(s) before this installation will be required.'
Global $ConstInf04						= '" "Install Now" "Cancel" 60 0 5 ' & $IBM_AppTitle
Global $ConstInf05						= 'After the installation is complete your computer will require a restart.'
Global $ConstInf06						= 'After the installation is complete, log off will be required.'

Global $ConstForceInf01					= 'D "WARNING!!! - The install of ' & $Friendly_AppTitle & ' ' & $Friendly_AppVersion & ' is now required to proceed as you have reached the postpone limit. '
Global $ConstForceInf02					= 'Do not disrupt the install. '
Global $ConstForceInf03					= @CRLF & @CRLF & 'Click `OK` to proceed. The install should take less than 10 minutes.'
Global $ConstForceInf04					= '" "OK" "" 10 0 0 ' & $IBM_AppTitle
Global $ConstForceInf05					= 'After the installation is complete your computer will require a restart.'
Global $ConstForceInf06					= 'After the installation is complete, log off will be required.'
	   $Message_Before_Forced 			= $ConstForceInf01
	   $Message_Before_Forced 			= $Message_Before_Forced & $ConstForceInf02


	If $Process_To_Kill_On_Install = '' Then
		$Message_Before = $ConstInf01
	Else
		$Message_Before = $ConstInf01 & @CRLF & $ConstInf02
	EndIf

	If $NeedReboot = True Or $NeedLogOff = True Then
		$Message_Before = $Message_Before & $ConstInf03
		$Message_Before_Forced = $Message_Before_Forced & @CRLF & @CRLF & $ConstForceInf03

			If $NeedReboot = True And $NeedLogOff = False Then
				$Message_Before = $Message_Before & @CRLF & $ConstInf05 & $ConstInf04
				$Message_Before_Forced = $Message_Before_Forced & @CRLF & $ConstForceInf05 & $ConstForceInf04
			ElseIf $NeedReboot = False And $NeedLogOff = True Then
				$Message_Before = $Message_Before & @CRLF & $ConstInf06 & $ConstInf04
				$Message_Before_Forced = $Message_Before_Forced & @CRLF & $ConstForceInf06 & $ConstForceInf04
			EndIf
	Else
		$Message_Before = $Message_Before & $ConstInf03 & $ConstInf04
		$Message_Before_Forced = $Message_Before_Forced & @CRLF & @CRLF & $ConstForceInf03 & $ConstForceInf04
	EndIf





Global $Message_During 					= $Friendly_AppTitle & ' ' & $Friendly_AppVersion & ' is installing...'
Global $Message_After 					= 'B "The installation of ' & $Friendly_AppTitle & ' ' & $Friendly_AppVersion & ' has completed successfully." "OK" "" 0 0'
Global $ConstInfA1 						= 'B "The installation of ' & $Friendly_AppTitle & ' ' & $Friendly_AppVersion & ' has completed successfully. ' & @CRLF & 'The system requires a restart for the application to function correctly." "OK" "" 0 0'
Global $ConstInfA2 						= 'B "The installation of ' & $Friendly_AppTitle & ' ' & $Friendly_AppVersion & ' has completed successfully. ' & @CRLF & 'The system requires a log off for the application to function correctly." "OK" "" 0 0'
Global $Message_After_Fail 				= 'B "The installation of ' & $Friendly_AppTitle & ' ' & $Friendly_AppVersion & ' failed.'
		$Message_After_Fail				= $Message_After_Fail & @CRLF & 'Please contact the Help Desk for further assistance." "OK" "" 0 0'



		If $NeedReboot = True And $NeedLogOff = False Then
			$Message_After = $ConstInfA1
		ElseIf $NeedReboot = False And $NeedLogOff = True Then
			$Message_After = $ConstInfA2
		EndIf


Global $Is_Silent 						= False											;~ DO NOT EDIT
Global $Is_FSilent 						= False											;~ DO NOT EDIT
Global $Uninstall_Mode 					= False											;~ DO NOT EDIT
Global $ScrMode 						= ''											;~ String for logs - Installation Uninstallation
Global $isInstalledGVar					= 0

;~ Logs section
Global $Log_File
Global Const $Log_Dir 					= $ProgramFiles & '\Logs\' & $IBM_AppTitle & '\'
Global $Installation_Log 				= $Log_Dir & $IBM_AppTitle & '_WRAP_INSTALL.log'
Global $Uninstallation_Log 				= $Log_Dir & $IBM_AppTitle & '_WRAP_UNINSTALL.log'
Global $Msi_Installation_Log 			= $Log_Dir & $IBM_AppTitle & '_MSI_INSTALL.log'
Global $Msi_Uninstallation_Log 			= $Log_Dir & $IBM_AppTitle & '_MSI_UNINSTALL.log'
Global $logTAB							= ''
Global $logSeparator 					= '------'

;~ Tag
Global Const $Tag_File 					= $ProgramFiles & '\Logs\' & $IBM_AppTitle & '.tag'

;~ IBMSRC
Global $IBMSRC_Dir 						= @WindowsDir & '\IBMSRC\' & $IBM_Apptitle

;~ Last_RC and Result
Global $Last_RC 						= 0
Global $Result 							= 1

;~ ######################################################################################
;~ Custom code to execute before removal of conflicting product(s) start goes here
;~ Return non-zero to indicate error and stop rest of script
;~ ######################################################################################
Func PreRemoveProducts()
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "PreRemoveProducts"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function')
			;--------------------------------------------------------------------------------------------------------------------------









			;--------------------------------------------------------------------------------------------------------------------------
		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
		$logTAB = $logTABBak
	Return $FuncRetCode
EndFunc

;~ ######################################################################################
;~ Custom pre-installation code goes here
;~ Return non-zero to indicate error and stop rest of script
;~ ######################################################################################
Func PreInstall()
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "PreInstall"
		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function')
			;--------------------------------------------------------------------------------------------------------------------------









			;--------------------------------------------------------------------------------------------------------------------------
		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
		$logTAB = $logTABBak
	Return $FuncRetCode
EndFunc

;~ ######################################################################################
;~ Custom post-installation code goes here
;~ Return non-zero to indicate error and stop rest of script
;~ ######################################################################################
Func PostInstall()
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "PostInstall"
		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function')
			;--------------------------------------------------------------------------------------------------------------------------








			;--------------------------------------------------------------------------------------------------------------------------
		CreateTag($IBM_AppTitle)

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
		$logTAB = $logTABBak
	Return $FuncRetCode
EndFunc


;~ ######################################################################################
;~ Custom pre-uninstallation code goes here
;~ Return non-zero to indicate error and stop rest of script
;~ ######################################################################################
Func PreUninstall()
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "PreUninstall"
		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function')
			;--------------------------------------------------------------------------------------------------------------------------









			;--------------------------------------------------------------------------------------------------------------------------
		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
		$logTAB = $logTABBak
	Return $FuncRetCode
EndFunc

;~ ######################################################################################
;~ Custom post-uninstallation code goes here
;~ Return non-zero to indicate error and stop rest of script
;~ ######################################################################################
Func PostUninstall()
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "PostUninstall"
		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function')
			;--------------------------------------------------------------------------------------------------------------------------









			;--------------------------------------------------------------------------------------------------------------------------
		RemoveTag($IBM_AppTitle)

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
		$logTAB = $logTABBak
	Return $FuncRetCode
EndFunc

;~ ######################################################################################
;~ # END Feel free to edit everything in here
;~ ######################################################################################


;~ ######################################################################################
;~ # Do NOT modify anything in here unless you really know what you're doing
;~ ######################################################################################

		$Last_RC = Initialize()
		If $Last_RC <> 0 Then Finalize($Last_RC)
		DoAction()
		;Finalize($Last_RC)


;------------------------------------------------------------------------------------------------------------------------
Func Uninstall()
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $tmp, $msiexec_cmdline
	Local $FuncRetCode					= 0
	Local $FnName						= "Uninstall"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

		KillProcesses($Log_File, $Process_To_Kill_On_Remove) 							;Kill processes

		$FuncRetCode = PreUninstall() 													;PreUninstall
			RetCodeChecker($FuncRetCode, "0,3010")
			$FuncRetCode = 0
			;--------------------------------------------------------------------------------------------------------------------------









			;--------------------------------------------------------------------------------------------------------------------------
			RetCodeChecker($FuncRetCode, "0,3010")

		$FuncRetCode = PostUninstall()													;PostUninstall
			RetCodeChecker($FuncRetCode, "0,3010")

		$FuncRetCode = ClearARP()														;Clear ARP custom entry
			RetCodeChecker($FuncRetCode, "0,3010")

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak

	Finalize($FuncRetCode)
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
Func Install($instance = 0)
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $tmp, $msiexec_cmdline, $message
	Local $FuncRetCode					= 0
	Local $FnName						= "Install"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

		If $Use_Five_Strike = True Then FiveStrike() 									;Five strike

		If Not $Is_FSilent = True Then 													;Progress bar
			BannerDisplay($Message_During)
		EndIf

		KillProcesses($Log_File, $Process_To_Kill_On_Install) 							;Kill processes

		$FuncRetCode = PreRemoveProducts() 												;Pre Remove previous versions
			RetCodeChecker($FuncRetCode, "0,3010")

		$FuncRetCode = RemoveProducts($Products_To_Remove) 								;Remove previous versions
			RetCodeChecker($FuncRetCode, "0,3010")

		$FuncRetCode = PreInstall() 													;Preinstall
			RetCodeChecker($FuncRetCode, "0,3010")
			$FuncRetCode = 0
			;--------------------------------------------------------------------------------------------------------------------------









			;--------------------------------------------------------------------------------------------------------------------------
			RetCodeChecker($FuncRetCode, "0,3010")

		$FuncRetCode = PostInstall() 													;PostInstall
			RetCodeChecker($FuncRetCode, "0,3010")

		If $Substitute_ARP = True Then AdjustARP() 										;Substitute original/vendor ARP entry

		;If Not $Is_FSilent = True Then 													;Completion message
		;	ShellExecute(@WindowsDir & '\UlaDialog3.exe', $Message_After, @ScriptDir, '', @SW_SHOWDEFAULT)
		;	BannerRemove()
		;EndIf

		Sleep(300)

		$FuncRetCode = 0 																;Required for failed message

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
	Finalize($FuncRetCode)
EndFunc
;------------------------------------------------------------------------------------------------------------------------

;------------------------------------------------------------------------------------------------------------------------
Func Rollback()
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "Rollback"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			RemoveTag($IBM_AppTitle)

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
Func FiveStrike()
	Local $StrikeCount, $Message, $RegPath, $Cancel_Option = True
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "FiveStrike"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			$RegPath = $Five_Strike_Registry_Root & '\' & $IBM_Apptitle
			$StrikeCount = RegRead($RegPath, 'strikecount') 							;Check five strike count

			If $StrikeCount = '' Then
				WriteToLog($Log_File, $logTAB & '>INFO: ' & $RegPath & ' do not exist, adding Reg key.')
				$StrikeCount = 0
				RegWrite($RegPath, 'strikecount', 'REG_SZ', $StrikeCount)
			EndIf

			If $StrikeCount = $HowManyStrikes Then $Cancel_Option = False

			If Not $Is_Silent = True And Not $Is_FSilent = True Then 					;Show five strike message
				If $Cancel_Option = False Then
					$Message = $Message_Before_Forced
				Else
					$Message = $Message_Before
				EndIf

				WriteToLog($Log_File, $logTAB & '>INFO: Strike count: ' & $StrikeCount)
				$Last_RC = ShellExecuteWait(@WindowsDir & '\UlaDialog3.exe', $Message, @ScriptDir, '', @SW_SHOWDEFAULT)

				If $Last_RC = 0 And @error = 0 Or $Last_RC = 3010 Then 					;Check return code
					WriteToLog($Log_File, $logTAB & '>INFO: User Selected "Install Now". Installation will start now.')
				Else
					WriteToLog($Log_File, $logTAB & '>INFO ERROR: User selected "Cancel or timeout occurred" ' & $StrikeCount & ' time(s). Install will exit now. Exit code ' & $Last_RC & '.')
					$StrikeCount = $StrikeCount + 1

					RegWrite($RegPath, 'strikecount', 'REG_SZ', $StrikeCount)

					WriteToLog($Log_File, $logTAB & '>INFO: Increasing strike count to ' & $StrikeCount & '.')
					Finalize($ExitCode_Five_Strike)
				EndIf
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
;~ Remove Uladialog3.exe, close logs, exit with specified return code. This is the ONLY function in script that calls Exit() directly and it should remain so.
Func Finalize($VarLast_RC)
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $i = 1
	Local $FuncRetCode					= 0
	Local $FailWindow					= 1
	Local $FnName						= "Finalize"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			If $VarLast_RC = 0 Or $VarLast_RC = 3010 Then
				$Result = 0
			EndIf

			For $ExitCodeElement In $ExitCodeArray 										;Check custom return codes
				If $ExitCodeArray[$i] = $VarLast_RC Then
					$Result = 0
					ExitLoop
				EndIf

				If $i <> $ExitCodeArray[0] Then
					$i = $i + 1
				EndIf
			Next
            
            $i = 0
            For $ExitCodeElementA In $ExitCodeArrayMsgFail 								;Check custom return codes
				If $ExitCodeArrayMsgFail[$i] = $VarLast_RC Then
					$FailWindow = 0
					ExitLoop
				EndIf

				If $i <> $ExitCodeArrayMsgFail[0] Then
					$i = $i + 1
				EndIf
			Next

            If $VarLast_RC <> $ExitCode_Five_Strike Then
				If Not $Is_FSilent = True And $Uninstall_Mode = False And ($FailWindow = 0 OR $Result = 1) Then	;Show fail message
					ShellExecute(@WindowsDir & '\UlaDialog3.exe', $Message_After_Fail, @ScriptDir, '', @SW_SHOWDEFAULT)
						WriteToLog($Log_File, $logTAB & '>INFO: ' & @WindowsDir & '\UlaDialog3.exe -> ' & $Message_After_Fail)
				EndIf

				If $Is_FSilent = False And $Result = 0 And $Uninstall_Mode = False And $isInstalledGVar = 0 And $FailWindow = 1 Then
					ShellExecute(@WindowsDir & '\UlaDialog3.exe', $Message_After, @ScriptDir, '', @SW_SHOWDEFAULT)
						WriteToLog($Log_File, $logTAB & '>INFO: ' & @WindowsDir & '\UlaDialog3.exe -> ' & $Message_After)
				EndIf
			EndIf

			If $Result = 1 Then
				Rollback()
			Else
				If $Uninstall_Mode = True Then
					RemoveFiveStrike()
				EndIf
			EndIf

			If FileExists(@WindowsDir & '\UlaDialog3.exe') Then 						;Remove UlaDialog3.exe
				FileDelete(@WindowsDir & '\UlaDialog3.exe')
				WriteToLog($Log_File, $logTAB & '>INFO: UlaDialog3.exe removed from ' & @WindowsDir & '.')
			EndIf

			WriteToLog($Log_File, $logTAB & '>INFO: Return code from last operation: ' & $VarLast_RC & '.')
			WriteToLog($Log_File, $logTAB & '>************* ' & $IBM_AppTitle & ' STOPPED *************')

			BannerRemove()

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Exit($VarLast_RC)
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
;~ Parse commandline, create folder structure for logs (if needed), initialize logs, install UlaDialog3.exe
Func Initialize()
	Local $isInstalledVar = 0
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "Initialize"

			If IsServerOS() = True Then Exit($ExitCode_Server) 									;Check OS type

			If StringInStr($CmdLineRaw, '/remove', 0) <> 0 Then $Uninstall_Mode = True 			;Check uninstall mode

			If $Uninstall_Mode = True Then 														;Create log dir and log
				$Log_File 	= $Uninstallation_Log
				$ScrMode	= 'Uninstall'

				If Not FileExists($Log_Dir) Then
					DirCreate($Log_Dir)
					WriteToLog($Log_File, $logTAB & '>INFO: ************* ' & $IBM_AppTitle & ' STARTED *************')
					WriteToLog($Log_File, $logTAB & '>INFO: Log folder (' & $Log_Dir & ') not found, created.')
					WriteToLog($Log_File, $logTAB & '>INFO: Executed with commandline: ' & $IBM_AppTitle & '.exe ' & $CmdLineRaw)
				Else
					WriteToLog($Log_File, $logTAB & '>INFO: ************* ' & $IBM_AppTitle & ' STARTED *************')
					WriteToLog($Log_File, $logTAB & '>INFO: Executed with commandline: ' & $IBM_AppTitle & '.exe ' & $CmdLineRaw)
				EndIf
				WriteToLog($Log_File, $logTAB & '>INFO: /remove selected, running uninstallation.')

				$isInstalledVar = isPKGInstalled($App_GUID, $App_Version, $IBM_AppTitle)
				If $isInstalledVar == 1 Then 					;Checking if application is installed
					WriteToLog($Log_File, $logTAB & '>INFO: ' & $IBM_AppTitle & ' is installed, going to uninstall.')
				ElseIf $isInstalledVar == 2 Then
					WriteToLog($Log_File, $logTAB & '>INFO: ' & $IBM_AppTitle & ' is not installed, Vendor installed, going to uninstall.')
				Else
					WriteToLog($Log_File, $logTAB & '>INFO: ' & $IBM_AppTitle & ' is not installed, nothing to uninstall.')
					$FuncRetCode = 0
					Finalize($FuncRetCode)
				EndIf
			Else
				$Log_File 	= $Installation_Log
				$ScrMode	= 'Install'

				If Not FileExists($Log_Dir) Then
					DirCreate($Log_Dir)
					WriteToLog($Log_File, $logTAB & '>INFO: ************* ' & $IBM_AppTitle & ' STARTED *************')
					WriteToLog($Log_File, $logTAB & '>INFO: Log folder (' & $Log_Dir & ') not found, created.')
					WriteToLog($Log_File, $logTAB & '>INFO: Executed with commandline: ' & $IBM_AppTitle & '.exe ' & $CmdLineRaw)
				Else
					WriteToLog($Log_File, $logTAB & '>INFO: ************* ' & $IBM_AppTitle & ' STARTED *************')
					WriteToLog($Log_File, $logTAB & '>INFO: Executed with commandline: ' & $IBM_AppTitle & '.exe ' & $CmdLineRaw)
				EndIf

				$isInstalledVar = isPKGInstalled($App_GUID, $App_Version, $IBM_AppTitle)
				If $isInstalledVar == 1 Then 			;Checking if application is installed
					WriteToLog($Log_File, $logTAB & '>INFO: ' & $IBM_AppTitle & ' is installed, nothing to install.')
					$FuncRetCode = 0
					$isInstalledGVar = $isInstalledVar
					Finalize($FuncRetCode)
				ElseIf $isInstalledVar == 2 Then
					WriteToLog($Log_File, $logTAB & '>INFO: ' & $IBM_AppTitle & ' is not installed, vendor PKG present going to install.')
				Else
					WriteToLog($Log_File, $logTAB & '>INFO: ' & $IBM_AppTitle & ' is not installed, going to install.')
				EndIf

				FileInstall('.\Include\UlaDialog3.exe', @WindowsDir & '\UlaDialog3.exe', 1) 	;Install UlaDialog3.exe
				WriteToLog($Log_File, $logTAB & '>INFO: UlaDialog3.exe installed to ' & @WindowsDir & '.')

				If StringInStr($CmdLineRaw, '/silent', 0) <> 0 Then								;Check 'silent' switch
					$Is_Silent = True
					WriteToLog($Log_File, $logTAB & '>INFO: /silent selected, no 5th Strike will be displayed.')
				EndIf

				If StringInStr($CmdLineRaw, '/fsilent', 0) <> 0 Then							;Check 'fsilent' switch
					$Is_FSilent = True
					WriteToLog($Log_File, $logTAB & '>INFO: /fsilent selected, no prompts and no 5th Strike will be displayed.')
				EndIf
			EndIf

		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
Func DoAction()
	If $Uninstall_Mode = True Then
		Uninstall()
	Else
		Install()
	EndIf
EndFunc
;------------------------------------------------------------------------------------------------------------------------



#include ".\Include\functions.au3"