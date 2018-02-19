#cs ----------------------------------------------------------------------------

	Functions list:
		KillProcesses($Log_File, $Process_List)
		RemoveProducts($Products_To_Remove)
		RemoveProduct($GUID)
		SelfDelete()
		AdjustARP()
		ClearARP()
		WriteToLog($Log_Filename, $Message)
		IsServerOS()
		CreateTag($VarPKGName='')
		ClearREGISTRY($AGUID = $App_GUID, $RegPath = '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\')
		RemoveTag($VarPKGName='')
		CopyFile($Src_Dir, $Dest_Dir, $File_Name)
		RemoveFile($Src_Dir, $File_Name)
		RemoveDirectory($Dir_Path, $willFail)
		RemoveDirectoryIfEmpty($Dir_Path)
		InstallEXE($Src_Dir, $EXE_FileName, $EXE_Param, $EXE_PreParam, $ExicCodes = '0,3010')
		UninstallEXE($Src_Dir, $EXE_FileName, $EXE_Param, $EXE_UPreParam, $ExicCodes = '0,3010')
		CheckFreeDiskSpace($MB_Size)
		RemoveFiveStrike()
		BannerDisplay($banner_text)
		BannerRemove()
		RetCodeChecker($FuRetCode, $ProperRCList)
		MSIInstall($VarSRC_Dir, $VarMSI_Filename, $VarMST_Filename, $VarMSP_FIlename = '', $VarMSIEXEC_Param = '', $ExicCodes = '0,3010')
		MSIUninstall($VarGUID, $VarMSIEXEC_Param = '', $ExicCodes = '0,3010')
		isPKGInstalled($VarGUID, $VarVersion, $VarTAGName = '') ; 0 - nothing is installed, 1 - IBM PKG installed, 2 - Vendor installed
		PrevPKGUninstall($VarPKGName, $VarGUID, $VarVersion, $VarInstDir = '', $VarVendorName = '', $ExicCodes = '0,3010')
		SetEnvVar($VarName, $VarValue)
		RegHideARP($GUID)
		SelfCopySrc()
		isCMDProcessRunning($proc, $procCMD = '', $strComputer = '.')
		SwitchService($service_name)
		SetNoModifyARP($AGUID)
		removeAllPrevBuilds()

#ce ----------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
;~ Kills processes specified in $process_list by executable name (/IM switch). Rest of parameters if force (/F) and kill child processes (/T)
Func KillProcesses($Log_File, $Process_List)
	Local $tmp, $i
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "KillProcesses"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			If $Process_List = '' Then
				WriteToLog($Log_File, $logTAB & '>INFO: Process list is empty, nothing to kill.')
				WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
				$logTAB = $logTABBak
				Return 0
			EndIf

			$tmp = StringSplit($Process_List, ';')
			For $i = 1 To $tmp[0]
				WriteToLog($Log_File, $logTAB & '>INFO: Executing "' & @SystemDir & '\taskkill.exe /F /IM ' & $tmp[$i] & ' /T')
				ShellExecuteWait(@SystemDir & '\taskkill.exe', '/F /IM "' & $tmp[$i] & '" /T', @ScriptDir, '', @SW_HIDE)
					If @error <> 0 Then
						WriteToLog($Log_File, $logTAB & '>INFO: Error durring kill process.')
						$FuncRetCode = 1
					EndIf
			Next

			If $FuncRetCode <> 0 Then
				WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function failed.')
				$logTAB = $logTABBak
				Finalize(1)
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
;~ Returns RC when there was a product to remove and that uninstall ended with error. Otherwise, returns 0
;~ Returns return code when there was a product to remove and that uninstall ended with error. Otherwise, returns 0
Func RemoveProducts($VarProducts_To_Remove)
	Local $tmp, $i
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "RemoveProducts"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			If $Remove_Self = True AND $MSI_Filename <> '' Then
				$tmp = $App_GUID
				If $x64_product = True Then $tmp = $tmp & '_x64'

				$FuncRetCode = RemoveProduct($tmp)

				If $FuncRetCode <> 0 And $FuncRetCode <> 3010 Then
					WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function. RemoveProduct function ended with RC <> 0,3010')
					$logTAB = $logTABBak
					Return $FuncRetCode
				EndIf
			EndIf

			If $VarProducts_To_Remove = '' Then
				WriteToLog($Log_File, $logTAB & '>INFO: Products list is empty, nothing to remove')
				WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
				$logTAB = $logTABBak
				Return 0
			EndIf

			$tmp = StringSplit($VarProducts_To_Remove, ';')
			For $i = 1 To $tmp[0]
				$FuncRetCode = RemoveProduct($tmp[$i])
				If $FuncRetCode <> 0 And $FuncRetCode <> 3010 Then
					WriteToLog($Installation_Log, $logTAB & '>INFO: One or more previous versions encountered an error during uninstall - please check the uninstall log. Exiting now.')
					WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
					$logTAB = $logTABBak
					Return $FuncRetCode
				EndIf
			Next

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
;~ Removes MSI product, speficied by GUID. Returns 1 in case of any error, 0 otherwise.
Func RemoveProduct($GUID)
	Local $MsiexecCommandline, $Last_RC, $temp, $notInstalled = False
	Local $x64 = ''
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "RemoveProduct"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			;If StringRight($GUID, 4) = '_x64' Then ;Set bitness and GUID
				$x64 = '64'
			;	$GUID = StringLeft($GUID, StringLen($GUID) - 4)
			;EndIf

			;Detect software
			$temp = RegRead('HKEY_LOCAL_MACHINE' & $x64 & '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $GUID, 'DisplayName')

			If $temp = '' Then
				WriteToLog($Log_File, $logTAB & '>INFO: HKEY_LOCAL_MACHINE' & $x64 & '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $GUID & '\DisplayName, not found.')
					$temp = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $GUID, 'DisplayName')
					If $temp = '' Then
						WriteToLog($Log_File, $logTAB & '>INFO: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $GUID & '\DisplayName, not found.')
						$notInstalled = True
					Else
						$notInstalled = False
					EndIf
			Else
				$notInstalled = False
			EndIf

			If $notInstalled = True Then
				WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
				$logTAB = $logTABBak
				Return 0
			Else
				$MsiexecCommandline = '/x ' & $GUID & ' /l*v+ "' & $Log_Dir & $GUID & '.MSI_UNINSTALL.log" /qn REBOOT=ReallySuppress'
				WriteToLog($Log_File, $logTAB & '>INFO: ' & $MsiexecCommandline & ' executing.')
				$FuncRetCode = ShellExecuteWait(@SystemDir & '\msiexec.exe', $MsiexecCommandline, @ScriptDir, '', @SW_SHOWDEFAULT)

				;Check return code
				Switch $FuncRetCode
					Case 0
						WriteToLog($Log_File, $logTAB & '>INFO: ' & $GUID & ' was uninstalled successfully. Exit Code: ' & $FuncRetCode)
					Case 3010
						WriteToLog($Log_File, $logTAB & '>INFO: ' & $GUID & ' was uninstalled successfully. Exit Code: ' & $FuncRetCode & '. Rebooot required.')
						$FuncRetCode = 0
					Case 1605
						WriteToLog($Log_File, $logTAB & '>INFO: ' & $GUID & ' was NOT found on this machine, nothing to uninstall. Exiting uninstall...')
						$FuncRetCode = 0
					Case Else
						WriteToLog($Log_File, $logTAB & '>INFO: Uninstallation of ' & $GUID & ' has failed. Exit Code: ' & $FuncRetCode & '.')
						$FuncRetCode = 1
				EndSwitch
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
;~ Allows script to delete itself
Func SelfDelete()
	Local $temp
	Local $x64 = ''
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "SelfDelete"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			;Set temporary dir
			$temp = _TempFile(@TempDir)
			$temp = $temp & '.cmd'

			;Create cmd file
			FileWriteLine($temp, ':begin')
			FileWriteLine($temp, 'rd /S /Q "' & @ScriptDir & '"')
			FileWriteLine($temp, 'if exist "' & @ScriptDir & '" goto begin')
			FileWriteLine($temp, 'del /Q /F "' & $temp & '" > nul')

			;Delete script
			ShellExecute(@ComSpec, '/C "' & $temp & '"', @TempDir, '', @SW_HIDE)

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
;~ Hides original Add/Remove Programs entry, adds new one with Uninstall method pointing to script, makes copy of script in C:\Windows\IBMSRC\$IBM_Apptitle
Func AdjustARP()
	Local $Source_Key, $Target_Key, $temp, $TMPApp_GUID
	Local $x64 = ''
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "AdjustARP"
	Local $FuncRetCode					= 0

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			$x64 = '64'


			$temp = RegRead('HKEY_LOCAL_MACHINE' & $x64 & '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $App_GUID, 'DisplayName')

			If $temp = '' Then
				WriteToLog($Log_File, $logTAB & '>INFO: HKEY_LOCAL_MACHINE' & $x64 & '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $App_GUID & '\DisplayName, not found.')
					$temp = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $App_GUID, 'DisplayName')
					If $temp = '' Then
						WriteToLog($Log_File, $logTAB & '>INFO: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $App_GUID & '\DisplayName, not found.')
					Else
						$Source_Key = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $App_GUID
						$Target_Key = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $App_GUID & '_' & $IBM_Apptitle
					EndIf
			Else
				$Source_Key = 'HKEY_LOCAL_MACHINE' & $x64 & '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $App_GUID
				$Target_Key = 'HKEY_LOCAL_MACHINE' & $x64 & '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $App_GUID & '_' & $IBM_Apptitle
			EndIf

			;Create new entry
			RegWrite($Target_Key, 'DisplayName', 'REG_SZ', RegRead($Source_Key, 'DisplayName'))
			RegWrite($Target_Key, 'DisplayVersion', 'REG_SZ', RegRead($Source_Key, 'DisplayVersion'))
			RegWrite($Target_Key, 'Manufacturer', 'REG_SZ', RegRead($Source_Key, 'Manufacturer'))
			RegWrite($Target_Key, 'Publisher', 'REG_SZ', RegRead($Source_Key, 'Publisher'))
			RegWrite($Target_Key, 'NoModify', 'REG_DWORD', '1')
			RegWrite($Target_Key, 'NoRepair', 'REG_DWORD', '1')
			RegWrite($Target_Key, 'DisplayIcon', 'REG_SZ', @WindowsDir & '\IBMSRC\' & $IBM_Apptitle & '\' & @ScriptName)

			RegHideARP($App_GUID)

			FileCopy(@ScriptFullPath, @WindowsDir & '\IBMSRC\' & $IBM_Apptitle & '\' & @ScriptName, 9) ;Copy script to use during uninstallation

			RegWrite($Target_Key, 'UninstallString', 'REG_SZ', '"' & @WindowsDir & '\IBMSRC\' & $IBM_Apptitle & '\' & @ScriptName & '" /remove') ;Set Uninstallstring for new entry

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
;~ Removes modified ARP entry
Func ClearARP()
	Local $temp, $TMPApp_GUID
	Local $x64 = ''
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "ClearARP"
	Local $FuncRetCode					= 0

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			$x64 = '64'

			If $Substitute_ARP = True Then

				$temp = RegRead('HKEY_LOCAL_MACHINE' & $x64 & '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $App_GUID & '_' & $IBM_Apptitle, 'DisplayName')

				If $temp = '' Then
					WriteToLog($Log_File, $logTAB & '>INFO: HKEY_LOCAL_MACHINE' & $x64 & '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $App_GUID & '_' & $IBM_Apptitle & '\DisplayName, not found.')
						$temp = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $App_GUID & '_' & $IBM_Apptitle, 'DisplayName')
						If $temp = '' Then
							WriteToLog($Log_File, $logTAB & '>INFO: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $App_GUID & '_' & $IBM_Apptitle & '\DisplayName, not found.')
						Else
							$TMPApp_GUID = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $App_GUID & '_' & $IBM_Apptitle
						EndIf
				Else
					$TMPApp_GUID = 'HKEY_LOCAL_MACHINE' & $x64 & '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $App_GUID & '_' & $IBM_Apptitle
				EndIf

					RegDelete($TMPApp_GUID)


				If StringCompare(@ScriptFullPath, @WindowsDir & '\IBMSRC\' & $IBM_Apptitle & '\' & @ScriptName, 0) = 0 Then
					SelfDelete()
				Else
					ShellExecuteWait(@ComSpec, '/C rd /s /q "' & @WindowsDir & '\IBMSRC\' & $IBM_Apptitle & '"', @ScriptDir, '', @SW_HIDE)
				EndIf
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
;~ Removes modified ARP entry
Func ClearREGISTRY($AGUID, $RegPath)
	Local $temp, $TMPApp_GUID
	Local $x64 = ''
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "ClearREGISTRY"
	Local $FuncRetCode					= 0

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

				$x64 = '64'

					$TMPApp_GUID = 'HKEY_LOCAL_MACHINE' & $x64 & $RegPath & $AGUID
					WriteToLog($Log_File, $logTAB & '>INFO: trying to remove: ' & $TMPApp_GUID & '.')

					$FuncRetCode = RegDelete($TMPApp_GUID)

					If $FuncRetCode <> 2 Then
						WriteToLog($Log_File, $logTAB & '>INFO: removed: ' & $TMPApp_GUID & '.')
					Else
						$TMPApp_GUID = 'HKEY_LOCAL_MACHINE' & $RegPath & $AGUID
						WriteToLog($Log_File, $logTAB & '>INFO: trying to remove: ' & $TMPApp_GUID & '.')

						$FuncRetCode = RegDelete($TMPApp_GUID)

						If $FuncRetCode <> 2 Then
							WriteToLog($Log_File, $logTAB & '>INFO: removed: ' & $TMPApp_GUID & '.')
						Else
							WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
							$logTAB = $logTABBak
							Finalize(1)
						EndIf
					EndIf



		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
;~ Writes message, prepended with timestamp to specified log file. File will be created if it doesn't exist, folder structure will NOT be created
Func WriteToLog($Log_Filename, $Message)
	FileWriteLine($Log_Filename, @YEAR & '/' & @MON & '/' & @MDAY & ' ' & @HOUR & ':' & @MIN & ':' & @SEC & ' ' & $Message)
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
Func IsServerOS()
    Local $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Local $colSettings = $objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
	Local $versionShort, $versionLong, $productType

    For $objOperatingSystem In $colSettings
		$versionLong = $objOperatingSystem.Version
		Local $versionArray = StringSplit($versionLong, ".")

		$versionShort = $versionArray[1] & "." & $versionArray[2]
		$productType = $objOperatingSystem.ProductType
    Next

	If $productType <> 1 Then Return True

	Return False
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;TAG operations----------------------------------------------------------------------------------------------------------
Func CreateTag($VarPKGName='')
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "CreateTag"
	Local $FuncRetCode					= 0

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			If $VarPKGName == '' Then
				$VarPKGName = $Tag_File
			Else
				$VarPKGName = $ProgramFiles & '\Logs\' & $VarPKGName & '.tag'
			EndIf

			If Not FileExists($VarPKGName) Then
				WriteToLog($Log_File, $logTAB & '>INFO: TAG "' & $VarPKGName & '" does NOT exist.')

				$FuncRetCode = FileWrite($VarPKGName, '')

				If $FuncRetCode = 1 Then
					WriteToLog($Log_File, $logTAB & '>INFO: TAG "' & $VarPKGName & '" has been created.')
					$FuncRetCode = 0
				Else
					WriteToLog($Log_File, $logTAB & '>INFO ERROR: TAG "' & $VarPKGName & '" has NOT been created.')
					Finalize(1)
				EndIf

			Else
				WriteToLog($Log_File, $logTAB & '>INFO: TAG "' & $VarPKGName & '" exists.')
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
Func RemoveTag($VarPKGName='')
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "RemoveTag"
	Local $FuncRetCode					= 0

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			If $VarPKGName == '' Then
				$VarPKGName = $Tag_File
			Else
				$VarPKGName = $ProgramFiles & '\Logs\' & $VarPKGName & '.tag'
			EndIf

			If FileExists($VarPKGName) Then
				WriteToLog($Log_File, $logTAB & '>INFO: TAG "' & $VarPKGName & '" exists.')

				$FuncRetCode = FileDelete($VarPKGName)

				If $Uninstall_Mode = True Then
					If $FuncRetCode = 1 Then
						WriteToLog($Log_File, $logTAB & '>INFO: TAG "' & $VarPKGName & '" has been removed.')
						$FuncRetCode = 0
					Else
						WriteToLog($Log_File, $logTAB & '>INFO ERROR: TAG "' & $VarPKGName & '" has NOT been removed.')
						WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
						$logTAB = $logTABBak
						Finalize(1)
					EndIf
				Else
					If $FuncRetCode = 1 Then
						WriteToLog($Log_File, $logTAB & '>INFO: TAG "' & $VarPKGName & '" has been removed.')
						$FuncRetCode = 0
					Else
						WriteToLog($Log_File, $logTAB & '>INFO ERROR: TAG "' & $VarPKGName & '" has NOT been removed.')
						WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
						$logTAB = $logTABBak
						Finalize(1)
					EndIf
				EndIf
			Else
				WriteToLog($Log_File, $logTAB & '>INFO: TAG "' & $VarPKGName & '" does NOT exist.')
				$FuncRetCode = 0
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;Files operations--------------------------------------------------------------------------------------------------------
Func CopyFile($Src_Dir, $Dest_Dir, $File_Name)
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "CopyFile"
	Local $FuncRetCode					= 0

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			;Check if file exists
			If FileExists($Src_Dir & '\' & $File_Name) Then
				WriteToLog($Log_File, $logTAB & '>INFO: File "' & $Src_Dir & $File_Name & '" exists.')

				;Copy file
				$FuncRetCode = FileCopy($Src_Dir & '\' & $File_Name, $Dest_Dir, 9) ;$FC_OVERWRITE (1) + $FC_CREATEPATH (8)

				If $FuncRetCode = 1 Then
					WriteToLog($Log_File, $logTAB & '>INFO: File "' & $File_Name & '" has been copied to "' & $Dest_Dir & '".')
					$FuncRetCode = 0
				Else
					WriteToLog($Log_File, $logTAB & '>INFO ERROR: File "' & $File_Name & '" has NOT been copied to "' & $Dest_Dir & '".')
					WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
					$logTAB = $logTABBak
					Finalize(1)
				EndIf

			Else
				WriteToLog($Log_File, $logTAB & '>INFO ERROR: File "' & $Src_Dir & $File_Name & '" does NOT exist.')
				WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
				$logTAB = $logTABBak
				Finalize(1)
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
Func RemoveFile($Src_Dir, $File_Name)
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "RemoveFile"
	Local $FuncRetCode					= 0

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			If FileExists($Src_Dir & '\' & $File_Name) Then
				WriteToLog($Log_File, $logTAB & '>INFO: "' & $Src_Dir & '\' & $File_Name & '" exists.')

				$FuncRetCode = FileDelete($Src_Dir & '\' & $File_Name)
				If $FuncRetCode = 1 Then
					WriteToLog($Log_File, $logTAB & '>INFO: File "' & $Src_Dir & '\' & $File_Name & '" has been removed.')
					$FuncRetCode = 0
				Else
					WriteToLog($Log_File, $logTAB & '>INFO ERROR: File "' & $Src_Dir & '\' & $File_Name & '" has NOT been removed.')
					WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
					$logTAB = $logTABBak
					Finalize(1)
				EndIf

			Else
				WriteToLog($Log_File, $logTAB & '>INFO: "' & $Src_Dir & '\' & $File_Name & '" does NOT exist.')
				$FuncRetCode = 0
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;Directory operations----------------------------------------------------------------------------------------------------
Func RemoveDirectory($Dir_Path, $willFail = True)
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "RemoveDirectory"
	Local $FuncRetCode					= 0

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			If FileExists($Dir_Path) Then
				WriteToLog($Log_File, $logTAB & '>INFO: Directory "' & $Dir_Path & '" exists.')
				$FuncRetCode = DirRemove($Dir_Path, 1)

				If $Uninstall_Mode = True Then
					If $FuncRetCode = 1 Then
						WriteToLog($Log_File, $logTAB & '>INFO: Directory "' & $Dir_Path & '" has been removed.')
						$FuncRetCode = 0
					Else
						WriteToLog($Log_File, $logTAB & '>INFO ERROR: Directory "' & $Dir_Path & '" has NOT been removed.')
						WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
						$logTAB = $logTABBak
						If $willFail = False Then
							$FuncRetCode = 1
							Return $FuncRetCode
						Else
							Finalize(1)
						EndIf
					EndIf
				Else
					If $FuncRetCode = 1 Then
						WriteToLog($Log_File, $logTAB & '>INFO: Directory "' & $Dir_Path & '" has been removed.')
						$FuncRetCode = 0
					Else
						WriteToLog($Log_File, $logTAB & '>INFO ERROR: Directory "' & $Dir_Path & '" has NOT been removed.')
						WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
						$logTAB = $logTABBak
						If $willFail = False Then
							$FuncRetCode = 1
							Return $FuncRetCode
						Else
							Finalize(1)
						EndIf
					EndIf
				EndIf
			Else
				WriteToLog($Log_File, $logTAB & '>INFO: Directory "' & $Dir_Path & '" does NOT exist.')
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
Func RemoveDirectoryIfEmpty($Dir_Path)
	Local $Dir_Size = DirGetSize($Dir_Path, 1)
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "RemoveDirectoryIfEmpty"
	Local $FuncRetCode					= 0

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			If FileExists($Dir_Path) = 0 Then
				WriteToLog($Log_File, $logTAB & '>INFO: Folder "' & $Dir_Path & '" doesnt exist.')
				$FuncRetCode = 0
			Else
				If Not @error Then
					WriteToLog($Log_File, $logTAB & '>INFO: Folder "' & $Dir_Path & '" exists.')

					If Not $Dir_Size[1] And Not $Dir_Size[2] Then
						WriteToLog($Log_File, $logTAB & '>INFO: Folder "' & $Dir_Path & '" is empty.')

						$FuncRetCode = DirRemove($Dir_Path)

						If $FuncRetCode = 1 Then
							WriteToLog($Log_File, $logTAB & '>INFO: Folder "' & $Dir_Path & '" has been removed.')
							$FuncRetCode = 0
						Else
							WriteToLog($Log_File, $logTAB & '>INFO ERROR: Folder "' & $Dir_Path & '" has NOT been removed.')
							WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
							$logTAB = $logTABBak
							Finalize(1)
						EndIf
					Else
						WriteToLog($Log_File, $logTAB & '>INFO: Folder "' & $Dir_Path & '" is NOT empty.')
					EndIf
				Else
					WriteToLog($Log_File, $logTAB & '>INFO: Folder "' & $Dir_Path & '" does NOT exist.')
				EndIf
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;EXE operations----------------------------------------------------------------------------------------------------------
Func InstallEXE($Src_Dir, $EXE_FileName, $EXE_Param, $EXE_PreParam = '', $ExicCodes = '0,3010')
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "InstallEXE"
	Local $FuncRetCode					= 0

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			;Check if installer exists
			If FileExists($Src_Dir & "\" & $EXE_Filename) Then
				WriteToLog($Log_File, $logTAB & '>INFO: Installer "' & $Src_Dir & "\" & $EXE_Filename & '" exists.')
				;Set command line
				WriteToLog($Log_File, $logTAB & '>INFO: Running command line: "' & $Src_Dir & '\' & $EXE_Filename & ' '& $EXE_Param & '.')
				;Run installation
				If $EXE_PreParam <> '' Then
					$FuncRetCode = RunWait(@ComSpec & ' /c ' & $EXE_PreParam & ' ' & $Src_Dir & ' ' & $EXE_Filename & ' ' & $EXE_Param)
				Else
					$FuncRetCode = ShellExecuteWait($EXE_Filename, $EXE_Param, $Src_Dir, '', @SW_SHOWDEFAULT)
				EndIf
					$FuncRetCode = RetCodeChecker($FuncRetCode, $ExicCodes)
			Else
				WriteToLog($Log_File, $logTAB & '>INFO ERROR: Installer "' & $Src_Dir & "\" & $EXE_Filename & '" does NOT exist.')
				WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
				$logTAB = $logTABBak
				Finalize(1)
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
Func UninstallEXE($Src_Dir, $EXE_FileName, $EXE_Param, $EXE_UPreParam = '', $ExicCodes = '0,3010')
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "UninstallEXE"
	Local $FuncRetCode					= 0

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			;Check if uninstaller exists
			If FileExists($Src_Dir & "\" & $EXE_Filename) Then
				WriteToLog($Log_File, $logTAB & '>INFO: Uninstaller "' & $Src_Dir & "\" & $EXE_Filename & '" exists.')
				;Set command line
				WriteToLog($Log_File, $logTAB & '>INFO: Running command line: "' & $Src_Dir & '\' & $EXE_Filename & ' '& $EXE_Param & '.')
				;Run uninstallation
				If $EXE_UPreParam <> '' Then
					$FuncRetCode = RunWait(@ComSpec & ' /c ' & $EXE_UPreParam & ' ' & $Src_Dir & ' ' & $EXE_Filename & ' ' & $EXE_Param)
				Else
					$FuncRetCode = ShellExecuteWait($EXE_Filename, $EXE_Param, $Src_Dir, '', @SW_SHOWDEFAULT)
				EndIf
					$FuncRetCode = RetCodeChecker($FuncRetCode, $ExicCodes)
			Else
				WriteToLog($Log_File, $logTAB & '>INFO ERROR: Uninstaller "' & $Src_Dir & "\" & $EXE_Filename & '" does NOT exist.')
				WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
				$logTAB = $logTABBak
				Finalize(1)
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;Other operations--------------------------------------------------------------------------------------------------------
Func CheckFreeDiskSpace($MB_Size)
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "CheckFreeDiskSpace"
	Local $FuncRetCode					= 0
	Local $Free_Disk_Space = DriveSpaceFree($C_Drive & '\')

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			If $Free_Disk_Space > $MB_Size Then
				WriteToLog($Log_File, $logTAB & '>INFO: There is enough free disk space: ' & $Free_Disk_Space & ' MB.')
				$FuncRetCode = 0
			Else
				WriteToLog($Log_File, $logTAB & '>INFO ERROR: There is NOT enough free disk space: ' & $Free_Disk_Space & ' MB.')
				WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
				$logTAB = $logTABBak
				Finalize($ExitCode_NotEnoughFreeDiskSpace)
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
Func RemoveFiveStrike()
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "RemoveFiveStrike"
	Local $FuncRetCode					= 0
	Local $Five_Strike = $Five_Strike_Registry_Root & '\' & $IBM_Apptitle

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			$FuncRetCode = RegDelete($Five_Strike)
			If $FuncRetCode = 0 Or $FuncRetCode = 1  Then
				WriteToLog($Log_File, $logTAB & '>INFO: Strike registry key "' & $Five_Strike & '" has been removed or does NOT exist.')
				$FuncRetCode = 0
			Else
				WriteToLog($Log_File, $logTAB & '>INFO ERROR: Strike registry key "' & $Five_Strike & '" has NOT been removed.')
				WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
				$logTAB = $logTABBak
				Finalize(1)
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;~ ######################################################################################
;~ Message handling code (banners, dialogs etc.).
;~ Code logic loosely mimics 'UlaDialog' code by Robert Sommerville and Randy Conrad.
;~ Modification by packager is expressly prohibited!
;~ ######################################################################################
;------------------------------------------------------------------------------------------------------------------------
Func BannerDisplay($banner_text)
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "BannerDisplay"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			SplashTextOn('', $banner_text, 600, 50, -1, -1, 48, '', 12, 500) ;~ 	no title, text as supplied by param, width, height, centered in x, centered in y, on top + text centered vertically, default font, font size, font weight

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
		$logTAB = $logTABBak
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
Func BannerRemove()
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "BannerRemove"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			SplashOff()

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function')
		$logTAB = $logTABBak
EndFunc
;------------------------------------------------------------------------------------------------------------------------



;------------------------------------------------------------------------------------------------------------------------
Func RetCodeChecker($FuRetCode, $ProperRCList)
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "RetCodeChecker"
	Local $tmp, $i, $IsSuccess

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')
		WriteToLog($Log_File, $logTAB & '>INFO: comparing $FuRetCode = ' & $FuRetCode & ' against $ProperRCList = ' & $ProperRCList & '.')

			$IsSuccess = False
			$tmp = StringSplit($ProperRCList, ',')

			For $i = 1 To $tmp[0]
				If $tmp[$i] = $FuRetCode Then
					$IsSuccess = True
				EndIf
			Next

			If $IsSuccess = True Then
				WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function as success.')
				$logTAB = $logTABBak
				Return 0
			Else
				WriteToLog($Log_File, $logTAB & '>FUNCTION Fail: ' & $FnName & ' function - Finalize script.')
				$logTAB = $logTABBak
				Finalize(1)
			EndIf
EndFunc
;------------------------------------------------------------------------------------------------------------------------



;------------------------------------------------------------------------------------------------------------------------
Func MSIInstall($VarSRC_Dir, $VarMSI_Filename, $VarMST_Filename, $VarMSP_FIlename = '', $VarMSIEXEC_Param = '', $ExicCodes = '0,3010')
	Local $tmp, $msiexec_cmdline
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "MSIInstall"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			$msiexec_cmdline = '/i "'
			If $SRC_Dir <> '' Then
				$msiexec_cmdline = $msiexec_cmdline & $SRC_Dir & '\'
			EndIf

				$msiexec_cmdline = $msiexec_cmdline & $MSI_Filename & '"'

			If $MST_Filename <> '' Then
				If $SRC_Dir <> '' Then
					$msiexec_cmdline = $msiexec_cmdline & ' TRANSFORMS="' & $SRC_Dir & '\' & $MST_Filename & '"'
				Else
					$msiexec_cmdline = $msiexec_cmdline & ' TRANSFORMS="' & $MST_Filename & '"'
				EndIf
			EndIf

			If $MSP_FIlename <> '' Then
				If $SRC_Dir <> '' Then
					$msiexec_cmdline = $msiexec_cmdline & ' PATCH="' & $SRC_Dir & '\' & $MSP_FIlename & '"'
				Else
					$msiexec_cmdline = $msiexec_cmdline & ' PATCH="' & $MSP_FIlename & '"'
				EndIf
			EndIf

			If $MSIEXEC_Param <> '' Then
				$msiexec_cmdline = $msiexec_cmdline & ' ' & $MSIEXEC_Param
			EndIf

				$msiexec_cmdline = $msiexec_cmdline & ' /l*v+ "' & $Msi_Installation_Log & '" /qn REBOOT=ReallySuppress'
					WriteToLog($Log_File, $logTAB & '>INFO: Running msiexec.exe. Command line: ' & $msiexec_cmdline & '.')
				$FuncRetCode = ShellExecuteWait(@SystemDir & '\msiexec.exe', $msiexec_cmdline, @ScriptDir, '', @SW_SHOWDEFAULT)
					$FuncRetCode = RetCodeChecker($FuncRetCode, $ExicCodes)

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------



;------------------------------------------------------------------------------------------------------------------------
Func MSIUninstall($VarGUID, $VarPKGName, $VarMSIEXEC_Param = '', $ExicCodes = '0,3010')
	Local $tmp, $msiexec_cmdline
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "MSIUninstall"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			If $VarMSIEXEC_Param <> '' Then
				$VarMSIEXEC_Param = ' ' & $VarMSIEXEC_Param
			EndIf

				$msiexec_cmdline = '/x ' & $VarGUID & $VarMSIEXEC_Param & ' /l*v+ "' & $ProgramFiles & '\Logs\' & $IBM_AppTitle & '\' & $VarPKGName & '_Uninstall.log' & '" /qn  REBOOT=ReallySuppress'
					WriteToLog($Log_File, $logTAB & '>INFO: Running msiexec.exe. Command line: ' & $msiexec_cmdline & '.')
				$FuncRetCode = ShellExecuteWait(@SystemDir & '\msiexec.exe', $msiexec_cmdline, @ScriptDir, '', @SW_SHOWDEFAULT)

			If $FuncRetCode = 0 and @error = 0 Then
				WriteToLog($Log_File, $logTAB & '>INFO: ' & $VarPKGName & ' was uninstalled successfuly.')
			ElseIf $FuncRetCode = 3010 Then
				$FuncRetCode = 0
				WriteToLog($Log_File, $logTAB & '>INFO: ' & $VarPKGName & ' was uninstalled successfuly. Reboot required...')
			Else
				WriteToLog($Log_File, $logTAB & '>INFO: Uninstallation ' & $VarPKGName & ' has failed. Exit Code: ' & $FuncRetCode & '.')
				Finalize($FuncRetCode)
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------



;------------------------------------------------------------------------------------------------------------------------
Func PrevMSIUninstall($VarGUID, $VarPKGName, $VarLogName, $VarMSIEXEC_Param = '', $ExicCodes = '0,3010')
	Local $tmp, $msiexec_cmdline
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "PrevMSIUninstall"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			If $VarMSIEXEC_Param <> '' Then
				$VarMSIEXEC_Param = ' ' & $VarMSIEXEC_Param
			EndIf

				$msiexec_cmdline = '/x ' & $VarGUID & $VarMSIEXEC_Param & ' /l*v+ "' & $VarLogName & '" /qn  REBOOT=ReallySuppress'
					WriteToLog($Installation_Log, $logTAB & '>INFO: Running msiexec.exe. Command line: ' & $msiexec_cmdline & '.')
				$FuncRetCode = ShellExecuteWait(@SystemDir & '\msiexec.exe', $msiexec_cmdline, @ScriptDir, '', @SW_SHOWDEFAULT)

			If $FuncRetCode = 0 and @error = 0 Then
				WriteToLog($Installation_Log, $logTAB & '>INFO: ' & $VarPKGName & ' was uninstalled successfuly.')
			ElseIf $FuncRetCode = 3010 Then
				$FuncRetCode = 0
				WriteToLog($Installation_Log, $logTAB & '>INFO: ' & $VarPKGName & ' was uninstalled successfuly. Reboot required...')
			Else
				WriteToLog($Installation_Log, $logTAB & '>INFO: Uninstallation ' & $VarPKGName & ' has failed. Exit Code: ' & $FuncRetCode & '.')
				Finalize($FuncRetCode)
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------







;------------------------------------------------------------------------------------------------------------------------
Func isPKGInstalled($VarGUID, $VarVersion, $VarTAGName = '') ; 0 - nothing is installed, 1 - IBM PKG installed, 2 - Vendor installed
	Local $tmp, $isReg = False
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "isPKGInstalled"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			$tmp = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $VarGUID, 'DisplayVersion')
			WriteToLog($Log_File, $logTAB & '>FUNCTION INFO: registry version = ' & $tmp & '')

			If $VarVersion = '' Then
				WriteToLog($Log_File, $logTAB & '>FUNCTION INFO: version not provided, checking for GUID.')
					If $tmp == '' Then
						$tmp = RegRead('HKEY_LOCAL_MACHINE64\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $VarGUID, 'DisplayVersion')
						WriteToLog($Log_File, $logTAB & '>FUNCTION INFO: registry version = ' & $tmp & '')
						If $tmp == '' Then
							WriteToLog($Log_File, $logTAB & '>INFO: ' & $VarGUID & ' is not installed.')
							$FuncRetCode = 0
						Else
							$isReg = True
						EndIf
					Else
						$isReg = True
					EndIf
			Else
					If $tmp <> $VarVersion Then
						$tmp = RegRead('HKEY_LOCAL_MACHINE64\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $VarGUID, 'DisplayVersion')
						WriteToLog($Log_File, $logTAB & '>FUNCTION INFO: registry version = ' & $tmp & '')
						If $tmp <> $VarVersion Then
							WriteToLog($Log_File, $logTAB & '>INFO: ' & $VarGUID & ' is not installed.')
							$FuncRetCode = 0
						Else
							$isReg = True
						EndIf
					Else
						$isReg = True
					EndIf
			EndIf


			If $isReg = True Then
				If $VarTAGName <> '' Then
					If FileExists($ProgramFiles & '\Logs\' & $VarTAGName & '.tag') Then
						If $VarTAGName <> '' Then
							WriteToLog($Log_File, $logTAB & '>INFO: ' & $VarTAGName & ' is installed.')
						Else
							WriteToLog($Log_File, $logTAB & '>INFO: ' & $VarGUID & ' is installed.')
						EndIf
						$FuncRetCode = 1
					Else
						WriteToLog($Log_File, $logTAB & '>INFO: GUID:' & $VarGUID & ' version: ' & $VarVersion & ' is installed, but it is not proper package.')
						$FuncRetCode = 2
					EndIf
				Else
					WriteToLog($Log_File, $logTAB & '>INFO: GUID:' & $VarGUID & ' version: ' & $VarVersion & ' is installed, but it is not proper package.')
					$FuncRetCode = 2
				EndIf
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------



;------------------------------------------------------------------------------------------------------------------------
Func PrevPKGUninstall($VarPKGName, $VarGUID, $VarVersion, $VarInstDir = '', $VarVendorName = '')
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "PrevPKGUninstall"
	Local $msiexec_cmdline
	Local $PrevGUID 					= $VarGUID
	Local $PrevVersionPKGName 			= $VarPKGName
	Local $PrevTAGFile 					= $ProgramFiles & '\Logs\' & $PrevVersionPKGName & ".tag"
	Local $PrevVersionIBMSrc 			= @WindowsDir & '\IBMSRC\' & $PrevVersionPKGName
	Local $PrevVersionIBMSrcEXE 		= $PrevVersionIBMSrc & '\' & $PrevVersionPKGName & '.exe'

	Local $Msi_Uninstall_LogPkg 		= $Log_Dir & $PrevVersionPKGName & '_WRAP_UNINSTALL.log'
	Local $Msi_Uninstall_Log 			= $Log_Dir & $PrevVersionPKGName & '_MSI_UNINSTALL.log'
	Local $PrevPKGLogUninstall 			= $ProgramFiles & '\Logs\' & $PrevVersionPKGName & '\' & $PrevVersionPKGName & '_WRAP_UNINSTALL.log'

	Local $PrevVersionNumber 			= RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $PrevGUID, 'DisplayVersion')

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			If $PrevVersionNumber <> '' Then
				WriteToLog($Log_File, $logTAB & '>FUNCTION INFO: $PrevVersionNumber = ' & $PrevVersionNumber & '.')
			Else
				$PrevVersionNumber = RegRead('HKEY_LOCAL_MACHINE64\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $PrevGUID, 'DisplayVersion')
				If $PrevVersionNumber <> '' Then
					WriteToLog($Log_File, $logTAB & '>FUNCTION INFO: $PrevVersionNumber = ' & $PrevVersionNumber & '.')
				EndIf
			EndIf
			

			If $PrevVersionNumber == $VarVersion Then
				If FileExists($PrevTAGFile) Then 										;previous version (package)
					WriteToLog($Installation_Log, $logTAB & '>INFO: ' & $PrevVersionPKGName & ' is installed, preparing uninstall...')

					If FileExists($PrevVersionIBMSrcEXE) Then
						WriteToLog($Installation_Log, $logTAB & '>INFO: ' & $PrevVersionIBMSrcEXE & ' exist, preparing uninstall...')
						$FuncRetCode = ShellExecuteWait($PrevVersionIBMSrcEXE, '/remove', $PrevVersionIBMSrc, '', @SW_SHOWDEFAULT)

							If NOT FileExists(@WindowsDir & '\UlaDialog3.exe') Then
								WriteToLog($Log_File, $logTAB & '>INFO: UlaDialog3.exe reinstallation to ' & @WindowsDir & '.')
								FileInstall('.\Include\UlaDialog3.exe', @WindowsDir & '\UlaDialog3.exe', 1)
							EndIf
						
						If $FuncRetCode = 0 and @error = 0 Then
							WriteToLog($Installation_Log, $logTAB & '>INFO: ' & $PrevVersionPKGName & ' was uninstalled successfuly.')
						ElseIf $FuncRetCode = 3010 Then
							$FuncRetCode = 0
							WriteToLog($Installation_Log, $logTAB & '>INFO: ' & $PrevVersionPKGName & ' was uninstalled successfuly. Reboot required...')
						Else
							WriteToLog($Installation_Log, $logTAB & '>INFO: Uninstallation ' & $PrevVersionPKGName & ' has failed. Exit Code: ' & $FuncRetCode & '.')
							Finalize($FuncRetCode)
						EndIf

						If FileExists($PrevPKGLogUninstall) Then						;copying LOG file from uninstall previous package version
							WriteToLog($Installation_Log, $logTAB & '>INFO: "' & $PrevPKGLogUninstall & '" exists.')
							$FuncRetCode = FileCopy($PrevPKGLogUninstall, $Log_Dir, 1)
							If $FuncRetCode <> 1 Then
								WriteToLog($Installation_Log, $logTAB & '>INFO: File ' & $PrevPKGLogUninstall & ' copy failed.')
								Finalize(1)
							Else
								WriteToLog($Installation_Log, $logTAB & '>INFO: File ' & $PrevPKGLogUninstall & ' was copied.')
								$FuncRetCode = 0
							EndIf
						Else
							WriteToLog($Installation_Log, $logTAB & '>INFO: "' & $PrevPKGLogUninstall & '" does NOT exists.')
						EndIf
					Else
						WriteToLog($Installation_Log, $logTAB & '>INFO: ' & $PrevVersionIBMSrcEXE & ' do not exist, preparing MSI uninstall...')
							PrevMSIUninstall($PrevGUID, $VarPKGName, $Msi_Uninstall_Log)
							;MSIUninstall($PrevGUID, $VarPKGName)
							RemoveDirectory($PrevVersionIBMSrc, True)

						$FuncRetCode = RegDelete('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $PrevGUID & '_' & $PrevVersionPKGName)
						If $FuncRetCode = 0 Or $FuncRetCode = 1  Then
							WriteToLog($Installation_Log, $logTAB & '>INFO: ARP ' & $PrevVersionPKGName & 'entry has been removed or does NOT exist.')
							$FuncRetCode = 0
						Else
							WriteToLog($Installation_Log, $logTAB & '>INFO: ARP ' & $PrevVersionPKGName & ' entry has NOT been removed.')
							Finalize(1)
						EndIf

						If $VarInstDir <> '' Then
							RemoveDirectory($VarInstDir, True)
						EndIf
					Endif

					RemoveTag($VarPKGName)
					;--------------------------------------------------------
				Else															;previous version (vendor)
					WriteToLog($Installation_Log, $logTAB & '>INFO: ' & $VarVendorName & ' exist but is not a IBM package. Uninstalling...')
					PrevMSIUninstall($PrevGUID, $VarPKGName, $Msi_Uninstall_Log)
				EndIf
			Else
				WriteToLog($Installation_Log, $logTAB & '>INFO: ' & $VarVendorName & ' does NOT exist.')
			EndIf

		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------



;------------------------------------------------------------------------------------------------------------------------
Func SetEnvVar($VarName, $VarValue)
	Local $tmp
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "SetEnvVar"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			Local $tmp
			Local $EnvironmentValue = $VarValue
			Local $EnvironmentKeyName= $VarName
			Local $EnvironmentReg = "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Session Manager\Environment"

				$tmp = RegRead($EnvironmentReg , $EnvironmentKeyName)

				if $tmp <> '' Then
					WriteToLog($Log_File, $logTAB & '>FUNCTION INFO: ENV variable founded: ' & $tmp & ' , removing.')
					RegDelete($EnvironmentReg , $EnvironmentKeyName)
				EndIf

				RegWrite($EnvironmentReg, $EnvironmentKeyName, "REG_SZ", $EnvironmentValue)
				WriteToLog($Log_File, $logTAB & '>FUNCTION INFO: ENV variable writing with new value: ' & $EnvironmentValue & ' .')

				$tmp = RegRead($EnvironmentReg , $EnvironmentKeyName)
				WriteToLog($Log_File, $logTAB & '>FUNCTION INFO: checking if ENV variable: ' & $tmp & ' is correct: ' & $VarValue & '.')

				if $tmp = $VarValue Then
					WriteToLog($Log_File, $logTAB & '>INFO: Environment variable: ' & $EnvironmentKeyName & ' was created which value ' & $EnvironmentValue & "." )
					$FuncRetCode = 0
				Else
					WriteToLog($Log_File, $logTAB & '>INFO: Environment variable: ' & $EnvironmentKeyName & ' was not created.')
					Finalize(1)
				EndIf


			$tmp = EnvGet($VarName)

			If $tmp <> '' Then
				If $tmp == $VarValue Then
						WriteToLog($Log_File, $logTAB & '>FUNCTION INFO: EnvVariable ' & $VarName & ' value was set as: ' & $VarValue & ' ending function.')
					$FuncRetCode = 0
						WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
					$logTAB = $logTABBak
					Return $FuncRetCode
				Else
					 EnvSet($VarName, $VarValue)
					 EnvUpdate()
					 $tmp = EnvGet($VarName)
						If $tmp == $VarValue Then
							WriteToLog($Log_File, $logTAB & '>FUNCTION INFO: EnvVariable ' & $VarName & ' value was set as: ' & $VarValue & ' ending function.')
							$FuncRetCode = 0
								WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
							$logTAB = $logTABBak
							Return $FuncRetCode
						Else
							WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
							$logTAB = $logTABBak
							Finalize(1)
						EndIf
				EndIf
			Else
				EnvSet($VarName, $VarValue)
				EnvUpdate()
				$tmp = EnvGet($VarName)
					If $tmp == $VarValue Then
						WriteToLog($Log_File, $logTAB & '>FUNCTION INFO: EnvVariable ' & $VarName & ' value was created as: ' & $VarValue & ' ending function.')
						$FuncRetCode = 0
							WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
						$logTAB = $logTABBak
						Return $FuncRetCode
					Else
						WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
						$logTAB = $logTABBak
						Finalize(1)
					EndIf
			EndIf


		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------



;------------------------------------------------------------------------------------------------------------------------
Func RegHideARP($GUID)
	Local $temp, $x64, $Source_Key
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "RegHideARP"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			$x64 = '64'

			$temp = RegRead('HKEY_LOCAL_MACHINE' & $x64 & '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $GUID, 'DisplayName')

			If $temp = '' Then
				WriteToLog($Log_File, $logTAB & '>INFO: HKEY_LOCAL_MACHINE' & $x64 & '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $GUID & '\DisplayName, not found.')
					$temp = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $GUID, 'DisplayName')
					If $temp = '' Then
						WriteToLog($Log_File, $logTAB & '>INFO: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $GUID & '\DisplayName, not found.')
					Else
						$Source_Key = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $GUID
					EndIf
			Else
				$Source_Key = 'HKEY_LOCAL_MACHINE' & $x64 & '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $GUID
			EndIf

			RegWrite($Source_Key, 'SystemComponent', 'REG_DWORD', '1') ;Hide original entry


		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------



;------------------------------------------------------------------------------------------------------------------------
Func SelfCopySrc()
	Local $tmp, $x64, $Source_Key
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "SelfCopySrc"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')


			FileCopy(@ScriptFullPath, @WindowsDir & '\IBMSRC\' & $IBM_Apptitle & '\' & @ScriptName, 9)


		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------



;------------------------------------------------------------------------------------------------------------------------
Func isCMDProcessRunning($proc, $procCMD = '', $strComputer = '.')
	Local $oProcessColl, $Process, $ProcessCMD, $oWMI
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "isCMDProcessRunning"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

		$oWMI = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & $strComputer & "\root\cimv2")
		$oProcessColl = $oWMI.ExecQuery("Select * from Win32_Process where Name= " & '"'& $Proc & '"')

		If $oProcessColl.Length <= 0 Then
			WriteToLog($Log_File, $logTAB & '>FUNCTION INFO: process ' & $proc & ' not running.')
			WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
			$logTAB = $logTABBak
			Return 1
		EndIf

	For $Process In $oProcessColl
		$ProcessCMD = $Process.Commandline
		If $procCMD <> '' Then
			If StringInStr($ProcessCMD, $procCMD) <> 0 Then
				WriteToLog($Log_File, $logTAB & '>FUNCTION INFO: process ' & $proc & ' with CMD ' & $procCMD & ' is running.')
				WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
				$logTAB = $logTABBak
				Return 0
			EndIf
		EndIf
	Next

	WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
	$logTAB = $logTABBak
	Return 1
EndFunc
;------------------------------------------------------------------------------------------------------------------------



;------------------------------------------------------------------------------------------------------------------------
Func SwitchService($service_name)
	Global $Service_States[20][3] ; [1] name [2] startup type [3] running?
	Local $i, $objWMIService, $objWMIQuery, $objService, $rc
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FuncRetCode					= 0
	Local $FnName						= "SwitchService"

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			$objWMIService = ObjGet('winmgmts:\\.\root\cimv2')

			If Not IsObj($objWMIService) Then
				WriteToLog($Log_File, $logTAB & '>INFO: Unable to open "winmgmts:\\.\root\cimv2" object.')
				Return 1
			EndIf

			For $i = 1 To UBound($Service_States) -1
				If StringCompare($service_name, $service_states[$i][1], 0) = 0 Then
					$objWMIQuery = $objWMIService.ExecQuery("Select * from Win32_Service where Name = '" & $service_name & "'")
					For $objService In $objWMIQuery
						If Not IsObj($objService) Then
							WriteToLog($Log_File, $logTAB & '>INFO: Win32_Service query returned null.')
							WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
							$logTAB = $logTABBak
							Return 1
						EndIf
						$rc = $objService.ChangeStartMode($service_states[$i][2])
						If $service_states[$i][3] = True Then $objService.StartService()
						$service_states[$i][1] = ''
						$service_states[$i][2] = ''
						$service_states[$i][3] = ''
					Next
						WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
						$logTAB = $logTABBak
					Return 0
				EndIf
			Next

			For $i = 1 To UBound($Service_States) -1
				If $service_states[$i][1] = '' Then ExitLoop
			Next

			$objWMIQuery = $objWMIService.ExecQuery("Select * from Win32_Service where Name = '" & $service_name & "'")
			For $objService In $objWMIQuery
				If Not IsObj($objService) Then
					WriteToLog($Log_File, $logTAB & '>INFO: Win32_Service query returned null.')
					WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
					$logTAB = $logTABBak
					Return 1
				EndIf
				$service_states[$i][1] = $service_name
				$service_states[$i][2] = $objService.StartMode
				$service_states[$i][3] = $objService.Started
				If StringCompare($service_states[$i][2], 'Automatic', 0) = 0 Then $service_states[$i][2] = 'Auto'

				If $objService.Started = True Then $objService.StopService()
				If $objService.StartMode <> 'Disabled' Then $objService.ChangeStartMode('Disabled')
				WriteToLog($Log_File, $logTAB & '>INFO: Stopping "' & $service_name & '" service.')
			Next

	WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
	$logTAB = $logTABBak
	Return 0
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
;~ Sets NoModify in ARP entry
Func SetNoModifyARP($AGUID)
	Local $Source_Key, $temp
	Local $x64 = ''
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "SetNoModifyARP"
	Local $FuncRetCode					= 0

		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')

			$x64 = '64'


			$temp = RegRead('HKEY_LOCAL_MACHINE' & $x64 & '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $AGUID, 'DisplayName')

			If $temp = '' Then
				WriteToLog($Log_File, $logTAB & '>INFO: HKEY_LOCAL_MACHINE' & $x64 & '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $AGUID & '\DisplayName, not found.')
					$temp = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $AGUID, 'DisplayName')
					If $temp = '' Then
						WriteToLog($Log_File, $logTAB & '>INFO: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $AGUID & '\DisplayName, not found.')
					Else
						$Source_Key = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $AGUID
					EndIf
			Else
				$Source_Key = 'HKEY_LOCAL_MACHINE' & $x64 & '\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' & $AGUID
			EndIf


			RegWrite($Source_Key, 'NoModify', 'REG_DWORD', '1')


		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------


;------------------------------------------------------------------------------------------------------------------------
Func removeAllPrevBuilds()
	Local $temp, $shortPKGName, $maxTemp, $nOpe, $buildNumber, $i, $j, $PrevVersionIBMSrc, $PrevVersionIBMSrcEXE, $PrevVersionPKGName
	Local $logTABBak					= $logTAB
		  $logTAB 						= $logTAB & $logSeparator
	Local $FnName						= "removeAllPrevBuilds"
	Local $FuncRetCode					= 0


		WriteToLog($Log_File, $logTAB & '>FUNCTION START: ' & $FnName & ' function.')


			$temp = StringSplit($IBM_AppTitle, "_")

			$maxTemp = $temp[0]

			For $i = 1 To $maxTemp-1
				If $i == 1 Then
					$shortPKGName = $temp[$i]
				Else
					$shortPKGName = $shortPKGName & '_' & $temp[$i]
				EndIf
			Next

			$nOpe = StringInStr($temp[$maxTemp], "B")


			If $nOpe == 0 Then
				$FuncRetCode = 1
			Else
				$buildNumber = StringTrimLeft ($temp[$maxTemp], $nOpe)
			EndIf


			If Int($buildNumber) == 1 Then
				$FuncRetCode = 0
			Else
				For $j = Int($buildNumber) To $j = 0
					$PrevVersionPKGName = $shortPKGName & '_B' & $buildNumber
					$PrevVersionIBMSrc = $IBMSRC_Dir & $PrevVersionPKGName
					$PrevVersionIBMSrcEXE = $PrevVersionIBMSrc & '\' & $PrevVersionPKGName & '.exe'

					If FileExists($PrevVersionIBMSrcEXE) Then
						WriteToLog($Installation_Log, $logTAB & '>INFO: ' & $PrevVersionIBMSrcEXE & ' exist, preparing uninstall...')
						$FuncRetCode = ShellExecuteWait($PrevVersionIBMSrcEXE, '/remove', $PrevVersionIBMSrc, '', @SW_SHOWDEFAULT)
					EndIf
				Next
			EndIf

;$IBM_AppTitle




		WriteToLog($Log_File, $logTAB & '>FUNCTION END: ' & $FnName & ' function.')
		$logTAB = $logTABBak
		Return $FuncRetCode
EndFunc
;------------------------------------------------------------------------------------------------------------------------