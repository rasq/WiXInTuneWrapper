#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=UlaDIalog3.ico
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Res_Description=ULA Script for user communication
#AutoIt3Wrapper_Res_Fileversion=1.0.0.82
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=Computer Sciences Corporation
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_Field=Author| Robert Sommerville and Randy Conrad
#AutoIt3Wrapper_Res_Field=Date| Jan 3, 2012
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs ----------------------------------------------------------------------------

	This has been modified to removed the check for the existance of MSI.
	It only checks for the registry entry 1 to appear, if it does not,
	then this dialog will presist until something kills the process.

	Make sure that you are always able to write the "HKEY_LOCAL_MACHINE\SOFTWARE\CSC" "MessageClose" = "1
	registry entry. If for some reason you can not, then just run a UlaDialog2.exe
	kill process when you want this EXE / Banner to go away.

#ce ----------------------------------------------------------------------------

#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#include <Timers.au3>
#include <Date.au3>
#include <ProgressConstants.au3>


If $CmdLine[1] = "B" Then
	Global $Type 				= $CmdLine[1] 		; setsinteractive dialog type, D=dilaog with timer, no buttons, DB= Dialog with ony one button, DBB= Dialog with 2 buttons, Banner= Banner only and waits on reg
	Global $Prompt 				= $CmdLine[2] 		; Prompt text
	Global $Button1 			= $CmdLine[3] 		; first button label
	Global $Button2 			= $CmdLine[4] 		; second button label
	Global $iCountdown 			= $CmdLine[5] * 60	; Sets the time-out in minutes
	Global $DefaultExitCode 	= $CmdLine[6] 		; set exit code when user lets dialog time out.
	Global $iTotal_Time 		= $iCountdown		; Copies the starting  to another variable.
	Global $Progress, $Label1, $Label2, $Form1, $UserAction1, $UserAction2
Else

	Global $Type 				= $CmdLine[1] 		; setsinteractive dialog type, D=dilaog with timer, no buttons, DB= Dialog with ony one button, DBB= Dialog with 2 buttons, Banner= Banner only and waits on reg
	Global $Prompt 				= $CmdLine[2] 		; Prompt text
	Global $Button1 			= $CmdLine[3] 		; first button label
	Global $Button2 			= $CmdLine[4] 		; second button label
	Global $iCountdown			= $CmdLine[5] * 60	; Sets the time-out in minutes
	Global $DefaultExitCode 	= $CmdLine[6] 		; set exit code when user lets dialog time out.
	Global $iTotal_Time 		= $iCountdown		; Copies the starting  to another variable.
	Global $Progress, $Label1, $Label2, $Form1, $UserAction1, $UserAction2
EndIf

If ProcessExists("explorer.exe") Then
	Call("_StartupDialogs")
Else
	Exit (0)
EndIf

Func _StartupDialogs()
	If $Type = "D" Then Call("_Dialogbox")
	If $Type = "B" Then Call("_Dialogbox2")
EndFunc   ;==>_StartupDialogs


Func _Dialogbox()
	; Creates GUI Window
	; BitOR($WS_MINIMIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS), $WS_EX_TOPMOST)
	; BitOR($WS_CAPTION, $WS_POPUP, $WS_BORDER, $WS_CLIPSIBLINGS), $WS_EX_WINDOWEDGE)
	If $iCountdown <> 0 Then
		Local $yOffset = 170
	Else
		Local $yOffset = 170
	EndIf

	$Form1 = GUICreate("ULA IT Communication", 520, $yOffset, -1, -1, BitOR($WS_MINIMIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS), $WS_EX_TOPMOST)

	If $iCountdown > 0 Then
		;Local $yOffsetA = $yOffset - 28
		$Progress = GUICtrlCreateProgress(10, 142, 300, 15, $PBS_SMOOTH) ; Creates Progress Bar
	Else
	EndIf

	GUICtrlSetColor($Progress, 0xff0000) ; Sets Progress Bar Colour
	;$yOffset = $yOffset - 33

	If $Button1 <> "" Then $UserAction1 = GUICtrlCreateButton($Button1, 335, 137, 80, 23, 0) ; Creates Restart Now Button
	If $Button2 <> "" Then $UserAction2 = GUICtrlCreateButton($Button2, 425, 137, 80, 23, 0) ; Creates Restart Latter Button

	If $iCountdown <> 0 Then
		$Label1 = GUICtrlCreateLabel($Prompt, 10, 10, 490, 120)
		;Local $yOffsetB = $yOffset - 20
		$Label2 = GUICtrlCreateLabel(_SecsToTime($iCountdown) & " minutes.", 10, 127, 100, 13)
	Else
		$Label1 = GUICtrlCreateLabel($Prompt, 10, 10, 490, 90)
	EndIf

	GUISetState(@SW_SHOW)

	_Timer_SetTimer($Form1, 1000, '_Countdown')

	If $Button1 <> "" And $Button2 <> "" Then
		While 1
			$nMsg = GUIGetMsg()
			Switch $nMsg
				Case $UserAction1
					Exit (0)
				Case $UserAction2
					Exit (1)
			EndSwitch
		WEnd
	EndIf

	If $Button1 <> "" And $Button2 = "" Then
		While 1
			$nMsg = GUIGetMsg()
			Switch $nMsg
				Case $UserAction1
					Exit (0)
			EndSwitch
		WEnd
	EndIf

	If $Button1 = "" And $Button2 <> "" Then
		While 1
			$nMsg = GUIGetMsg()
			Switch $nMsg
				Case $UserAction2
					Exit (1)
			EndSwitch
		WEnd
	EndIf
EndFunc   ;==>_Dialogbox

Func _Dialogbox2()
	$Form1 = GUICreate("ULA IT Communication", 370, 100, -1, -1, BitOR($WS_MINIMIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS), $WS_EX_TOPMOST) ; Creates GUI Window
	;BitOR($WS_MINIMIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_GROUP, $WS_BORDER, $WS_CLIPSIBLINGS), $WS_EX_TOPMOST)
	; BitOR($WS_CAPTION, $WS_POPUP, $WS_BORDER, $WS_CLIPSIBLINGS), $WS_EX_WINDOWEDGE)
	If $iCountdown > 0 Then
		$Progress = GUICtrlCreateProgress(10, 75, 230, 15, $PBS_SMOOTH) ; Creates Progress Bar
	Else
	EndIf
	GUICtrlSetColor($Progress, 0xff0000) ; Sets Progress Bar Colour
	If $Button1 <> "" Then $UserAction1 = GUICtrlCreateButton($Button1, 280, 70, 80, 23, 0) ; Creates Restart Now Button
	;If $Button2 <> "" Then $UserAction2 = GUICtrlCreateButton($Button2, 290, 70, 80, 23, 0) ; Creates Restart Latter Button
	If $iCountdown <> 0 Then
		$Label1 = GUICtrlCreateLabel($Prompt, 10, 10, 330, 50)
		$Label2 = GUICtrlCreateLabel(_SecsToTime($iCountdown) & " minutes.", 10, 60, 100, 13)

	Else
		$Label1 = GUICtrlCreateLabel($Prompt, 10, 10, 330, 50)
	EndIf
	GUISetState(@SW_SHOW)

	_Timer_SetTimer($Form1, 1000, '_Countdown')



	;If $Button1 <> "" And $Button2 <> "" Then

		;While 1
			;$nMsg = GUIGetMsg()
			;Switch $nMsg
				;Case $UserAction1
				;	Exit (0)
				;Case $UserAction2
				;	Exit (1)
			;EndSwitch
		;WEnd

	;EndIf

	If $Button1 <> ""  Then ;And $Button2 = "" Then

		While 1
			$nMsg = GUIGetMsg()
			Switch $nMsg
				Case $UserAction1
					Exit (0)
			EndSwitch
		WEnd

	EndIf

	If $Button1 = "" Then ;And $Button2 <> "" Then

		While 1
			$nMsg = GUIGetMsg()
			Switch $nMsg
				Case $UserAction2
					Exit (1)
			EndSwitch
		WEnd

	EndIf
EndFunc   ;==>_Dialogbox2

Func _Countdown($hWnd, $iMsg, $iIDTimer, $dwTime)
	$iCountdown -= 1
	$percent_value = Floor(($iCountdown / $iTotal_Time) * 100)
	$percent_value = 100 - $percent_value
	If $iCountdown > 0 Then
		GUICtrlSetData($Label2, _SecsToTime($iCountdown) & " minutes.")
		GUICtrlSetData($Progress, $percent_value)
	ElseIf $iCountdown = 0 Then
		GUICtrlSetData($Label2, _SecsToTime($iCountdown) & " minutes.")
		GUICtrlSetData($Progress, $percent_value)
		_Timer_KillTimer($hWnd, $iIDTimer)
		;ControlClick($Form1, '', $UserAction1) ; Default action when Countdown equals 0
		If $DefaultExitCode = ("0") Then
			Exit (0) ; Default action when Countdown equals 0
		EndIf
		If $DefaultExitCode = ("1") Then
			Exit (1) ; Default action when Countdown equals 0
		EndIf
	EndIf
EndFunc   ;==>_Countdown

Func _SecsToTime($iSecs)
	Local $iHours, $iMins, $iSec_s
	_TicksToTime($iSecs * 1000, $iHours, $iMins, $iSec_s)
	Return StringFormat("%02i:%02i", $iMins, $iSec_s)
EndFunc   ;==>_SecsToTime
