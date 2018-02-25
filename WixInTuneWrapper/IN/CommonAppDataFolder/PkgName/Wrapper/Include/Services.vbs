'****************************************************************************************************************************************************
' FUNCTIONS:
'	SetService(ServiceName, Mode, StartupType)
'	
' FUNCTION EXAMPLE USAGE (IN and OUT):
'	Call SetService("AxInstSV", "Start", "Automatic")
'	Return 0, 1
'	
'	Acceptable Mode values (case insensitive): Start, Stop
'	Acceptable StartupType values (case insensitive): Automatic (Delayed Start), Automatic, Manual, Disabled
'***************************************************************************************************************************************************
'! Control over system services, it can stop, start or change it's mode.
'!
'! @param  ServiceName      Existing Service Name.
'! @param  Mode             Acceptable Mode values (case insensitive): Start, Stop.
'! @param  StartupType      Acceptable StartupType values (case insensitive): Automatic (Delayed Start), Automatic, Manual, Disabled.
'!
'! @return 0,1
Function SetService(ServiceName, Mode, StartupType)
	Dim objListOfServicesAll, objServiceAll, objListOfServices, objService, S_startmode
    Dim strCurrentFNName                   : strCurrentFNName = "SetService"
	Dim IsService						: IsService = False
	Dim strLogTABBak						: strLogTABBak = strLogTAB
		strLogTAB = strLogTAB & strLogSeparator
            
        fnAddLog(StartingFN(0) & strCurrentFNName & StartingFN(1) & "ServiceName - " & ServiceName & ", Mode - " & Mode & ", StartupType - " & StartupType & DotFN(0))
	
    'TODO param check
    
    
	'----- Check correctness of ServiceName parameter:
		Set objListOfServicesAll = objWMIService.ExecQuery("SELECT * FROM Win32_Service")
			For Each objServiceAll In objListOfServicesAll
				If objServiceAll.Name = ServiceName Then
					IsService = True
				End If
			Next
		Set objListOfServicesAll = Nothing
		
		If IsService = False Then
			fnAddLog("FUNCTION ERROR: service " & ServiceName & " does not exist.")
			fnAddLog("FUNCTION INFO: ending SetService - fail.")
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			SetService = 1
			Exit Function
		End If
		
	'----- Check correctness of Mode parameter:
		Mode = LCase(Mode)					'Convert Mode value to lowercase.
		
		If Mode <> "start" AND Mode <> "stop" Then
			fnAddLog("FUNCTION ERROR: " & Mode & " is not a proper service mode.")
			fnAddLog("FUNCTION INFO: ending SetService - fail.")
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			SetService = 1
			Exit Function
		End If
		
	'----- Check correctness of StartupType parameter:
		StartupType = LCase(StartupType)	'Convert StartupType value to lowercase.
		
		If StartupType <> "automatic (delayed start)" AND StartupType <> "automatic" AND StartupType <> "manual" AND StartupType <> "disabled" Then
			fnAddLog("FUNCTION ERROR: " & StartupType & " is not a proper service startup type.")
			fnAddLog("FUNCTION INFO: ending SetService - fail.")
			
			strLogTAB = strLogTABBak
			fnFinalize(1)
			SetService = 1
			Exit Function
		End If
		
		
		Set objListOfServices = objWMIService.ExecQuery("SELECT * FROM Win32_Service WHERE Name = '" & ServiceName & "'")
			For Each objService In objListOfServices
				If objService.State = "Running" Then
					fnAddLog("FUNCTION INFO: " & ServiceName & " is running.")
					If Mode = "start" Then
						fnAddLog("FUNCTION INFO: No need to start service.")
					Else
					'STOP service:
						objService.StopService()
						fnAddLog("FUNCTION INFO: " & ServiceName & " has been stopped.")
					End If
				Else
					fnAddLog("FUNCTION INFO: " & ServiceName & " is stopped.")
					If Mode = "stop" Then
						fnAddLog("FUNCTION INFO: No need to stop service.")
					Else
					'START service:
						S_startmode = objService.StartMode
						If S_startmode = "Disabled" Then
							objService.ChangeStartMode("Automatic")
							Wscript.Sleep 1000
						End If
						objService.StartService()
						fnAddLog("FUNCTION INFO: " & ServiceName & " has been started.")
					End If
				End If
				
				S_startmode = objService.StartMode
				If S_startmode = "Auto" Then
					S_startmode = "automatic"
				End If
				
				'Change service Startup Type:
				If S_startmode <> StartupType Then
					objService.ChangeStartMode(StartupType)
					fnAddLog("FUNCTION INFO: Startup Type has been changed to: " & StartupType & ".")
				Else
					fnAddLog("FUNCTION INFO: No need to change Startup Type.")
				End If
				
				Wscript.Sleep 2000
			Next
		Set objListOfServices = Nothing
		
		
		fnAddLog("FUNCTION INFO: ending SetService - success.")
		
		strLogTAB = strLogTABBak
		SetService = 0
        wscript.Sleep 2000
    
		Exit Function
End Function