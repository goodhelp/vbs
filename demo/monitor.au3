Const $Dxva2 = DllOpen("Dxva2.dll")
Global Const $tagSTRUCT = "HANDLE hPhysicalMonitor;WCHAR szPhysicalMonitorDescription[128];"
$hMonitor = _WinAPI_MonitorFromWindow(_WinAPI_GetDesktopWindow(), 1) ;返回主显示器句柄
$hMonitor = _GetPhysicalMonitorsFromHMONITOR($hMonitor) ;获取物理显示器句柄
_RestoreMonitorFactoryDefaults($hMonitor );
Func _GetPhysicalMonitorsFromHMONITOR(Const $pMonitor)
        Local $Number = DllStructCreate("DWORD")
        Local $i, $tagPhysical, $M = 1
        $Ret = DllCall($Dxva2, "bool", "GetNumberOfPhysicalMonitorsFromHMONITOR", _
                        "handle", $pMonitor, _
                        "ptr", DllStructGetPtr($Number))		
        $NumberOfMonitors = DllStructGetData($Number, 1)
        For $i = 1 To $NumberOfMonitors
                $tagPhysical &= $tagSTRUCT
        Next
        $MonitorArray = DllStructCreate($tagSTRUCT)
        $Ret = DllCall($Dxva2, "bool", "GetPhysicalMonitorsFromHMONITOR", _
                        "handle", $pMonitor, _
                        "DWORD", $NumberOfMonitors, _
                        "ptr", DllStructGetPtr($MonitorArray))
        Return DllStructGetData($MonitorArray, 1) ;返回第一个显示器句柄
EndFunc   ;==>_GetPhysicalMonitorsFromHMONITOR

Func _RestoreMonitorFactoryDefaults(Const $h_monitor)
        DllCall($Dxva2, "bool", "RestoreMonitorFactoryDefaults", "ptr", $h_monitor)
        ;Return (@error) ? (SetError(1, _WinAPI_GetLastErrorMessage(), False)) : (True)
EndFunc   ;==>_RestoreMonitorFactoryDefaults

Func _WinAPI_MonitorFromWindow($hWnd, $iFlag = 1)
	Local $aRet = DllCall('user32.dll', 'handle', 'MonitorFromWindow', 'hwnd', $hWnd, 'dword', $iFlag)
	If @error Then Return SetError(@error, @extended, 0)

	Return $aRet[0]
EndFunc   ;==>_WinAPI_MonitorFromWindow

Func _WinAPI_GetDesktopWindow()
	Local $aResult = DllCall("user32.dll", "hwnd", "GetDesktopWindow")
	If @error Then Return SetError(@error, @extended, 0)

	Return $aResult[0]
EndFunc   ;==>_WinAPI_GetDesktopWindow