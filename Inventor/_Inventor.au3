Opt("WinTitleMatchMode", 2)

Global $sInventorTitle = "Autodesk Inventor"
Global $sInventorFrame = "[CLASS:AfxFrameOrView90u; INSTANCE:1],[CLASS:AfxFrameOrView110u; INSTANCE:2]"

Func _Inventor_FrameGetPos()
	Local $aPos
	Local $aFrame = StringSplit($sInventorFrame, ",")
	For $vFrame In $aFrame
		$aPos = ControlGetPos($sInventorTitle, "", $vFrame)
		If IsArray($aPos) Then
			Return $aPos
			Exit
		EndIf
    Next
EndFunc

Func _Inventor_FrameExists()
	Local $aPos = _Inventor_FrameGetPos()
	Return IsArray($aPos)
EndFunc

Func _Inventor_GetPos()
	Local $aPos = WinGetPos($sInventorTitle, "")
	Return $aPos
EndFunc

Func _Inventor_Active()
	Return WinActive($sInventorTitle)
EndFunc

Func _Inventor_Exists()
	Return WinExists($sInventorTitle)
EndFunc