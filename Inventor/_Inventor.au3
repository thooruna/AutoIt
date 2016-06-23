Opt("WinTitleMatchMode", 2)

Global $sInventorTitle = "Autodesk Inventor"
Global $sInventorFrame = "[CLASS:AfxFrameOrView90u; INSTANCE:1],[CLASS:AfxFrameOrView110u; INSTANCE:2]"

Func _Inventor_FrameGetPos()
	Local $aPosF
	Local $aFrame = StringSplit($sInventorFrame, ",")
	For $vFrame In $aFrame
		$aPosF = ControlGetPos($sInventorTitle, "", $vFrame)
		If IsArray($aPosF) Then
			Return $aPosF
		EndIf
    Next
EndFunc

Func _Inventor_FrameExists()
	$aPosF = _Inventor_FrameGetPos()
	Local $sTitle = WinGetTitle("[ACTIVE]")
	Return ((StringLeft($sTitle, 4) = "test") Or (StringLeft($sTitle, StringLen($sInventorTitle)) = $sInventorTitle)) And IsArray($aPosF)
EndFunc

Func _Inventor_GetPos()
	Local $aPosI = WinGetPos($sInventorTitle, "")
	Return $aPosI
EndFunc

Func _Inventor_Active()
	Return WinActive($sInventorTitle)
EndFunc

Func _Inventor_Exists()
	Return WinExists($sInventorTitle)
EndFunc