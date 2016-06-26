#include <Array.au3>
#include <ButtonConstants.au3>
#include <ColorConstantS.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>

#include "_Inventor.au3"

Opt("MouseCoordMode", 1)

Global $g_aButtons[0]
Global $g_hGUI = 0
Global $g_iSize = 40
Global $g_sModule = "Toolbar"
Global $g_sToolbarName = "Common Tools"


Func InventorToolbar_Create()
	Local $hGUI = 0
	Local $sPath = _Inventor_VBAProjectPath()

	Local $aPos = _Inventor_FrameGetPos()
	If Not IsArray($aPos) Then
		Local $aPos[] = [0, 0]
	EndIf

	$g_aButtons = _Inventor_MacroArray($g_sModule)
	_ArrayColInsert($g_aButtons,1)

	$hGUI = GUICreate($g_sToolbarName, $g_iSize*UBound($g_aButtons, $UBOUND_ROWS), $g_iSize, $aPos[0], $aPos[1], $WS_POPUP, BitOr($WS_EX_TOPMOST, $WS_EX_TRANSPARENT, $WS_EX_TOOLWINDOW),0)

	Local $i
	For $i = 0 To UBound($g_aButtons, $UBOUND_ROWS) - 1
		$g_aButtons[$i][1] = GUICtrlCreateButton($g_aButtons[$i][0], $g_iSize * $i, 0, $g_iSize, $g_iSize, $BS_BITMAP)

		Local $sIcon = $sPath & $g_sModule & "." & $g_aButtons[$i][0] & ".bmp"
		If FileExists($sIcon) Then
			GUICtrlSetImage(-1, $sIcon)
		EndIF
	Next

	GUISetState(@SW_HIDE)

	Return $hGUI
EndFunc

Func InventorToolbar_MouseInRegion($aPosM, $aPosI, $aPosF)
	Return _
		($aPosM[0] >= $aPosI[0] + $aPosF[0]) And ($aPosM[0] <= $aPosI[0] + $aPosF[0] + $aPosF[2]) And _
		($aPosM[1] >= $aPosI[1] + $aPosF[1]) And ($aPosM[1] <= $aPosI[1] + $aPosF[1] + 150)
EndFunc

Func InventorToolbar_Display()
	Local $aPosM = MouseGetPos()
	Local $aPosI = _Inventor_GetPos()
	Local $aPosF = _Inventor_FrameGetPos()

	If IsArray($aPosI) and IsArray($aPosF) Then
		WinMove($g_hGUI, "",  $aPosI[0] + $aPosF[0] + 10, $aPosI[1] + $aPosF[1] + 2)
	EndIf

	If _Inventor_FrameExists($g_sToolbarName) and InventorToolbar_MouseInRegion($aPosM, $aPosI, $aPosF) Then
		GUISetState(@SW_SHOW)
	Else
		GUISetState(@SW_HIDE)
	EndIf
EndFunc

Func InventorToolbar_Initialize()
	If $g_hGUI = 0 Then
		$g_hGUI = InventorToolbar_Create()
	EndIf
EndFunc

Func InventorToolbar_Terminate()
	If $g_hGUI > 0 Then
		GUIDelete($g_hGUI)
	EndIf
EndFunc

Func Main()
	Local $iMsg
	Local $iButton

	While 1
		If _Inventor_Exists() Then
			InventorToolbar_Initialize()
			InventorToolbar_Display()

			$iMsg = GUIGetMsg()
			$iButton = _ArraySearch($g_aButtons, $iMsg, 0, 0, 0, 2, 1, 1)
			If $iButton >= 0 Then
				_Inventor_MacroExecute($g_sModule, $g_aButtons[$iButton][0])
			EndIf
		Else
			InventorToolbar_Terminate()
		EndIf

		Sleep(100)
	Wend

	InventorToolbar_Terminate()
EndFunc

Main()
