#include <ButtonConstants.au3>
#include <ColorConstantS.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>

#include "_Inventor.au3"

; Tray
Opt("TrayOnEventMode", 1)
Opt("TrayMenuMode", 1) ; Default tray menu items (Script Paused/Exit) will not be shown.

   Local $idButton1, $idButton2, $idButton3, $idButton4
   Local $idButton5, $idButtonclose

Func RunInventorMacro($strModule, $strFunction)
	; Get a reference to Inventor.  This assumes Inventor is running.
	Local $invApp = ObjGet("", "Inventor.Application")

	If @error Then TrayTip ( "Error", "Inventor VBA Error 1", 1, 3)

	; Get a reference to the default VBA project.
	; It will always be the first item in the collection.
	Local $invVBAProject = $invApp.VBAProjects.Item(1)

	If @error Then TrayTip ( "Error", "Inventor VBA Error 2", 1, 3)

	; Get a reference to the code module.
	Local $invModule = $invVBAProject.InventorVBAComponents.Item($strModule)

	If @error Then TrayTip ( "Error", "Inventor VBA Error 3", 1, 3)

	; Get a reference to the function.
	Local $invSub = $invModule.InventorVBAMembers.Item($strFunction)

	If @error Then TrayTip ( "Error", "Inventor VBA Error 4", 1, 3)

	; Execute the macro.
	Local $result = $invSub.Execute

	If @error Then TrayTip ( "Error", "Inventor VBA Error 5", 1, 3)
EndFunc   ;==>RunInventorMacro

Func ToolBarMove($hGUI)
	If _Inventor_Exists() Then
		Local $aPosI = _Inventor_GetPos()
		Local $aPosF = _Inventor_FrameGetPos()
		If IsArray($aPosI) and IsArray($aPosF) Then
			WinMove($hGUI, "",  $aPosI[0] + $aPosF[0] + 10, $aPosI[1] + $aPosF[1] + 2)
		EndIf
	EndIf
EndFunc

Func CreateToolBar()
	Local $iHeight = 40
	Local $aPos = _Inventor_FrameGetPos()
	Local $hGUI = GUICreate("test", $iHeight*6, $iHeight, $aPos[0], $aPos[1], $WS_POPUP, BitOr($WS_EX_TOPMOST, $WS_EX_TRANSPARENT, $WS_EX_TOOLWINDOW),0)

	$idButton1 = GUICtrlCreateButton("1", 0, 0, $iHeight, $iHeight, $BS_BITMAP)
	GUICtrlSetImage(-1, "P:\inventor\2016\Inventor\Macros\OpenDrawing.OpenDrawing.Large.bmp")
	;GUICtrlSetImage(-1, "shell32.dll", 5)
	$idButton2 = GUICtrlCreateButton("2", $iHeight, 00, $iHeight, $iHeight, $BS_BITMAP)
	GUICtrlSetImage(-1, "P:\inventor\2016\Inventor\Macros\IV Customisations\Customising 03b Ribbon.PNG")
	;GUICtrlSetImage(-1, "shell32.dll", 7)
	$idButton3 = GUICtrlCreateButton("3", $iHeight*2, 00, $iHeight, $iHeight, $BS_ICON)
	GUICtrlSetImage(-1, "shell32.dll", 22)
	$idButton4 = GUICtrlCreateButton("4", $iHeight*3, 0, $iHeight, $iHeight, $BS_ICON)
	GUICtrlSetImage(-1, "shell32.dll", 23)
	$idButton5 = GUICtrlCreateButton("5", $iHeight*4, 0, $iHeight, $iHeight, $BS_ICON)
	GUICtrlSetImage(-1, "shell32.dll", 32)
	$idButtonclose = GUICtrlCreateButton("close", $iHeight*5, 0, $iHeight, $iHeight, $BS_ICON)
	GUICtrlSetImage(-1, "shell32.dll", 28)
	GUISetState(@SW_HIDE)

	Return $hGUI
EndFunc

Func _Main()
   Local $hGUI = CreateToolBar()

	; Run the GUI until the dialog is closed
	While 1
		If _Inventor_FrameExists() Then
			GUISetState(@SW_SHOW)
		Else
			GUISetState(@SW_HIDE)
		EndIf

		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
			   ExitLoop
			Case $idButton1
			   RunInventorMacro("OpenDrawing", "OpenDrawing")
			Case $idButton2
			   RunInventorMacro("Quantity", "Quantity")
			Case $idButton3
				;
			Case $idButton4
				;
			Case $idButton5
				;
			Case $idButtonclose
				ExitLoop
			Case Else
		EndSwitch

		ToolBarMove($hGUI)
	Wend

	GUIDelete()
EndFunc   ;==>_Main

_Main()
