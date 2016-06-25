#include <Array.au3>
#include <File.au3>

Opt("WinTitleMatchMode", 2)

Global $sInventorExe = "Inventor.exe"
Global $sInventorTitle = "Autodesk Inventor"
Global $sInventorFrame = "[CLASS:AfxFrameOrView90u; INSTANCE:1],[CLASS:AfxFrameOrView110u; INSTANCE:2]"

Func _Inventor_Active()
	Return WinActive($sInventorTitle)
EndFunc

Func _Inventor_Exists()
	Return ProcessExists($sInventorExe) ;And WinExists($sInventorTitle)
EndFunc

Func _Inventor_FrameExists($sWindowName)
	$aPosF = _Inventor_FrameGetPos()
	Local $sTitle = WinGetTitle("[ACTIVE]")
	Return ((StringLeft($sTitle, StringLen($sWindowName)) = $sWindowName) Or (StringLeft($sTitle, StringLen($sInventorTitle)) = $sInventorTitle)) And IsArray($aPosF)
EndFunc

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

Func _Inventor_GetPos()
	Local $aPosI = WinGetPos($sInventorTitle, "")
	Return $aPosI
EndFunc

Func _Inventor_MacroExecute($sModule, $sMember)
	If _Inventor_Exists() Then
		Local $oInventor = ObjGet("", "Inventor.Application")
		Local $oModule = $oInventor.VBAProjects.Item(1).InventorVBAComponents.Item($sModule)
		Local $oMember = $oModule.InventorVBAMembers.Item($sMember)
		$oMember.Execute
	EndIf
EndFunc

Func _Inventor_MacroArray($sModule)
	Local $aButtons[0]

	If _Inventor_Exists() Then
		Local $oInventor = ObjGet("", "Inventor.Application")
		Local $oModule = $oInventor.VBAProjects.Item(1).InventorVBAComponents.Item($sModule)

		Local $vMember
		For $vMember In $oModule.InventorVBAMembers
			_ArrayAdd($aButtons, $vMember.Name)
		Next
	EndIf

	Return $aButtons
EndFunc

Func _Inventor_VBAProjectPath()
	If _Inventor_Exists() Then
		Local $sDrive = ""
		Local $sDir = ""
		Local $sFileName = ""
		Local $sExtension = ""

		Local $oInventor = ObjGet("", "Inventor.Application")
		_PathSplit($oInventor.FileOptions.DefaultVBAProjectFileFullFilename, $sDrive, $sDir, $sFileName, $sExtension)

		Return $sDrive & $sDir
	EndIf
EndFunc


