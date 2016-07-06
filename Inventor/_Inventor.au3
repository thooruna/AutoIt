#include <Array.au3>
#include <File.au3>

Opt("WinTitleMatchMode", 2)

Global $g_sInventorExe = "Inventor.exe"
Global $g_sInventorTitle = "Autodesk Inventor"
Global $g_sInventorFrame = "[CLASS:AfxFrameOrView90u; INSTANCE:1],[CLASS:AfxFrameOrView110u; INSTANCE:2]"
Global $g_oInventor = Null

Func _Inventor_Active()
	Return WinActive($g_sInventorTitle)
EndFunc

Func _Inventor_Exists()
	$bReturn = ProcessExists($g_sInventorExe)
	If $bReturn Then
		If Not($g_oInventor) Then
			$g_oInventor = ObjGet("", "Inventor.Application")
		EndIf
	Else
		$g_oInventor = Null
	EndIf

	Return $bReturn ;And WinExists($g_sInventorTitle)
EndFunc

Func _Inventor_FrameActive()
	Local $sControl = ControlGetFocus("[ACTIVE]")

	If $sControl = "AfxFrameOrView110u2" Then
		Return True
	Else
		Return False
	EndIf
EndFunc

Func _Inventor_FrameExists($sWindowName)
	$aPosF = _Inventor_FrameGetPos()
	Local $sTitle = WinGetTitle("[ACTIVE]")
	Return ((StringLeft($sTitle, StringLen($sWindowName)) = $sWindowName) Or (StringLeft($sTitle, StringLen($g_sInventorTitle)) = $g_sInventorTitle)) And IsArray($aPosF)
EndFunc

Func _Inventor_FrameGetPos()
	Local $aPosF
	Local $aFrame = StringSplit($g_sInventorFrame, ",")
	For $vFrame In $aFrame
		$aPosF = ControlGetPos($g_sInventorTitle, "", $vFrame)
		If IsArray($aPosF) Then
			Return $aPosF
		EndIf
    Next
EndFunc

Func _Inventor_GetPos()
	Local $aPosI = WinGetPos($g_sInventorTitle, "")
	Return $aPosI
EndFunc

Func _Inventor_MacroExecute($sModule, $sMember)
	If _Inventor_Exists() Then
		Local $oModule = $g_oInventor.VBAProjects.Item(1).InventorVBAComponents.Item($sModule)
		Local $oMember = $oModule.InventorVBAMembers.Item($sMember)
		Local $oResult = $oMember.Execute
	EndIf
EndFunc

Func _Inventor_MacroArray($sModule)
	Local $aButtons[0]

	If _Inventor_Exists() Then
		Local $oModule = $g_oInventor.VBAProjects.Item(1).InventorVBAComponents.Item($sModule)

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

		Local $g_oInventor = ObjGet("", "Inventor.Application")
		_PathSplit($g_oInventor.FileOptions.DefaultVBAProjectFileFullFilename, $sDrive, $sDir, $sFileName, $sExtension)

		Return $sDrive & $sDir
	EndIf
EndFunc


