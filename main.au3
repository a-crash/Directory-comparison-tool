#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.1
 Author:         Alex Kamuro

 Script Function:
	  Main GUI

#ce ----------------------------------------------------------------------------

#NoTrayIcon
#include <GUIConstantsEx.au3>
#include <EditConstants.au3>
#include "comparison.au3"

Local $hMainGUI = GUICreate("Directory comparison tool", 410, 140)
   GUICtrlCreateLabel("Original folder:", 10, 20, 100, 25)
   Local $hCompDir1 = GUICtrlCreateInput("", 110, 20, 200, 20, $ES_READONLY)
   Local $hCompDirButton1 = GUICtrlCreateButton("Browse", 320, 20, 75, 20)
   GUICtrlCreateLabel("New folder:", 10, 50, 100, 25)
   Local $hCompDir2 = GUICtrlCreateInput("", 110, 50, 200, 20, $ES_READONLY)
   Local $hCompDirButton2 = GUICtrlCreateButton("Browse", 320, 50, 75, 20)
   Local $hCompStartBtn = GUICtrlCreateButton("Compare", 145, 85, 120, 40)
   GUICtrlSetState(-1, $GUI_DISABLE)

GUISetState(@SW_SHOW, $hMainGUI)

While 1
   Local $aMsg = GUIGetMsg(1)
   Switch $aMsg[1]
   Case $hMainGUI
	  Switch $aMsg[0]
		 Case $GUI_EVENT_CLOSE
			Exit
		 Case $hCompDirButton1
			Local $sPath = FileSelectFolder("Select Folder", "")
			GUICtrlSetData($hCompDir1,$sPath)
			CheckButton()
		 Case $hCompDirButton2
			Local $sPath = FileSelectFolder("Select Folder", "")
			GUICtrlSetData($hCompDir2,$sPath)
			CheckButton()
		 Case $hCompStartBtn
			Local $sOldDir = GUICtrlRead($hCompDir1)
			Local $sNewDir = GUICtrlRead($hCompDir2)
			CompareStart($sOldDir, $sNewDir)
	  EndSwitch
   Case $hCompGUI
	  Switch $aMsg[0]
		 Case $GUI_EVENT_CLOSE
			GUISetState(@SW_HIDE, $hCompGUI)
		 Case $hCopyBtn
			CopyToFolder(GUICtrlRead($hCompDir2))
	  EndSwitch
   EndSwitch
WEnd

Func CheckButton()
   If GUICtrlRead($hCompDir1) <> "" And GUICtrlRead($hCompDir2) <> "" Then
	  GUICtrlSetState($hCompStartBtn, $GUI_ENABLE)
   Else
	  GUICtrlSetState($hCompStartBtn, $GUI_DISABLE)
   EndIf
EndFunc		;==>CheckButton
