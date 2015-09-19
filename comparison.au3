#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.1
 Author:         Alex Kamuro

 Script Function:
	  Comparison GUI and functions

#ce ----------------------------------------------------------------------------

#include <Array.au3>
#include <File.au3>
#include <Crypt.au3>
#include <GuiListView.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ProgressConstants.au3>
#include <FileConstants.au3>

Global $hCompGUI = GUICreate("Comparison result", 700, 400, -1, -1, BitOR($WS_MINIMIZEBOX,$WS_MAXIMIZEBOX,$WS_SIZEBOX))
   Local $hPathLabel = GUICtrlCreateLabel("Path:", 10, 10, 580)
   GUICtrlSetResizing(-1, $GUI_DOCKTOP + $GUI_DOCKBOTTOM + $GUI_DOCKLEFT)
   Local $hCountLabel = GUICtrlCreateLabel("Count: ", 10, 30, 100)
   GUICtrlSetResizing(-1, $GUI_DOCKALL)
   Local $hProgress = GUICtrlCreateProgress(150, 32, 300, 10, $PBS_MARQUEE)
   GUICtrlSetState($hProgress, $GUI_HIDE)
   GUICtrlSetResizing($hProgress, $GUI_DOCKALL)
   Local $hCompList = GUICtrlCreateListView("Type|Path", 10, 55, 680, 310, $LVS_REPORT)
   GUICtrlSetResizing(-1, $GUI_DOCKBORDERS)
   _GUICtrlListView_SetColumnWidth($hCompList, 0, 100)
   Local $hCopyBtn = GUICtrlCreateButton("Copy to folder...", 580, 15, 100, 25)
   GUICtrlSetTip($hCopyBtn, "Copy the modified and new files to another folder")
   GUICtrlSetResizing($hCopyBtn, $GUI_DOCKTOP + $GUI_DOCKRIGHT)

Local $iNewDirLength = 0

Func CompareStart($sOldDir, $sNewDir)
   $iNewDirLength = StringLen($sNewDir)
   GUICtrlSetData($hPathLabel, "Path: " & $sNewDir)
   _GUICtrlListView_DeleteAllItems($hCompList)

   _Crypt_Startup()
   GUICtrlSetState($hCopyBtn, $GUI_HIDE)
   GUICtrlSetState($hProgress, $GUI_SHOW)
   GUICtrlSendMsg($hProgress, $PBM_SETMARQUEE, 1, 100)
   GUISetState(@SW_SHOW, $hCompGUI)
   WinActivate($hCompGUI)
   Compare($sOldDir, $sNewDir)
   GUICtrlSetState($hProgress, $GUI_HIDE)
   GUICtrlSendMsg($hProgress, $PBM_SETMARQUEE, 0, 0)
   If _GUICtrlListView_GetItemCount($hCompList) > 0 Then
	  GUICtrlSetState($hCopyBtn, $GUI_SHOW)
   EndIf
   _Crypt_Shutdown()
   MsgBox($MB_ICONINFORMATION, "", "Done.", 0, $hCompGUI)
EndFunc		;==>CompareStart

Func CompareLists($aOldList, $aNewList)
   Local $aFiles = []

   For $i = 1 To UBound($aOldList)-1
	  $search = _ArraySearch($aNewList,$aOldList[$i],1)
	  If $search == -1 Then _ArrayAdd($aFiles, $aOldList[$i])
   Next

   For $i = 1 To UBound($aNewList)-1
	  $search = _ArraySearch($aOldList,$aNewList[$i],1)
	  If $search == -1 Then _ArrayAdd($aFiles, $aNewList[$i])
   Next

   Return $aFiles
EndFunc		;==>CompareLists

Func CompareFiles($aOldList, $aNewList, $sOldDir, $sNewDir)
   Local $aFiles = []

   For $i = 1 To UBound($aOldList)-1
	  If Not FileExists($sOldDir & '\' & $aOldList[$i]) Or Not FileExists($sNewDir & '\' & $aOldList[$i]) Then ContinueLoop

	  Local $dHash1 = _Crypt_HashFile($sOldDir & '\' & $aOldList[$i], $CALG_SHA1)
	  Local $dHash2 = _Crypt_HashFile($sNewDir & '\' & $aOldList[$i], $CALG_SHA1)
	  If $dHash1 <> $dHash2 Then _ArrayAdd($aFiles, $aOldList[$i])
   Next

   Return $aFiles
EndFunc		;==>CompareFiles

Func Compare($sOldDir, $sNewDir)

   Local $aOldFiles = _FileListToArray($sOldDir)
   Local $aNewFiles = _FileListToArray($sNewDir)

   ; search for new or deleted files
   Local $aFilesDiff = CompareLists($aOldFiles, $aNewFiles)
   For $i = 1 To UBound($aFilesDiff)-1
	  If FileExists($sNewDir & '\' & $aFilesDiff[$i]) Then
		 GUICtrlCreateListViewItem("Added|" & StringTrimLeft($sNewDir, $iNewDirLength) & '\' & $aFilesDiff[$i], $hCompList)
		 GUICtrlSetColor(-1, 0x006600)
		 _GUICtrlListView_SetColumnWidth($hCompList, 1, $LVSCW_AUTOSIZE)
		 GUICtrlSetData($hCountLabel, "Count: " & _GUICtrlListView_GetItemCount($hCompList))
	  Else
		 GUICtrlCreateListViewItem("Removed|" & StringTrimLeft($sNewDir, $iNewDirLength) & '\' & $aFilesDiff[$i], $hCompList)
		 GUICtrlSetColor(-1, 0xCC0000)
		 _GUICtrlListView_SetColumnWidth($hCompList, 1, $LVSCW_AUTOSIZE)
		 GUICtrlSetData($hCountLabel, "Count: " & _GUICtrlListView_GetItemCount($hCompList))
	  EndIf
   Next

   ; comparison of existing files
   Local $aFilesDiff = CompareFiles($aOldFiles, $aNewFiles, $sOldDir, $sNewDir)
   For $i = 1 To UBound($aFilesDiff)-1
	  GUICtrlCreateListViewItem("Changed|" & StringTrimLeft($sNewDir, $iNewDirLength) & '\' & $aFilesDiff[$i], $hCompList)
	  GUICtrlSetColor(-1, 0x3333CC)
	  _GUICtrlListView_SetColumnWidth($hCompList, 1, $LVSCW_AUTOSIZE)
	  GUICtrlSetData($hCountLabel, "Count: " & _GUICtrlListView_GetItemCount($hCompList))
   Next

   ; comparison in subfolders
   For $i = 1 To UBound($aOldFiles)-1
	  Local $sDir1 = $sOldDir & '\' & $aOldFiles[$i]
	  Local $sDir2 = $sNewDir & '\' & $aOldFiles[$i]
	  If Not IsDir($sDir1) Then ContinueLoop
	  If FileExists($sDir2) And IsDir($sDir2) Then
		 Compare($sDir1, $sDir2)
	  EndIf
   Next

EndFunc		;==>Compare

Func CopyToFolder($sSourseDir)
   Local $sPath = FileSelectFolder("Select Folder", "")
   If $sPath = "" Then Return

   GUICtrlSetState($hCopyBtn, $GUI_HIDE)
   GUICtrlSetState($hProgress, $GUI_SHOW)
   GUICtrlSendMsg($hProgress, $PBM_SETMARQUEE, 1, 100)

   For $i=0 To _GUICtrlListView_GetItemCount($hCompList)
	  Local $sType = _GUICtrlListView_GetItemText($hCompList, $i, 0)
	  If $sType <> "Added" And $sType <> "Changed" Then ContinueLoop

	  Local $sFile = _GUICtrlListView_GetItemText($hCompList, $i, 1)
	  If IsDir($sSourseDir & $sFile) Then ContinueLoop

	  FileCopy($sSourseDir & $sFile, $sPath & $sFile, $FC_OVERWRITE + $FC_CREATEPATH)
   Next

   GUICtrlSetState($hProgress, $GUI_HIDE)
   GUICtrlSendMsg($hProgress, $PBM_SETMARQUEE, 0, 0)
   GUICtrlSetState($hCopyBtn, $GUI_SHOW)

   MsgBox($MB_ICONINFORMATION, "", "Copying to '" & $sPath & "\' completed.", 0, $hCompGUI)
EndFunc		;==>CopyToFolder

Func IsDir($sFilePath)
   Return StringInStr(FileGetAttrib($sFilePath), "D") > 0
EndFunc		;==>IsDir