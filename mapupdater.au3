#NoTrayIcon
; ===============================================================================================================
;
; Description:      Download and install maps created by openfietsmap.nl
; Author:           Andrwe Lord Weber
; Date:             2013-07-07
; Licence:          Creative Common Attribution-ShareAlike 3.0 Unported (http://creativecommons.org/licenses/by-sa/3.0/)
; Credits:          (_Zip.au3) PsaltyDS for the original idea on which this UDF is based.
;                   (_Zip.au3) torels for the basic framework on which this UDF is based.
;                   (_Zip.au3) wraithdu for the Zip-UDF.
; Sources:          AutoIt3 -> http://www.autoitscript.com/
;                   _Zip.au3 -> http://www.autoitscript.com/forum/topic/116565-zip-udf-zipfldrdll-library/
;
; NOTES:
;   This script is mainly for downloading maps of openfietsmap.nl but could also work for other map sources.
;   If you want to give it a shoot just change the URL ($sUrl) in function _getFiles to the page where the maps
;   in zip-format can be downloaded.
;
;   I've tested and written this script for Garmin Dakota 20.
;   It should work with every navi supporting openstreetmap maps and providing a drive where the maps can be
;   copied to.
;
; ===============================================================================================================

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ListviewConstants.au3>
#include <EditConstants.au3>
#include <ProgressConstants.au3>
#include <String.au3>
#include <File.au3>
#include <_Zip.au3>
#include <Array.au3>
 
; #FUNCTION# ====================================================================================================
; Name...........:  _isInternetConnected
; Description....:  check if internet connection is available
; Syntax.........:  _isInternetConnected()
; Parameters.....:  
; Return values..:  Success     - 1
;                   Failure     - 0
; Author.........:  guinness
; Modified.......:
; Remarks........:  
; Related........:
; Link...........:  http://www.autoitscript.com/wiki/AutoIt_Snippets
; Example........:
; ===============================================================================================================
 Func _isInternetConnected()
    Local $aReturn = DllCall('connect.dll', 'long', 'IsInternetConnected')
    If @error Then
        Return SetError(1, 0, False)
    EndIf
    Return $aReturn[0] = 0
 EndFunc
 
; #FUNCTION# ====================================================================================================
; Name...........:  _moveFile
; Description....:  move given file to target directory using Windows API with progress bar
; Syntax.........:  _moveFile($fromFile, $toFile)
; Parameters.....:  
; Return values..:  Success     - 1
;                   Failure     - 0
; Author.........:  Andrwe Lord Weber
; Credits:.......:  Jos (http://www.autoitscript.com/forum/user/19-jos) for original idea (_FileCopy)
; Modified.......:  
; Remarks........:  
; Related........:
; Link...........:  http://www.autoitscript.com/wiki/Snippets_(_Files_%26_Folders_)
; Example........:
; ===============================================================================================================
 Func _moveFile($fromFile,$tofile)
    Local $FOF_RESPOND_YES = 16
    $winShell = ObjCreate("shell.application")
    Return $winShell.namespace($tofile).MoveHere($fromFile,$FOF_RESPOND_YES)
 EndFunc
 
 Func _addStatusText($StatusEdit, $sText)
	GUICtrlSetData($StatusEdit, GUICtrlRead($StatusEdit) & '- ' & $sText & @CRLF)
 EndFunc
 
 Func _addElement(ByRef $aArray, $vValue = '')
	If Not IsArray($aArray) Then
		  $aArray = _ArrayCreate($vValue)
	   Else
		  _ArrayAdd($aArray, $vValue)
	   EndIf
 EndFunc
 
 Func _loadFile($StatusEdit, $StatusProgress, $StatusLabel, $EndButton, $sUrl, $sFile)
	Local $download, $dlSize = 2000000000, $dlDone = 0, $dlDoneOld = 0, $dlRate = 1, $dlRateOld
	GUICtrlSetData($EndButton, 'Download stoppen')
    GUICtrlSetData($StatusProgress, 1)
	$sUrl = $sUrl & $sFile
	_addStatusText($StatusEdit, 'Lade Datei ' & $sUrl & ' nach ' & @TempDir)
	   $download = InetGet($sUrl, @TempDir & "\" & $sFile, 2, 1)
	   $dlSize = InetGetSize($sUrl, 2)
	   Do
		  If GUIGetMsg() == $EndButton Or GUIGetMsg() == $GUI_EVENT_CLOSE Then
			 InetClose($download)
			 FileDelete(@TempDir & "\" & $sFile)
			 GUICtrlSetData($StatusProgress, 0)
			 GUICtrlSetData($StatusLabel, '')
			 GUICtrlSetData($EndButton, 'Beenden')
			 Return 1
		  EndIf
		  $dlDone = InetGetInfo($download, 0)
		  $dlRateOld = $dlRate
		  $dlRate = Floor((($dlDone - $dlDoneOld) * 10 / 1024))
		  If $dlRate < 1 Then $dlRate = $dlRateOld
		  GUICtrlSetData($StatusLabel, 'Datei-Größe (MB):' & @TAB & Floor(($dlSize / 1024 / 1024)) & @TAB & 'ca. Restzeit (Min): ' & @TAB & Floor($dlSize / 1024 / $dlRate / 60) & @CRLF & 'davon geladen (MB):' & @TAB & Floor(($dlDone / 1024 / 1024)) & @TAB & 'Download-Rate (KB/s): ' & @TAB & $dlRate)
		  GUICtrlSetData($StatusProgress, ($dlDone * 100 / $dlSize))
		  $dlDoneOld = $dlDone
		  Sleep(100)
	   Until InetGetInfo($download, 2)
	   If InetGetInfo($download, 3) Then
		  GUICtrlSetData($StatusProgress, 100)
		  Return 1
	   Else
		  _addStatusText($StatusEdit, 'Beim herunterladen von ' & $aUrl & ' ist folgender Fehler aufgetreten:' & @CRLF & InetGetInfo($download, 4)
		  Return 0
	   EndIf
 EndFunc
 
 Func _getFiles($WinMain, $StatusEdit, $StatusProgress, $StatusLabel, $EndButton)
	Local $WinError, $Label, $TextArea, $sNewFile, $sNewFilePath, $sFile, $sFilePath, $dlReturn = 1, $sUrl = "http://osm.pleiades.uni-wuppertal.de/openfietsmap/EU_2013/GPS/"   ; <------- Change if you want to try it
	Local $WinCardSel, $i, $lvItem, $aItem
	Local $aFiles, $aLvItems, $aLines
	_addStatusText($StatusEdit, 'Ermittle Datei-URLs')
	Local $bData = InetRead($sUrl)
	Dim $aContent = StringRegExp(BinaryToString($bData), '(?i).*a\shref="([^"]*\.zip)".*<td[^>]*>([^<]*)</td><td[^<]*>[^<]*</td><td[^>]*><?b?>?([^<]*)<.*', 3)
	For $i = 0 To UBound($aContent) - 1  Step 3
	   _addElement($aLines, $aContent[$i] & '|' & $aContent[$i + 1] & '|' & $aContent[$i + 2])
	Next
	
	$WinCardSel = GUICreate("Karten wählen", 500, 280, -1, -1, $WS_MINIMIZEBOX + $WS_CAPTION + $WS_POPUP + $WS_SYSMENU + $WS_SIZEBOX, -1, $WinMain)
	GUICtrlCreateLabel("Folgende Karten wurden gefunden." & @CRLF & "Bitte auswählen, welche installiert werden sollen:", 5, 5)
	Local $list = GUICtrlCreateListView("|Datei                     |Änderungsdatum|Beschreibung", 5, 5, 490, 230, $GUI_SS_DEFAULT_LISTVIEW + $LVS_NOSORTHEADER, $LVS_EX_CHECKBOXES)
	For $sLine In $aLines
	   _addElement($aLvItems, GUICtrlCreateListViewItem('|' & $sLine, $list))
	Next
	Local $ChooseButton = GUICtrlCreateButton("Auswahl installieren", 110, 240, 105, 35)
	Local $AllButton = GUICtrlCreateButton("Alle Installieren", 225, 240, 105, 35)
	GUISetState()
	
	While 1
	   Switch GUIGetMsg()
		  Case $ChooseButton
			 For $lvItem In $aLvItems
				If BitAnd(GUICtrlRead($lvItem, 1),$GUI_CHECKED) Then
				   $aItem = StringSplit(GUICtrlRead($lvItem), '|')
				   _addElement($aFiles, $aItem[2])
				EndIf
			 Next
			 ExitLoop
		  Case $AllButton
			 For $lvItem In $aLvItems
				$aItem = StringSplit(GUICtrlRead($lvItem), '|')
				_addElement($aFiles, $aItem[2])
			 Next
			 ExitLoop
		  Case $GUI_EVENT_CLOSE
			 GUIDelete($WinCardSel)
			 _addStatusText($StatusEdit, 'Es wurden keine Karten zur Installation ausgewählt.')
			 Return 1
	   EndSwitch
	WEnd
	GUIDelete($WinCardSel)
	
	Dim $aImages[UBound($aFiles)]
	_addStatusText($StatusEdit, 'Lade und installiere ' & UBound($aFiles) & ' Karte(n).')
	
	if not IsArray($aFiles) Then
	   $WinError = GUICreate("Error", 600, 800, -1, -1, $WS_MINIMIZEBOX + $WS_CAPTION + $WS_POPUP + $WS_SYSMENU + $WS_SIZEBOX)
	   $Label = GUICtrlCreateLabel("Konnte Karten-URLs nicht ermitteln. " & @CRLF & "Gefundene URLs: " & UBound($aFiles), 5, 5)
	   $TextArea = GUICtrlCreateEdit(BinaryToString($bData), 5, 40, 590, 755, $ES_AUTOVSCROLL + $WS_VSCROLL + $ES_MULTILINE + $ES_READONLY)
	   GUISetState()
	   Return 1
    EndIf
	For $sFile In $aFiles
	   $sFilePath = @TempDir & '\' & $sFile
	   If GUIGetMsg() == $EndButton Then
		  InetClose($download)
		  FileDelete($sFilePath)
		  GUICtrlSetData($StatusProgress, 0)
		  GUICtrlSetData($StatusLabel, '')
		  GUICtrlSetData($EndButton, 'Beenden')
		  Return 1
	   EndIf
	   $sNewFile = StringReplace(StringLower($sFile), '.zip', '.img')
	   $sNewFilePath = @TempDir & '\' & $sNewFile
	   If FileExists($sNewFilePath) Then
		  If MsgBox(292, 'Karte existiert', 'Die Datei ' & $sNewFilePath & ' existiert bereits.' & @CRLF & 'Überschreiben?') == 6 Then
			 $dlReturn = _loadFile($StatusEdit, $StatusProgress, $StatusLabel, $EndButton, $sUrl, $sFile)
		  Else
			 $dlReturn = 2
		  EndIf
	   Else
		  $dlReturn = _loadFile($StatusEdit, $StatusProgress, $StatusLabel, $EndButton, $sUrl, $sFile)
	   EndIf
	   
	   If $dlReturn == 1 Then
		  If _unzipFiles($StatusEdit, $EndButton, $sFilePath, $sNewFilePath) Then
			 _addStatusText($StatusEdit, 'Lösche erfolgreich entpackte Datei ' & $sFilePath)
			 FileDelete($sFilePath)
			 _ArrayAdd($aImages, $sNewFilePath)
		  EndIf
	   ElseIf $dlReturn == 2 Then
		  _ArrayAdd($aImages, $sNewFilePath)
	   EndIf
    Next
	GUICtrlSetData($StatusLabel, '')
	GUICtrlSetData($EndButton, 'Beenden')
	Return $aImages
 EndFunc
 
 Func _unzipFiles($StatusEdit, $EndButton, $sFilePath, $sNewFilePath)
	Local $szDrive, $szDir, $szFName, $szExt, $aPath
    GUICtrlSetData($EndButton, 'Entpacken stoppen')
	If Not FileExists($sFilePath) Then
	   Return 0
    EndIf
    If _Zip_ItemExists($sFilePath, 'garmin\gmapsupp.img') Then
	   _addStatusText($StatusEdit, 'Entpacke ' & $sFilePath & ' nach ' & $sNewFilePath)
	   if _Zip_Unzip($sFilePath, 'garmin\gmapsupp.img', @TempDir, 513) Then
		  If Not FileExists(@TempDir & '\gmapsupp.img') Then
			 _addStatusText($StatusEdit, 'Die Datei ' & @TempDir & '\gmapsupp.img' & ' existiert nicht.')
			 Return 0
		  EndIf
		  If Not FileMove(@TempDir & '\gmapsupp.img', $sNewFilePath, 9) Then
			 _addStatusText($StatusEdit, 'Fehler beim Verschieben der Datei ' & @TempDir & '\gmapsupp.img' & ' nach ' & $sNewFilePath)
			 Return 0
		  EndIf
		  Return 1
	   Else
		  _addStatusText($StatusEdit, 'Beim Entpacken von ' & $sFile & ' ist ein Fehler mit folgendem Code aufgetreten:' & @error)
		  Return 0
	   EndIf
    EndIf
 EndFunc

 Func _getDrive($WinMain, $StatusEdit)
	Local $list, $ChooseButton, $CancelButton, $aDrive, $sDrive, $aDrives, $lvItem, $WinDriveSel
	Local $aDrives = DriveGetDrive('REMOVABLE')

    If UBound($aDrives) > 0 Then
	   $WinDriveSel = GUICreate("vorhandene Speichermedien", 500, 280, -1, -1, $WS_MINIMIZEBOX + $WS_CAPTION + $WS_POPUP + $WS_SYSMENU + $WS_SIZEBOX, -1, $WinMain)

	   $list = GUICtrlCreateListView("Laufwerk  | Größe (GB) | Garmin-Gerät?  ", 5, 5, 490, 230)
	   $ChooseButton = GUICtrlCreateButton("Auswahl bestätigen", 110, 240, 105, 35)
	   $CancelButton = GUICtrlCreateButton("Abbrechen", 225, 240, 105, 35)
	   for $sDrive in $aDrives
		  If DriveStatus($sDrive) == 'READY' Then
			 $lvItem = $sDrive & '|' & (DriveSpaceTotal($sDrive)/1024) & '|'
			 If FileExists($sDrive & '\garmin') Then
				$lvItem = $lvItem & 'yes'
			 EndIf
			 GUICtrlCreateListViewItem($lvItem, $list)
		  EndIf
	   Next
	   GUISetState()

	   While 1
		  Switch GUIGetMsg()
			 Case $ChooseButton
				$aDrive = StringSplit(GUICtrlRead(GUICtrlRead($list)), '|')
				GUIDelete($WinDriveSel)
				Return $aDrive
			 Case $CancelButton
				_addStatusText($StatusEdit, 'Karten-Update wird abgebrochen, da kein valides Gerät ausgewählt wurde.')
				GUIDelete($WinDriveSel)
				Return 0
			 Case $GUI_EVENT_CLOSE
				_addStatusText($StatusEdit, 'Karten-Update wird abgebrochen, da kein valides Gerät ausgewählt wurde.')
				GUIDelete($WinDriveSel)
				Return 0
		  EndSwitch
	   WEnd
	Else
	   Do
		  If MsgBox(5, 'Warte auf Speichermedium', 'Bitte das Navigationsgerät anschließen.') == 2 Then
			 _addStatusText($StatusEdit, 'Karten-Update wird abgebrochen, da kein valides Gerät angeschlossen wurde.')
			 Return 0
		  EndIf
	   Until UBound(DriveGetDrive('REMOVABLE')) > 0
	   Return _getDrive($WinMain, $StatusEdit)
    EndIf
 EndFunc

 Func main()
    Local $WinMain, $StartButton, $StatusEdit, $StatusProgress, $StatusLabel, $EndButton, $aImages, $sImage, $sTargetDir, $aDrive
	$WinMain = GUICreate('Kartenupdate-Tool', 400, 440, -1, -1, $WS_MINIMIZEBOX + $WS_CAPTION + $WS_POPUP + $WS_SYSMENU + $WS_SIZEBOX)
    $StatusEdit = GUICtrlCreateEdit('', 5, 5, 390, 330, $ES_READONLY + $ES_MULTILINE + $ES_AUTOVSCROLL + $WS_VSCROLL)
	$StatusProgress = GUICtrlCreateProgress(5, 335, 390, 20, $PBS_SMOOTH)
	$StatusLabel = GUICtrlCreateLabel('', 5, 360, 390, 35)
	$StartButton = GUICtrlCreateButton('Update starten', 110, 400, 105, 35)
	$EndButton = GUICtrlCreateButton('Beenden', 225, 400, 105, 35)
	GUISetState()
	While 1
		Switch GUIGetMsg()
			 Case $GUI_EVENT_CLOSE
				ExitLoop
			 Case $StartButton
				while not _isInternetConnected()
				   GUISetState(@SW_DISABLE, $WinMain)
				   Switch MsgBox( 6, "keine Internetverbindung", "Der Rechner scheint keine Netzwerkverbindung zu haben. Bitte zu erst ein Verbindung herstellen.")
					  Case 2
						 Exit
					  Case 11
						 ExitLoop
				   EndSwitch
			    WEnd
				GUISetState(@SW_ENABLE, $WinMain)
				GUISetState(@SW_SHOWNORMAL, $WinMain)
				$aImages = _getFiles($WinMain, $StatusEdit, $StatusProgress, $StatusLabel, $EndButton)
				If IsArray($aImages) And UBound($aImages) > 0 Then
				   $aDrive = _getDrive($WinMain, $StatusEdit)
				   If IsArray($aDrive) And UBound($aDrive) > 0 Then
					  If $aDrive[3] == 'yes' Then
						 _addStatusText($StatusEdit, 'Ausgewähltes Laufwerk wurde als Garmin-Gerät erkannt.' & @CRLF & '  Nutze Standard-Ordner "garmin" als Ziel.')
						 $sTargetDir = $aDrive[1] & '\garmin\'
					  Else
						 $sTargetDir = FileSelectFolder('Bitte Zielpfad auf dem Laufwerk auswählen.', $aDrive[1])
					  EndIf
					  For $sImage In $aImages
						 If Not $sImage Then ContinueLoop
						 _addStatusText($StatusEdit, 'Verschiebe ' & $sImage & ' nach ' & $sTargetDir)
						 If Not _moveFile($sImage, $sTargetDir) Then
							_addStatusText($StatusEdit, 'Fehler beim Verschieben der Karte auf das Gerät. Fehler-Code: ' & @error)
						 EndIf
					  Next
					  _addStatusText($StatusEdit, 'Kartenupdate abgeschlossen.')
				   EndIf
			    EndIf
			 Case $EndButton
				ExitLoop
		EndSwitch
	WEnd
	
 EndFunc
 
 main()
