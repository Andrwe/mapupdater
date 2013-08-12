#NoTrayIcon
; ===============================================================================================================
;
; Description:      Download and install maps created by openfietsmap.nl
; Author:           Andrwe Lord Weber
; Date:             2013-08-12
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

Const $VERSION = '0.0.1'

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
; Return values..:  Success          - 1
;                   Unknown Language - 0
;                   Failure          - 0
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
 
 ; #FUNCTION# ====================================================================================================
; Name...........:  _getLanguage
; Description....:  returns language code according to ISO 639-3 used in lang.ini
; Syntax.........:  _getLanguage()
; Parameters.....:  
; Return values..:  Success     - language code
;                   Failure     - 0
; Author.........:  Andrwe Lord Weber
; Credits:.......:  http://www.autoitscript.com/autoit3/docs/appendix/OSLangCodes.htm
; Modified.......:  
; Remarks........:  
; Related........:  http://www-01.sil.org/iso639%2D3/codes.asp
; Link...........:  http://www.autoitscript.com/autoit3/docs/appendix/OSLangCodes.htm
; Example........:
; ===============================================================================================================
 Func _getLanguage()
    If _getConfig('general', 'LANG', '') Then Return _getConfig('general', 'LANG', '')
	Select
	   Case StringInStr("0413 0813", @OSLang)
		   ; Dutch
		   Return 0
	   Case StringInStr("0409 0809 0c09 1009 1409 1809 1c09 2009 2409 2809 2c09 3009 3409", @OSLang)
		   ; English
		   Return "eng"
	   Case StringInStr("040c 080c 0c0c 100c 140c 180c", @OSLang)
		   ; French
		   Return 0
	   Case StringInStr("0407 0807 0c07 1007 1407", @OSLang)
		   ; German
		   Return "deu"
	   Case StringInStr("0410 0810", @OSLang)
		   ; Italian
		   Return 0
	   Case StringInStr("0414 0814", @OSLang)
		   ; Norwegian
		   Return 0
	   Case StringInStr("0415", @OSLang)
		   ; Polish
		   Return 0
	   Case StringInStr("0416 0816", @OSLang)
		   ; Portuguese
		   Return 0
	   Case StringInStr("040a 080a 0c0a 100a 140a 180a 1c0a 200a 240a 280a 2c0a 300a 340a 380a 3c0a 400a 440a 480a 4c0a 500a", @OSLang)
		   ; Spanish
		   Return 0
	   Case StringInStr("041d 081d", @OSLang)
		   ; Swedish
		   Return 0
	   Case Else
		   Return 0
    EndSelect
 EndFunc
 
 Func _generateString($sLine, $aReplaces)
	Local $i
	for $i = 0 To UBound($aReplaces) -1
	   $sLine = StringReplace($sLine, '%' & $i & '%', $aReplaces[$i])
	Next
	Return $sLine
 EndFunc
 
 Func _getString($sSection, $sKey, $aReplaces = '')
	Return _generateString(IniRead(@ScriptDir & '\lang.ini', $sSection & '/' & _getLanguage(), $sKey, IniRead(@ScriptDir & '\lang.ini', $sSection & '/eng', $sKey, 'Could not find ' & $sKey & ' in your or default language')), $aReplaces)
 EndFunc
 
 Func _getConfig($sSection, $sKey, $sDefault)
	Local $sFile
	If FileExists(@ScriptDir & '\mapupdater.ini') Then
	   $sFile = @ScriptDir & '\mapupdater.ini'
	EndIf
	If FileExists(@UserProfileDir & '\mapupdater.ini') Then
	   $sFile = @UserProfileDir & '\mapupdater.ini'
	EndIf
	Return IniRead($sFile, $sSection, $sKey, $sDefault)
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
 
 Func _loadFile($StatusEdit, $StatusProgress, $StatusLabel, $EndButton, $sUrl, $sFile, $force = 0)
	Local $download, $dlSize = 2000000000, $dlDone = 0, $dlDoneOld = 0, $dlRate = 1, $dlRateOld
	GUICtrlSetData($EndButton, _getString('button', 'btnStopDownload'))
    GUICtrlSetData($StatusProgress, 1)
	$sUrl = $sUrl & $sFile
	_addStatusText($StatusEdit, _getString('info', 'infDownloadFile', _ArrayCreate($sUrl, @TempDir & '\' & $sFile)))
	$dlSize = InetGetSize($sUrl, 2 + $force)
	$download = InetGet($sUrl, @TempDir & '\' & $sFile, 2 + $force, 1)
	Local $err = InetGetInfo($download)
	_ArrayDisplay($err)
	Do
	   If GUIGetMsg() == $EndButton Or GUIGetMsg() == $GUI_EVENT_CLOSE Then
		  InetClose($download)
		  FileDelete(@TempDir & "\" & $sFile)
		  GUICtrlSetData($StatusProgress, 0)
		  GUICtrlSetData($StatusLabel, '')
		  GUICtrlSetData($EndButton, _getString('button', 'btnClose'))
		  Return 0
	   EndIf
	   $dlDone = InetGetInfo($download, 0)
	   $dlRateOld = $dlRate
	   $dlRate = Floor((($dlDone - $dlDoneOld) * 10 / 1024))
	   If $dlRate < 10 Then $dlRate = $dlRateOld
	   GUICtrlSetData($StatusLabel, _getString('info', 'infDownloadStatus', _ArrayCreate(@TAB & Floor(($dlSize / 1024 / 1024)) & @TAB, @TAB & Floor($dlSize / 1024 / $dlRate / 60) & @CRLF, @TAB & Floor(($dlDone / 1024 / 1024)) & @TAB, @TAB & $dlRate)))
	   GUICtrlSetData($StatusProgress, ($dlDone * 100 / $dlSize))
	   $dlDoneOld = $dlDone
	   Sleep(100)
	Until InetGetInfo($download, 2)
	If InetGetInfo($download, 3) Then
	   GUICtrlSetData($StatusProgress, 100)
	   InetClose($download)
	   Return 1
	Else
	   Local $err = InetGetInfo($download)
; known error-codes:
;  13 - file doesn't exist
;  31 - 
	   If $err[4] == 13 Then
		  _addStatusText($StatusEdit, _getString('error', 'errDownload', _ArrayCreate($sUrl, @CRLF & '   ' & _getString('error', 'errMissingFile', _ArrayCreate('')))))
	   Else
		  _ArrayDisplay($err)
		  _addStatusText($StatusEdit, _getString('error', 'errDownload', _ArrayCreate($sUrl, 'Error-Code: ' & $err & ' + ' & InetGetInfo($download, 5))))
	   EndIf
	   InetClose($download)
	   GUICtrlSetData($StatusProgress, 0)
	   GUICtrlSetData($StatusLabel, '')
	   GUICtrlSetData($EndButton, _getString('button', 'btnClose'))
	   Return 0
	EndIf
 EndFunc
 
 Func _getFiles($WinMain, $StatusEdit, $StatusProgress, $StatusLabel, $EndButton)
	Local $WinError, $Label, $TextArea, $sNewFile, $sNewFilePath, $sFile, $sFilePath, $dlReturn = 1
	Local $sUrl = _getConfig('advanced', 'URL', 'http://osm.pleiades.uni-wuppertal.de/openfietsmap/EU_2013/GPS/')
	Local $WinCardSel, $i, $lvItem, $aItem
	Local $aFiles, $aLvItems, $aLines, $aImages
	_addStatusText($StatusEdit, _getString('info', 'infDetermineUrls'))
	Local $bData = InetRead($sUrl)
	Dim $aContent = StringRegExp(BinaryToString($bData), _getConfig('advanced', 'LINK_REGEXP', '(?i).*a\shref="([^"]*\.zip)".*<td[^>]*>([^<]*)</td><td[^<]*>[^<]*</td><td[^>]*><?b?>?([^<]*)<.*'), 3)
	For $i = 0 To UBound($aContent) - 1  Step 3
	   _addElement($aLines, $aContent[$i] & '|' & $aContent[$i + 1] & '|' & $aContent[$i + 2])
	Next
	
	$WinCardSel = GUICreate(_getString('title', 'ttlChooseCards'), 500, 280, -1, -1, $WS_MINIMIZEBOX + $WS_CAPTION + $WS_POPUP + $WS_SYSMENU + $WS_SIZEBOX, -1, $WinMain)
	GUICtrlCreateLabel(_getString('useract', 'usrChooseCards', _ArrayCreate(@CRLF)), 5, 5)
	Local $list = GUICtrlCreateListView(_getString('header', 'hdrChooseCards'), 5, 5, 490, 230, $GUI_SS_DEFAULT_LISTVIEW + $LVS_NOSORTHEADER, $LVS_EX_CHECKBOXES)
	For $sLine In $aLines
	   _addElement($aLvItems, GUICtrlCreateListViewItem('|' & $sLine, $list))
	Next
	Local $ChooseButton = GUICtrlCreateButton(_getString('button', 'btnInstallChoice'), 110, 240, 105, 35)
	Local $AllButton = GUICtrlCreateButton(_getString('button', 'btnInstallAll'), 225, 240, 105, 35)
	GUISetState()
	
	While 1
	   Switch GUIGetMsg()
		  Case $GUI_EVENT_CLOSE
			 GUIDelete($WinCardSel)
			 _addStatusText($StatusEdit, _getString('error', 'errMissingCards'))
			 Return 1
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
	   EndSwitch
	WEnd
	GUIDelete($WinCardSel)

	_addStatusText($StatusEdit, _getString('info', 'infCountedCards', _ArrayCreate(UBound($aFiles))))
	
	if not IsArray($aFiles) Then
	   $WinError = GUICreate(_getString('error', 'ttlError'), 600, 800, -1, -1, $WS_MINIMIZEBOX + $WS_CAPTION + $WS_POPUP + $WS_SYSMENU + $WS_SIZEBOX)
	   $Label = GUICtrlCreateLabel(_getString('error', 'errDeterminedCards', _ArrayCreate(@CRLF, UBound($aFiles))), 5, 5)
	   $TextArea = GUICtrlCreateEdit(BinaryToString($bData), 5, 40, 590, 755, $ES_AUTOVSCROLL + $WS_VSCROLL + $ES_MULTILINE + $ES_READONLY)
	   GUISetState()
	   Return 1
    EndIf
	For $sFile In $aFiles
	   $sFilePath = @TempDir & '\' & $sFile
	   If GUIGetMsg() == $EndButton Or GUIGetMsg() == $GUI_EVENT_CLOSE Then
		  InetClose($download)
		  FileDelete($sFilePath)
		  GUICtrlSetData($StatusProgress, 0)
		  GUICtrlSetData($StatusLabel, '')
		  GUICtrlSetData($EndButton, _getString('button', 'btnClose'))
		  Return 1
	   EndIf
	   $sNewFile = StringReplace(StringLower($sFile), '.zip', '.img')
	   $sNewFilePath = @TempDir & '\' & $sNewFile
	   If FileExists($sNewFilePath) Then
		  If MsgBox(292, _getString('title', 'ttlExistingCard'), _getString('useract', 'usrOverwriteFile', _ArrayCreate($sNewFilePath, @CRLF))) == 6 Then
			 $dlReturn = _loadFile($StatusEdit, $StatusProgress, $StatusLabel, $EndButton, $sUrl, $sFile)
		  Else
			 $dlReturn = 2
		  EndIf
	   Else
		  $dlReturn = _loadFile($StatusEdit, $StatusProgress, $StatusLabel, $EndButton, $sUrl, $sFile)
	   EndIf
	   
	   If $dlReturn == 1 Then
		  If _unzipFile($StatusEdit, $EndButton, $sFilePath, $sNewFilePath) Then
			 _addStatusText($StatusEdit, _getString('info', 'infDeleteUnzipped', _ArrayCreate($sFilePath)))
			 FileDelete($sFilePath)
			 _addElement($aImages, $sNewFilePath)
		  EndIf
	   ElseIf $dlReturn == 2 Then
		  _addElement($aImages, $sNewFilePath)
	   EndIf
    Next
	GUICtrlSetData($StatusLabel, '')
	GUICtrlSetData($EndButton, _getString('button', 'btnClose'))
	Return $aImages
 EndFunc
 
 Func _unzipFile($StatusEdit, $EndButton, $sFilePath, $sNewFilePath, $bCard = 1)
	Local $szDrive, $szDir, $szFName, $szExt, $aPath
    GUICtrlSetData($EndButton, _getString('button', 'btnStopDecompression'))
	If Not FileExists($sFilePath) Then
	   Return 0
    EndIf
	If $bCard Then
	   If _Zip_ItemExists($sFilePath, 'garmin\gmapsupp.img') Then
		  _addStatusText($StatusEdit, _getString('info', 'infUnzipFile', _ArrayCreate($sFilePath, $sNewFilePath)))
		  if _Zip_Unzip($sFilePath, 'garmin\gmapsupp.img', @TempDir, 513) Then
			 If Not FileExists(@TempDir & '\gmapsupp.img') Then
				_addStatusText($StatusEdit, _getString('error', 'errMissingFile', @TempDir & '\gmapsupp.img'))
				Return 0
			 EndIf
			 If Not FileMove(@TempDir & '\gmapsupp.img', $sNewFilePath, 9) Then
				_addStatusText($StatusEdit, _getString('error', 'errMovingFile', _ArrayCreate(@TempDir & '\gmapsupp.img', $sNewFilePath)))
				Return 0
			 EndIf
			 Return 1
		  Else
			 _addStatusText($StatusEdit, _getString('error', 'errUnzipFile', _ArrayCreate($sFile, @error)))
			 Return 0
		  EndIf
	   EndIf
	Else
	   if _Zip_Unzip($sFilePath, '', $sNewFilePath, 513) Then
		  Return 1
	   Else
		  _addStatusText($StatusEdit, _getString('error', 'errUnzipFile', _ArrayCreate($sFile, @error)))
		  Return 0
	   EndIf
    EndIf
 EndFunc

 Func _getDrive($WinMain, $StatusEdit)
	Local $list, $ChooseButton, $CancelButton, $aDrive, $sDrive, $aDrives, $lvItem, $WinDriveSel
	Local $aDrives = DriveGetDrive('REMOVABLE')

    If UBound($aDrives) > 0 Then
	   $WinDriveSel = GUICreate(_getString('title', 'ttlDeterminedDrives'), 500, 280, -1, -1, $WS_MINIMIZEBOX + $WS_CAPTION + $WS_POPUP + $WS_SYSMENU + $WS_SIZEBOX, -1, $WinMain)

	   $list = GUICtrlCreateListView(_getString('header', 'hdrChooseDrive'), 5, 5, 490, 230)
	   $ChooseButton = GUICtrlCreateButton(_getString('button', 'btnConfirmChoice', 110, 240, 105, 35)
	   $CancelButton = GUICtrlCreateButton(_getString('button', 'btnCancel'), 225, 240, 105, 35)
	   for $sDrive in $aDrives
		  If DriveStatus($sDrive) == 'READY' Then
			 $lvItem = $sDrive & '|' & (DriveSpaceTotal($sDrive)/1024) & '|'
			 If FileExists($sDrive & '\garmin') Then
				$lvItem = $lvItem & _getString('header', 'hdrYes')
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
				_addStatusText($StatusEdit, _getString('error', 'errChoosenInvalidDrive'))
				GUIDelete($WinDriveSel)
				Return 0
			 Case $GUI_EVENT_CLOSE
				_addStatusText($StatusEdit, _getString('error', 'errChoosenInvalidDrive'))
				GUIDelete($WinDriveSel)
				Return 0
		  EndSwitch
	   WEnd
	Else
	   Do
		  If MsgBox(5, _getString('title', 'ttlWaitDrive'), _getString('useract', 'usrConnectNavi')) == 2 Then
			 _addStatusText($StatusEdit, _getString('error', 'errMissingDrive'))
			 Return 0
		  EndIf
	   Until UBound(DriveGetDrive('REMOVABLE')) > 0
	   Return _getDrive($WinMain, $StatusEdit)
    EndIf
 EndFunc

 Func _checkUpdate($StatusEdit, $StatusProgress, $StatusLabel, $EndButton)
	_addStatusText($StatusEdit, _getString('info', 'infCheckingUpdate'))
	Local $bData = InetRead('http://andrwe.org/mapupdater/version.txt', 1)
	Dim $aContent = StringRegExp(BinaryToString($bData), '^.*\s', 1)
	If $aContent[0] <> $VERSION Then
	   _addStatusText($StatusEdit, _getString('info', 'infFoundUpdate', _ArrayCreate($aContent[0], $VERSION, @CRLF)))
	   If MsgBox(36, _getString('title', 'ttlInstallUpdate'), _getString('useract', 'usrInstallUpdate')) == 7 Then
		  Return 0
	   EndIf
	   If _loadFile($StatusEdit, $StatusProgress, $StatusLabel, $EndButton, 'http://andrwe.org/mapupdater/', 'mapupdater-' & $aContent[0] & '.zip', 1) Then
		  _unzipFile($StatusEdit, $EndButton, @TempDir & '/mapupdater-' & $aContent[0] & '.zip', @ScriptDir, 0)
	   EndIf
	Else
	   _addStatusText($StatusEdit, _getString('info', 'infNoUpdate'))
	EndIf
 EndFunc

 Func main()
    Local $WinMain, $StartButton, $StatusEdit, $StatusProgress, $StatusLabel, $EndButton, $aImages, $sImage, $sTargetDir, $aDrive
	$WinMain = GUICreate(_getString('title', 'ttlMain'), 400, 440, -1, -1, $WS_MINIMIZEBOX + $WS_CAPTION + $WS_POPUP + $WS_SYSMENU + $WS_SIZEBOX)
    $StatusEdit = GUICtrlCreateEdit('', 5, 5, 390, 330, $ES_READONLY + $ES_MULTILINE + $ES_AUTOVSCROLL + $WS_VSCROLL)
	$StatusProgress = GUICtrlCreateProgress(5, 335, 390, 20, $PBS_SMOOTH)
	$StatusLabel = GUICtrlCreateLabel('', 5, 360, 390, 35)
	$StartButton = GUICtrlCreateButton(_getString('button', 'btnStartUpdate'), 110, 400, 105, 35)
	$EndButton = GUICtrlCreateButton(_getString('button', 'btnClose'), 225, 400, 105, 35)
	GUISetState()
	_addStatusText($StatusEdit, _getString('info', 'infDeterminedLang', _ArrayCreate(_getLanguage())))
	_checkUpdate($StatusEdit, $StatusProgress, $StatusLabel, $EndButton)
	While 1
		Switch GUIGetMsg()
			 Case $GUI_EVENT_CLOSE
				ExitLoop
			 Case $StartButton
				While Not _isInternetConnected()
				   GUISetState(@SW_DISABLE, $WinMain)
				   Switch MsgBox( 6, _getString('title', 'ttlConnectInet'), _getString('error', 'errConnectInet'))
					  Case 2
						 Exit
					  Case 11
						 ExitLoop
				   EndSwitch
			    WEnd
				GUISetState(@SW_ENABLE, $WinMain)
				GUISetState(@SW_SHOWNORMAL, $WinMain)
				GUICtrlSetState($StartButton, $GUI_DISABLE)
				$aImages = _getFiles($WinMain, $StatusEdit, $StatusProgress, $StatusLabel, $EndButton)
				If IsArray($aImages) And UBound($aImages) > 0 Then
				   $aDrive = _getDrive($WinMain, $StatusEdit)
				   If IsArray($aDrive) And UBound($aDrive) > 0 Then
					  If $aDrive[3] == _getString('header', 'hdrYes') Then
						 _addStatusText($StatusEdit, _getString('info', 'infFoundGarmin', _ArrayCreate(@CRLF)))
						 $sTargetDir = $aDrive[1] & '\garmin\'
					  Else
						 $sTargetDir = FileSelectFolder(_getString('useract', 'usrTargetPath'), $aDrive[1])
					  EndIf
					  For $sImage In $aImages
						 If Not $sImage Then ContinueLoop
;~ 						 $sFileSize = FileGetSize($sImage)
;~ 						 $sTarget
						 _addStatusText($StatusEdit, _getString('info', 'infMoveImage', _ArrayCreate($sImage, $sTargetDir)))
						 _moveFile($sImage, $sTargetDir)
;~ 						 If $sFileSize == 
;~ 							_addStatusText($StatusEdit, 'Fehler beim Verschieben der Karte auf das Gerät. Fehler-Code: ' & @error)
;~ 						 EndIf
					  Next
					  _addStatusText($StatusEdit, _getString('info', 'infFinishedUpdate'))
				   EndIf
			    EndIf
				GUICtrlSetState($StartButton, $GUI_ENABLE)
			 Case $EndButton
				ExitLoop
		EndSwitch
	WEnd
	
 EndFunc
 
 main()
