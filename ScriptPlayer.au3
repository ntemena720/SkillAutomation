#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <GuiConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <ColorConstants.au3>
#include <Array.au3>; arraydisplay
#include <File.au3>
#include <IE.au3>
#include <String.au3>
#include <Date.au3>
#include <ComboConstants.au3>
#include "GUIScrollbars_Ex.au3"
#include <GuiEdit.au3>
#include <GuiScrollBars.au3>
#RequireAdmin

Global $TrialEndDate = "2017/07/04" ; 00:00:00" ; YYYY/MM/DD Trial copy
Global $MainGUI, $CompNameGUI, $DashBoardGUI, $GUIFaultMenu ; Fault code Main Menu
Global $CNameOkButton = 999, $CNameCancelbutton = 999, $ProceedButton = 999, $MName
Global $ScriptLabel, $ObjectLabel, $FaultLabel,$EstTimeLabel
Global $RunList, $remotecomputer = ""
Global $AppFiles = @ScriptDir & "\App\*.ps1"
Global $TaskFiles = @ScriptDir & "\Task\*.ps1"
Global $SymptomFiles = @ScriptDir & "\Symp\*.ps1"
Global $faultfile =@ScriptDir & "\Util\fault.txt"
Global $modelfile =@ScriptDir & "\Util\model.txt"
Global $timelog = @ScriptDir & "\Util\time.txt"
Global $notefile = @ScriptDir & "\Util\note.txt"
Global $aFault[1], $aFaultLog[1] ; static number of faults discovered and action buttons
Global $ObjectsFound = 0, $FaultFound = 0
Global $NoteButton[1], $ModelFlButton[1], $ModelFdrButton[1], $HotFixButton[1], $ServiceButtonR[1],$ServiceButtonS[1], $RelatedLinkButton[1] ; static number of faults discovered and action buttons
Global $AddScriptButton = 999
Global $MaxFileLine = 0 ; number of fault txt lines -> max number of fault array
Global $maxstring = 0 ; maximum number of text in an array; this will dictate the width of results menu
Global $aFilemodel[1]; array of compname model
Global $ObjectMapMenu, $MapButton = 999, $ModelChoice, $DefaultListValue, $ObjectSelected, $ObjectLocation, $ObjectHeader, $ManualLink, $BuyLink
Global $aScriptTime[1][2], $ScriptCount = 0, $htimer, $hTimeFileOpen ; script time
Global $sFont = "Verdana"
Global $Color1 = "0xa60404" ; red
Global $Color2 = "0x222323" ;dark grey for >>
Global $Color3 = "0xcfd7d7" ; Light grey for the GUI
Global $Color4 = "0xe0e0e0" ; player background color
Global $Color5 = "0x990000" ; dark RED "0xebeeee" ; dark green gray List
Global $Color6 = "0xe0e0e0" ; Profile list background
Global $Color7 = "0xFF8000" ; List profile text


#Region Main
HotKeySet("+!n", "CreateNote") ; hotkey for note menu
$MainGUI = GUICreate("Script Player v0.53", 600, 320, -1, -1, -1, $WS_EX_TOOLWINDOW) ; 540 ;-1, $WS_EX_TOOLWINDOW
GUICtrlSetFont(-1, 9, 400, 0, $sFont)
GUISetBkColor($Color4)
CreateWindowMenu()

#Region App
ProcessSampleScript() ; process sample scripts to point to the right path

GUICtrlCreateLabel("Application Profile", 90, 8, 104, 17)
GUICtrlSetFont(-1, 9, 700, 0, "MS Sans Serif")
$AppList = GUICtrlCreateList("", 16, 32, 233, 58, BitOR($GUI_SS_DEFAULT_LIST, $LBS_EXTENDEDSEL))
;GUICtrlSetBkColor($AppList, $Color6)
$AddAppButton = GUICtrlCreateButton(">>", 264, 49, 33, 25, $BS_DEFPUSHBUTTON, $WS_EX_WINDOWEDGE)
GUICtrlSetFont(-1, 11, 500, 0, $sFont)
;GUICtrlSetColor(-1, $Color2)
;GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
_GUICtrlListBox_Dir($AppList, $AppFiles)

#EndRegion App

#Region Task
GUICtrlCreateLabel("Task Profile", 105, 104, 95, 17)
GUICtrlSetFont(-1, 9, 700, 0, "MS Sans Serif")
$TaskList = GUICtrlCreateList("", 16, 128, 233, 58, BitOR($GUI_SS_DEFAULT_LIST, $LBS_EXTENDEDSEL))
;GUICtrlSetBkColor($TaskList, $Color3)
$AddTaskButton = GUICtrlCreateButton(">>", 264, 143, 33, 25, -1, $WS_EX_WINDOWEDGE)
GUICtrlSetFont(-1, 11, 500, 0, $sFont)
;GUICtrlSetColor(-1, $Color2)
;GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
_GUICtrlListBox_Dir($TaskList, $TaskFiles)

#EndRegion Task

#Region Symptom
GUICtrlCreateLabel("Symptom Profile", 90, 200, 95, 17)
GUICtrlSetFont(-1, 9, 700, 0, "MS Sans Serif")
$SymptomList = GUICtrlCreateList("", 16, 224, 233, 58, BitOR($GUI_SS_DEFAULT_LIST, $LBS_EXTENDEDSEL))
;GUICtrlSetBkColor($SymptomList, $Color3)
$AddSympButton = GUICtrlCreateButton(">>", 264, 242, 33, 25)
GUICtrlSetFont(-1, 11, 500, 0, $sFont)
;GUICtrlSetColor(-1, $Color2)
;GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
_GUICtrlListBox_Dir($SymptomList, $SymptomFiles)

#EndRegion Symptom

$RunList = GUICtrlCreateList("", 312, 32, 185, 253, $LBS_NOSEL)
;GUICtrlSetBkColor($RunList, $Color3)
$RunButton = GUICtrlCreateButton("Run", 513, 72, 75, 35)
$ClearButton = GUICtrlCreateButton("Clear List", 513, 134, 75, 35)
$ExitButton = GUICtrlCreateButton("Exit", 513, 196, 75, 35)

SetDefaultSel($AppList)
SetDefaultSel($TaskList)
SetDefaultSel($SymptomList)

GUISetState(@SW_SHOW, $MainGUI)

#EndRegion Main
While 1
	$nMsg = GUIGetMsg(1)
	Switch $nMsg[1]
		Case $MainGUI
			Switch $nMsg[0]
				Case $GUI_EVENT_CLOSE, $ExitButton
					;DeleteFaultfile()
					Exit
				Case $AddAppButton
					GetSelectedScript($AppList)
				Case $AddTaskButton
					GetSelectedScript($TaskList)
				Case $AddSympButton
					GetSelectedScript($SymptomList)
				Case $RunButton
					MachineNameMenu()
				Case $ClearButton
					ClearAll()
				Case $ManualLink
					OpenBrowser("https://www.i-ttm.com/script-player-manual.html") ;
				Case $BuyLink
					OpenBrowser("https://www.i-ttm.com/store/c1/Featured_Products.html") ;
			EndSwitch
		Case $CompNameGUI
			Switch $nMsg[0]
				Case $GUI_EVENT_CLOSE, $CNameCancelbutton
					GUIDelete($CompNameGUI)
					GUISetState(@SW_RESTORE, $MainGUI)
					GUISetState(@SW_ENABLE, $MainGUI)
				Case $CNameOkButton
					GUISetState(@SW_MINIMIZE, $MainGUI)
					RunScripts()
			EndSwitch
		Case $DashBoardGUI
			Switch $nMsg[0]
				Case $GUI_EVENT_CLOSE
					GUIDelete($DashBoardGUI)
					ResetObject_Fault()
					GUISetState(@SW_RESTORE, $MainGUI)
					GUISetState(@SW_ENABLE, $MainGUI)
				Case $ProceedButton
					GUIDelete($DashBoardGUI)
					FaultMenu()
			EndSwitch
	EndSwitch
WEnd

Func GetSelectedScript($listbox) ; get text from the selected or highlighted entry on the listbox

	Local $SelScripts, $iSel

	$SelScripts = _GUICtrlListBox_GetSelItemsText($listbox)
	If UBound($SelScripts) = 2 Then ;IsArray($SelScripts) Then
		_GUICtrlListBox_AddString($RunList, $SelScripts[1]) ; to add value to the running list
		$iSel = _GUICtrlListBox_GetCurSel($listbox)
		_GUICtrlListBox_DeleteString($listbox, $iSel) ; to delete list from origin
	ElseIf UBound($SelScripts) = 1 Then
		MsgBox(16, "Script Player v0.53", "Please select a Profile") ; error check
	Else
		MsgBox(16, "Script Player v0.53", "Please select a Profile") ; all trap error check
	EndIf

EndFunc   ;==>GetSelectedScript

Func SetDefaultSel($Lbox) ; highlight first entry from all listbox
	_GUICtrlListBox_SetSel($Lbox, 0)
EndFunc   ;==>SetDefaultSel

Func ClearAll() ;  clear contents from the irunning list box and repopulate all listbox
	_GUICtrlListBox_ResetContent($RunList)
	_GUICtrlListBox_ResetContent($AppList)
	_GUICtrlListBox_Dir($AppList, $AppFiles)
	_GUICtrlListBox_ResetContent($TaskList)
	_GUICtrlListBox_Dir($TaskList, $TaskFiles)
	_GUICtrlListBox_ResetContent($SymptomList)
	_GUICtrlListBox_Dir($SymptomList, $SymptomFiles)
EndFunc   ;==>ClearAll

Func RunScripts()
	Local $FileScript = ""
	Local $computername = ""
	Local $ScriptLocation = ""
	$computername = MachineName()
	If CheckMachineOnline($computername) Then
		$ScriptCount = _GUICtrlListBox_GetListBoxInfo($RunList)
		;MsgBox(0, "Max row", $ScriptCount)
		ReDim $aScriptTime[$ScriptCount][2]
		;GUICtrlSetData($ScriptLabel, $ScriptCount)
		;MsgBox(0, "1111", "inside runscript")
		For $i = 0 To $ScriptCount - 1 Step 1
			$FileScript = _GUICtrlListBox_GetText($RunList, $i) ; & @CRLF
			$FileScript = StringStripWS($FileScript, 3) ; strip leading and trailing white space
			$ScriptLocation = LocateScriptDir($FileScript)
			MeasureTime("Start")
			RunPS($ScriptLocation) ; pass scriptfilepath
			$aScriptTime[$i][0] = $filescript
			$aScriptTime[$i][1] = MeasureTime("Stop")

		Next
	EndIf
	;_ArrayDisplay($aScriptTime, "Script Time")
	WriteTimeLog($aScriptTime)
	CountFaults()
EndFunc   ;==>RunScripts

Func LocateScriptDir($ScriptName)
	Local $AppDir = @ScriptDir & "\App\" & $ScriptName
	Local $TaskDir = @ScriptDir & "\Task\" & $ScriptName
	Local $SympDir = @ScriptDir & "\Symp\" & $ScriptName

	If FileExists($AppDir) Then
		CountObjects($AppDir)
		Return $AppDir
	ElseIf FileExists($TaskDir) Then
		CountObjects($TaskDir)
		Return $TaskDir
	ElseIf FileExists($SympDir) Then
		CountObjects($SympDir)
		Return $SympDir
	Else
		Return False
	EndIf
EndFunc   ;==>LocateScriptDir

Func CreateWindowMenu() ; menuv
	Local $ManualInfoMenu = GUICtrlCreateMenu("Help")
	Local $idInfoMenu = GUICtrlCreateMenu("About")
	Local $BuyMenu = GUICtrlCreateMenu("Purchase")
	GUICtrlCreateMenuItem("Report bugs to bugreport@i-ttm.com", $idInfoMenu)
	GUICtrlCreateMenuItem("Trial copy expires : " & _DateTimeFormat($TrialEndDate, 2), $idInfoMenu)
	$ManualLink = GUICtrlCreateMenuItem("Open Script Player Manual", $ManualInfoMenu)
	$BuyLink = GUICtrlCreateMenuItem("Get Script Player Software", $BuyMenu)
EndFunc   ;==>CreateWindowMenu

Func RunPS($filenamescript)
	Local $result, $online
	; Local $filepath = ' -File  "' & @ScriptDir & '\' & $filenamescript & '" ' & $remotecomputer
	Local $filepath = ' -File  "' & $filenamescript & '" ' & $remotecomputer
	;MsgBox(0,"", $filepath)
	;MsgBox(4096,""," powershell.exe -executionpolicy unrestricted -NoExit -NoProfile " &  $filepath )
	RunWait('powershell.exe -executionpolicy unrestricted -NoProfile ' & $filepath, @SystemDir, @SW_HIDE, 7) ; sw_maximize and -NoExit - Run
	;Run('powershell.exe -executionpolicy unrestricted -NoExit -NoProfile  -Command  " & "' &  $filepath     , @SystemDir, @SW_MAXIMIZE, 7)
EndFunc   ;==>RunPS

Func CheckMachineOnline($compname)
	Local $online, $status, $TestOnlineScriptBase
	$TestOnlineScriptBase = "$status = Test-Connection -computername " & $compname & " -quiet" & @CR & _
			"if ( $status  -eq $true) {write-host '1'} Else { write-host '0'} "
	;Run("powershell.exe -executionpolicy unrestricted -NoExit -NoProfile  "  & $TestOnlineScriptBase, @SystemDir, @SW_MAXIMIZE, 3)
	$status = Run("powershell.exe -executionpolicy unrestricted -NoProfile  " & $TestOnlineScriptBase, @SystemDir, @SW_HIDE, 3)
	MsgBox(0, "Script Player v0.53", "Checking if machine is online", 1)
	StdinWrite($status, @CRLF)
	StdinWrite($status)
	While 1
		$online &= StdoutRead($status)
		If @error Then ExitLoop
	WEnd
	If $online = 1 Then
		GUIDelete($CompNameGUI)
		DeleteFaultfile() ; +++++++++++++++++++++ write formula so it wont delete file from test scenario
		DashboardMenu()
		ReadTimeLog()
		$remotecomputer = $compname
		Return True
	Else
		MsgBox(16, "Script Player v0.53", "Machine not online")
		MachineNameMenu()
		Return False
	EndIf
EndFunc   ;==>CheckMachineOnline

Func MachineNameMenu()
	$sCompData = AutoClip()

	If ListNotEmpty() Then
		$CompNameGUI = GUICreate("Script Player v0.53", 230, 166, -1, -1, -1, $WS_EX_TOOLWINDOW, $MainGUI)
		GUISetBkColor($Color4)
		GUICtrlCreateLabel("Enter Machine Name", 24, 16, -1, -1)
		$MName = GUICtrlCreateInput($sCompData, 16, 40, 193, 21)
		GUICtrlSetState($MName, $GUI_FOCUS)
		;$HidePshell = GUICtrlCreateCheckbox("Unhide power shell window", 48, 112, 153, 17)
		$CNameOkButton = GUICtrlCreateButton("OK", 24, 72, 75, 25)
		GUICtrlSetState($CNameOkButton, $GUI_DEFBUTTON)
		$CNameCancelbutton = GUICtrlCreateButton("Cancel", 128, 72, 75, 25)
		GUISetState(@SW_SHOW, $CompNameGUI)
	Else
		MsgBox(16, "Script Player v0.53", " Please select a script to run.")
		Return
	EndIf

EndFunc   ;==>MachineNameMenu

Func MachineName()
	Local $compname = "" ; initialize computername
	Local $tempname = ""

	GUISetState(@SW_DISABLE, $MainGUI)
	;msg("assigning mname to tempname: " & $tempname)
	$tempname = GUICtrlRead($MName)
	If ($tempname = "") Or (@error > 0) Then
		Return False
	EndIf
	$compname = StringStripWS($tempname, 8)
	;msg( "compname returning value: " & $compname)
	Return $compname
	; +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
EndFunc   ;==>MachineName

Func DashboardMenu()
	$DashBoardGUI = GUICreate("Script Player v0.53", 350, 320, -1, -1, -1, $WS_EX_TOOLWINDOW, $MainGUI)
	GUISetBkColor($Color4)
	GUICtrlCreateLabel("# of Running Script(s)", 14, 208, 88, 50, $SS_CENTER)
	GUICtrlSetFont(-1, 9, 700, 0, $sFont)
	GUICtrlCreateLabel("# of Objects to Isolate", 248, 208, 88, 50, $SS_CENTER)
	GUICtrlSetFont(-1, 9, 700, 0, $sFont)
	GUICtrlCreateLabel("# of Suspect Faults Found", 135, 208, 88, 50, $SS_CENTER)
	GUICtrlSetFont(-1, 9, 700, 0, $sFont)
	GUICtrlCreateLabel("Est. Script Time Duration mm:ss", 65, 88, 220, 50)
    GUICtrlSetFont(-1, 9, 700, 0, $sFont)
	$ScriptLabel = GUICtrlCreateLabel("0", 28, 128, 66, 50, BitOR($SS_CENTER, $SS_SUNKEN))
	GUICtrlSetFont(-1, 27, 900, 0, $sFont)
	GUICtrlSetColor(-1, $COLOR_WHITE)
	GUICtrlSetBkColor($ScriptLabel, $COLOR_BLUE)
	$ObjectLabel = GUICtrlCreateLabel("0", 260, 128, 66, 50, BitOR($SS_CENTER, $SS_SUNKEN))
	GUICtrlSetFont(-1, 27, 900, 0, $sFont)
	GUICtrlSetColor(-1, $COLOR_WHITE)
	GUICtrlSetBkColor($ObjectLabel, $COLOR_BLUE)
	$FaultLabel = GUICtrlCreateLabel("0", 148, 128, 66, 50, BitOR($SS_CENTER, $SS_SUNKEN))
	GUICtrlSetFont(-1, 27, 900, 0, $sFont)
	GUICtrlSetColor(-1, $COLOR_WHITE)
	GUICtrlSetBkColor($FaultLabel, $COLOR_RED)
	$EstTimeLabel = GUICtrlCreateLabel("99:99", 75, 24, 205, 50, BitOR($SS_CENTER,$SS_SUNKEN))
	GUICtrlSetFont(-1, 27, 700, 0, $sFont)
	GUICtrlSetColor(-1, $COLOR_WHITE)
	GUICtrlSetBkColor($EstTimeLabel, $COLOR_GREEN)
	$ProceedButton = GUICtrlCreateButton("NEXT",130, 265, 99, 41)
	GUICtrlSetState($ProceedButton, $GUI_HIDE)
	GUISetState(@SW_SHOW, $DashBoardGUI)
EndFunc   ;==>DashboardMenu

Func ResetObject_Fault()
	$ObjectsFound = 0
	$FaultFound = 0
EndFunc   ;==>ResetObject_Fault

Func DeleteFaultfile()
	FileDelete($faultfile)
EndFunc   ;==>DeleteFaultfile

Func ListNotEmpty()

	Local $MaxRow
	$MaxRow = _GUICtrlListBox_GetListBoxInfo($RunList)
	If $MaxRow <> 0 Then
		Return True
	Else
		Return False
	EndIf
EndFunc   ;==>ListNotEmpty

Func CountFaults() ; open fault file and parse the log
	Local $sFileRead, $sFileRead
	If Not FileExists($faultfile) Then ; check for fault.txt
		_FileCreate($faultfile)
		;Return
	EndIf

	If Not FileExists($modelfile) Then ; check model .txt
		_FileCreate($modelfile)
	EndIf

	$MaxFileLine = _FileCountLines($faultfile)

	$hFaultFileOpen = FileOpen($faultfile, 0) ; Read fault.txt on Read Mode
	If $hFaultFileOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "Script Player v0.53", "Error processing Fault File.")
		Return
	EndIf

	$hModelFileOpen = FileOpen($modelfile, 1) ; Open model.txt in Append Mode
	If $hModelFileOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "Script Player v0.53", "Error processing Model File.")
		Return
	EndIf

	ReDim $aFault[$MaxFileLine]
	ReDim $aFaultLog[$MaxFileLine]
	;msg("max  file line is: " & $MaxFileLine)
	For $i = 1 To $MaxFileLine
		$sFileRead = FileReadLine($hFaultFileOpen, $i)
		If StringInStr($sFileRead, "#N#", 2) Then
			$FaultFound += 1
			$aFaultLog[$FaultFound - 1] = $sFileRead ; text string to a cell array
			$faultvariable = _StringExplode($sFileRead, "#") ; convert to array
			FormatFault($faultvariable, ($FaultFound - 1)) ; assign fault object to aFault array +++++++++++++++++++++++++
			UpdateFaultCount($FaultFound) ; update fault count on dashboard +++++++++++++++++++++++
		ElseIf StringInStr($sFileRead, "#Y#", 2) Then
			FileWriteLine($hModelFileOpen, $sFileRead) ; assign working objects to model.txt
		EndIf
	Next
	;msg("outside loop")
	FileClose($hFaultFileOpen)
	FileClose($hModelFileOpen)
	;msg("number of faults: " & $FaultFound)
	If $FaultFound = 0 Then
		MsgBox(0, "Script Player v0.53", "No faults found")
		Return
	Else

		GUICtrlSetData($FaultLabel, $FaultFound)
		If NotExpired() Then
			;msg("not expired 11")
			_ArrayRemoveBlanks($aFaultLog) ; array of fault  strings #N#
			RemoveBlanksandGetMax($aFault) ; and also get Max string ; array of fault string status "does not contain"
			GUICtrlSetState($ProceedButton, $GUI_SHOW)
		EndIf

	EndIf

EndFunc   ;==>CountFaults

Func FormatFault($FaultArray, $count) ;; assign status text for every negative log result
	;_ArrayDisplay($FaultArray, "inside formatfault")
	Switch $FaultArray[0]
		Case "File"
			;MsgBox(0,"File Case",$var)
			$aFault[$count] = StringUpper($FaultArray[2]) & " file not found on " & $FaultArray[3] ; - take out ""
		Case "FileSize"
			;MsgBox(0,"File Size Case",$var)
			$aFault[$count] = StringUpper($FaultArray[2]) & " has the wrong file size of : " & $FaultArray[4] & " Expected size: " & $FaultArray[5]
		Case "FileWord"
			;MsgBox(0,"File Word Case",$var)
			$aFault[$count] = StringUpper($FaultArray[2]) & " file does not contain the word: " & StringUpper($FaultArray[4])
		Case "FileCreate"
			;MsgBox(0,"File Create Case",$var)
			$aFault[$count] = StringUpper($FaultArray[2]) & " file has incorrect create date of : " & $FaultArray[4] & " Expected file create date : " & $FaultArray[5]
		Case "FileModified"
			;MsgBox(0,"File Modified Case",$var)
			$aFault[$count] = StringUpper($FaultArray[2]) & " file has incorrect modified date of : " & $FaultArray[4] & " Expected file modified date : " & $FaultArray[5]
		Case "Folder"
			;MsgBox(0,"Folder Case",$var)
			$aFault[$count] = StringUpper($FaultArray[2]) & " folder not found. Expected folder path : " & $FaultArray[3]
		Case "FolderCreate"
			;MsgBox(0,"Folder Create Case",$var)
			$aFault[$count] = StringUpper($FaultArray[2]) & " folder has incorrect create date of : " & $FaultArray[4] & " Expected folder create date : " & $FaultArray[5]
		Case "FolderModified"
			;MsgBox(0,"Folder Modified Case",$var)
			$aFault[$count] = StringUpper($FaultArray[2]) & " folder has incorrect modified date of : " & $FaultArray[4] & " Expected folder modified date : " & $FaultArray[5]
		Case "FolderNumber"
			;MsgBox(0,"Folder Number Case",$var)
			$aFault[$count] = StringUpper($FaultArray[2]) & " folder has incorrect number of files : " & $FaultArray[4] & " Expected number of files :  " & $FaultArray[5]
		Case "RegSubkey"
			;MsgBox(0,"RegSubkey Case",$var)
			$aFault[$count] = "Registry HIVE and KEY not accessible. Expected value : " & $FaultArray[3]
		Case "RegSubkeyValue"
			;MsgBox(0,"RegSubkeyValue Case ", $var)
			$aFault[$count] = "Registry SUBKEY VALUE not found. Expected value : " & $FaultArray[3]
		Case "Regdata"
			;MsgBox(0,"RegData Case",$var)
			$aFault[$count] = StringUpper($FaultArray[2]) & " data is incorrect. Expected value : " & $FaultArray[3] ; longest length
		Case "Process"
			;MsgBox(0,"Process Case",$var)
			$aFault[$count] = StringUpper($FaultArray[2]) & " process not found" ;;; showing $computername
		Case "Service"
			;MsgBox(0,"Service Case",$var)
			$aFault[$count] = StringUpper($FaultArray[2]) & " service " & $FaultArray[3] ;;;;
		Case "Online"
			;MsgBox(0,"Online Case",$var)
			$aFault[$count] = $FaultArray[2] & " is not online"
		Case "HotFix"
			;MsgBox(0,"HotFix Case",$var)
			$aFault[$count] = StringUpper($FaultArray[2]) & " patch not found"
	EndSwitch
	; _ArrayDisplay($aFault, " text describing the faults")


EndFunc   ;==>FormatFault

Func FaultMenu() ; Menu for display all the negative results and action buttons
	Local $labelwidth
	Local $topheight = 20
	Local $ModelIncluded = 0
	Local $EnableVScroll = 0
	ReDim $NoteButton[$FaultFound], $RelatedLinkButton[$FaultFound] ; ; resize action buttons array to the right number
	ReDim $ModelFlButton[$FaultFound], $ModelFdrButton[$FaultFound], $HotFixButton[$FaultFound], $ServiceButtonR[$FaultFound],$ServiceButtonS[$FaultFound] ; resize action buttons array to the right number

	$MenuWidth = 75 + ($maxstring * 6)
	;msg( "max string is: "  & $maxstring)
	;msg(" Menu width is: " & $MenuWidth)
	If $MenuWidth < 325 Then ; set minimum window width
		$MenuWidth = 325
	EndIf

	$MenuHeight =  45 + ($FaultFound * 60)
    If $MenuHeight > 400 Then  ; set maximum window height
		$ScrollHeight = $MenuHeight
		$MenuHeight = 400
		$EnableVScroll = 1
	Endif

	$GUIFaultMenu = GUICreate("Script Player v0.53", $MenuWidth, $MenuHeight, -1, -1, -1, $WS_EX_TOOLWINDOW, $MainGUI) ;, $WS_EX_TOOLWINDOW); ++++++ min and max view menu becomes unformatted
	;$GUIFaultMenu = GUICreate("Script Player v0.53", 250, 400, -1, -1 ) ;,  -1, $WS_EX_TOOLWINDOW ) ;,-1, $MainGUI) ;, $WS_EX_TOOLWINDOW); ++++++ min and max view menu becomes unformatted

	GUISetBkColor($Color4)
	For $i = 0 To UBound($aFault) - 1

		$RelatedLinkButton[$i] = -1 ; assign value to button so button won't misfire
		$labelwidth = 35 + (StringLen($aFault[$i]) * 6)
		GUICtrlCreateLabel($aFault[$i], 20, $topheight, $labelwidth, 18) ;, $SS_SUNKEN)
		GUICtrlSetColor(-1, $Color5)
		$topheight += 23

		If StringInStr($aFaultLog[$i], "Online", 2, 1, 1, 6) Then
			$NoteButton[$i] = GUICtrlCreateButton("Contact Info", 20, $topheight, 80, 30) ; diffrent button label for offline servers
			GUICtrlSetBkColor(-1, $Color3)
		Else
			$NoteButton[$i] = GUICtrlCreateButton("Notes", 20, $topheight, 80, 30) ; default button lable of notes
			GUICtrlSetBkColor(-1, $Color3)
		EndIf

		If StringInStr($aFaultLog[$i], "File", 2, 1, 1, 4) Then ; File Error
			$modelfilebuttonexist = 0
			$aItem = _StringExplode($aFaultLog[$i], "#")
			If ObjectHasModel($aItem[0] & "#Y#" & $aItem[2]) Then
				$ModelFlButton[$i] = GUICtrlCreateButton("Model", 120, $topheight, 80, 30)
				GUICtrlSetBkColor(-1, $Color3)
				$ModelIncluded = 1
				$modelfilebuttonexist = 1
			Else
				$ModelFlButton[$i] = -1
			EndIf
			If NoteshasLink($aFaultLog[$i]) Then
				If $modelfilebuttonexist Then
					$fileleftwidth = 220
				Else
					$fileleftwidth = 120
				EndIf
				$RelatedLinkButton[$i] = GUICtrlCreateButton("Open Link", $fileleftwidth, $topheight, 80, 30)
				GUICtrlSetBkColor(-1, $Color3)
			EndIf
		Else
			$ModelFlButton[$i] = -1
		EndIf

		If StringInStr($aFaultLog[$i], "Folder", 2, 1, 1, 6) Then ; Folder Error

			$modelfolderbuttonexist = 0
			$aItem = _StringExplode($aFaultLog[$i], "#")

			If ObjectHasModel($aItem[0] & "#Y#" & $aItem[2]) Then
				$ModelFdrButton[$i] = GUICtrlCreateButton("Model", 120, $topheight, 80, 30) ; $topheight, 80, 30)
				GUICtrlSetBkColor(-1, $Color3)
				$ModelIncluded = 1
				$modelfolderbuttonexist = 1
			Else
				$ModelFdrButton[$i] = -1
			EndIf

			If NoteshasLink($aFaultLog[$i]) Then
				If $modelfolderbuttonexist Then
					$folderleftwidth = 220
				Else
					$folderleftwidth = 120
				EndIf
				$RelatedLinkButton[$i] = GUICtrlCreateButton("Open Link", $folderleftwidth, $topheight, 80, 30)
				GUICtrlSetBkColor(-1, $Color3)
			EndIf
		Else
			$ModelFdrButton[$i] = -1
		EndIf


		If StringInStr($aFaultLog[$i], "HotFix", 2, 1, 1, 6) Then ; Hotfix KB error
			$HotFixButton[$i] = GUICtrlCreateButton("Search Web", 120, $topheight, 80, 30)
			GUICtrlSetBkColor(-1, $Color3)

			If NoteshasLink($aFaultLog[$i]) Then
				$leftwidth = 220
				$RelatedLinkButton[$i] = GUICtrlCreateButton("Open Link", $leftwidth, $topheight, 80, 30)
				GUICtrlSetBkColor(-1, $Color3)
			EndIf
		Else
			$HotFixButton[$i] = -1
		EndIf

		If StringInStr($aFaultLog[$i], "RegSubKey#", 2, 1, 1, 10) Then ;
			If NoteshasLink($aFaultLog[$i]) Then
				$RelatedLinkButton[$i] = GUICtrlCreateButton("Open Link", 120, $topheight, 80, 30)
				GUICtrlSetBkColor(-1, $Color3)
			EndIf
		EndIf

		If StringInStr($aFaultLog[$i], "RegSubKeyValue", 2, 1, 1, 14) Then ;
			If NoteshasLink($aFaultLog[$i]) Then
				$RelatedLinkButton[$i] = GUICtrlCreateButton("Open Link", 120, $topheight, 80, 30)
				GUICtrlSetBkColor(-1, $Color3)
			EndIf
		EndIf

		If StringInStr($aFaultLog[$i], "RegData", 2, 1, 1, 7) Then ;
			If NoteshasLink($aFaultLog[$i]) Then
				$RelatedLinkButton[$i] = GUICtrlCreateButton("Open Link", 120, $topheight, 80, 30)
				GUICtrlSetBkColor(-1, $Color3)
			EndIf
		EndIf

		If StringInStr($aFaultLog[$i], "Process", 2, 1, 1, 7) Then ; Hotfix KB error
			If NoteshasLink($aFaultLog[$i]) Then
				$RelatedLinkButton[$i] = GUICtrlCreateButton("Open Link", 120, $topheight, 80, 30)
				GUICtrlSetBkColor(-1, $Color3)
			EndIf
		EndIf

		If StringInStr($aFaultLog[$i], "Service", 2, 1, 1, 7) Then ; Service error
			$startbuttonexist = 0
			If StringInStr($aFaultLog[$i], "#stopped#", 2, -1) Then ; Stopped error
				$ServiceButtonS[$i] = GUICtrlCreateButton(" Start Service", 120, $topheight, 80, 30)
				GUICtrlSetBkColor(-1, $Color3)
				$startbuttonexist = 1
			Else
				$ServiceButtonS[$i] = -1
			EndIf

			If StringInStr($aFaultLog[$i], "#running#", 2, -1) Then ; Stopped error
				$ServiceButtonR[$i] = GUICtrlCreateButton(" Stop Service", 120, $topheight, 80, 30)
				GUICtrlSetBkColor(-1, $Color3)
				$startbuttonexist = 1
			Else
				$ServiceButtonR[$i] = -1
			EndIf

			If NoteshasLink($aFaultLog[$i]) Then
				If $startbuttonexist Then
					$leftmargin = 220
				Else
					$leftmargin = 120
				EndIf
				$RelatedLinkButton[$i] = GUICtrlCreateButton("Open Link", $leftmargin, $topheight, 80, 30)
				GUICtrlSetBkColor(-1, $Color3)
			EndIf

		Else
			$ServiceButtonR[$i] = -1
			$ServiceButtonS[$i] = -1

		EndIf

		$topheight += 34
	Next



	If $ModelIncluded Then
		GUICtrlCreateLabel("Model = Show known working machine to compare ", 20, $topheight + 15, 320, 18)
		GUICtrlSetFont(-1, 7, 300)
	EndIf

	; _ArrayDisplay($RelatedLinkButton,"Link button values")
	 ;_ArrayDisplay($HotFixButton,"hot fix button values") ;$ModelFlButton[$i]
	; _ArrayDisplay($ModelFlButton,"model file button values")
	; _ArrayDisplay($ModelFdrButton, " Model Folder button values")
	;_ArrayDisplay($NoteButton, "Notes button values")

	;GUISetState(@SW_MINIMIZE,$MainGUI)
	GUISetState(@SW_DISABLE, $MainGUI)
	GUISetState(@SW_SHOW, $GUIFaultMenu)

	If $EnableVScroll Then
		_GUIScrollbars_Generate($GUIFaultMenu, 0, $ScrollHeight ) ;;; scroll bar control
	Endif
	CleanModel()
	;***************************************************************************************************************************
	While 1
		$nMsg = GUIGetMsg(0)
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				GUISetState(@SW_ENABLE, $MainGUI)
				ResetObject_Fault()
				GUISetState(@SW_RESTORE, $MainGUI)
				GUIDelete($GUIFaultMenu)
				ExitLoop
			Case $nMsg > 0
				For $ii = 0 To UBound($NoteButton) - 1
					If $NoteButton[$ii] = $nMsg Then
						NotesMenu($ii)
					EndIf
				Next
				For $iii = 0 To UBound($ModelFlButton) - 1
					If $ModelFlButton[$iii] = $nMsg Then
						ModelFileMenu($iii)
					EndIf
				Next
				For $x = 0 To UBound($ModelFdrButton) - 1
					If $ModelFdrButton[$x] = $nMsg Then
						ModelFolderMenu($x)
					EndIf
				Next
				For $xxx = 0 To UBound($HotFixButton) - 1
					If $HotFixButton[$xxx] = $nMsg Then
						SearchWebMenu($xxx)
					EndIf
				Next
				For $y = 0 To UBound($ServiceButtonR) - 1
					If $ServiceButtonR[$y] = $nMsg Then
						ServiceStopMenu($y)
					EndIf
				Next

				For $yx = 0 To UBound($ServiceButtonS) - 1
					If $ServiceButtonS[$yx] = $nMsg Then
						ServiceStartMenu($yx)
					EndIf
				Next
				For $yy = 0 To UBound($RelatedLinkButton) - 1
					If $RelatedLinkButton[$yy] = $nMsg Then
						OpenLink($yy)
					EndIf
				Next
		EndSwitch
	WEnd

	;*****************************************************************************************************************************

EndFunc   ;==>FaultMenu

Func _ArrayRemoveBlanks(ByRef $arr) ; delete empty arrays
	_ArrayDisplay(" Display array with blanks", $arr)
	$idx = 0
	For $i = 0 To UBound($arr) - 1
		If $arr[$i] <> "" Then
			$arr[$idx] = $arr[$i]
			$idx += 1
		EndIf
	Next
	ReDim $arr[$idx]
EndFunc   ;==>_ArrayRemoveBlanks

Func RemoveBlanksandGetMax(ByRef $arr) ; delete empty array and get maximum number of status strings
	$idx = 0
	For $i = 0 To UBound($arr) - 1
		If $arr[$i] <> "" Then
			GetMaxLength($arr[$i]) ; Get max string
			$arr[$idx] = $arr[$i]
			$idx += 1
		EndIf
	Next
	ReDim $arr[$idx]
EndFunc   ;==>RemoveBlanksandGetMax

Func GetMaxLength($textstring) ; get the maximum number of strings on text status
	If $maxstring < StringLen($textstring) Then
		$maxstring = StringLen($textstring)
	EndIf
	;msg("Got Max Lenght")
EndFunc   ;==>GetMaxLength

Func CountObjects($scriptfile)
	Local $sFileRead, $MaxFileLine, $sFileRead

	If Not FileExists($scriptfile) Then
		MsgBox(16, "Script Player v0.53", "No Script Detected")
		Return
	EndIf

	$MaxFileLine = _FileCountLines(@ScriptFullPath)
	$hFileOpen = FileOpen($scriptfile, 0) ; Read Mode
	If $hFileOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "Script Player v0.53", "Error reading Script.")
		Return
	EndIf

	For $i = 1 To $MaxFileLine
		$sFileRead = FileReadLine($hFileOpen, $i)
		If StringInStr($sFileRead, "***************", 2) Then
			$ObjectsFound += 1
		EndIf
	Next
	FileClose($hFileOpen)
	; Global var = $MaxFileLine - 1
	GUICtrlSetData($ObjectLabel, $ObjectsFound)
EndFunc   ;==>CountObjects

Func ActionMenu($item)
	MsgBox(0, "ITEM for Action", $aFaultLog[$item])
EndFunc   ;==>ActionMenu

Func NotesMenu($item) ; assign and show notes for each objects
	;_ArrayDisplay($aFaultLog," fault log inside notes menu")
	$aNotes = _StringExplode($aFaultLog[$item], "#", 7) ; ++++++++++++   array input to search for links
	;MsgBox(0,"array_test","line before array")
	;_ArrayDisplay($aNotes," NOTES")
	;MsgBox(0,"array_test","line after arrayd")

	Switch $aNotes[0]
		Case "Hotfix"
			NoteGui($aNotes[5], 0)
		Case "Online"
			NoteGui($aNotes[4], 1) ; flag if contact button is pressed and use contact menu header
		Case "Service"
			NoteGui($aNotes[6], 0)
		Case "Process"
			NoteGui($aNotes[5], 0)
		Case "Regdata"
			NoteGui($aNotes[5], 0)
		Case "RegSubkeyValue"
			NoteGui($aNotes[5], 0)
		Case "RegSubkey"
			NoteGui($aNotes[5], 0)
		Case "FolderNumber"
			NoteGui($aNotes[7], 0)
		Case "FolderModified"
			NoteGui($aNotes[7], 0)
		Case "FolderCreate"
			NoteGui($aNotes[7], 0)
		Case "Folder"
			NoteGui($aNotes[5], 0)
		Case "FileModified"
			NoteGui($aNotes[7], 0)
		Case "FileCreate"
			NoteGui($aNotes[7], 0)
		Case "FileWord"
			NoteGui($aNotes[6], 0)
		Case "FileSize"
			NoteGui($aNotes[7], 0)
		Case "File"
			NoteGui($aNotes[5], 0)
	EndSwitch
EndFunc   ;==>NotesMenu

Func ModelFileMenu($item)
	Local $MaxModelFileLine, $hModelFileOpen, $ModelMissing = True
	Local $aModelLines[1], $ObjectMatch = 0, $ii = 0 ; count for match model objects

	$aValue = Open_Count_Model_File()
	$hModelFileOpen = $aValue[0]
	$MaxModelFileLine = $aValue[1]

	$aNotes = _StringExplode($aFaultLog[$item], "#", 5) ; get fault log string item from the Fault.txt
	;_ArrayDisplay($aNotes, "Array for Location and object")
	$faultfolder = _StringBetween($aFaultLog[$item], "c$\", "#") ; get location for object +++++++++++++++++++++++++++++++++++++++++++++

	$ObjectSelected = $aNotes[2] ; Temp variable for the object error selected used for the mapdrive function
	$ObjectHeader = $aNotes[0] ; object header
	ReDim $aModelLines[$MaxModelFileLine]
	;MsgBox(0,"search parameter", $ObjectHeader & "#Y#"& $aNotes[2])

	For $i = 1 To $MaxModelFileLine
		$sFileRead = FileReadLine($hModelFileOpen, $i) ; read model file
		If StringInStr($sFileRead, $ObjectHeader & "#Y#" & $aNotes[2], 2) Then ; if string in fault log matches object in the model.txt, do statement  +++++ insert objectheader
			If StringInStr($sFileRead, $faultfolder[0], 2) Then ; make sure object in fault.txt has the same folder in model.txt; jus tin case files are the same name but diffrent folders
				$ModelMissing = False ; Error maintenance
				$aModelLines[$ii] = $sFileRead ; assign all matching objects from model to an array
				$ii += 1
				$ObjectMatch += 1
			EndIf
		EndIf
	Next
	FileClose($hModelFileOpen)
	;_ArrayDisplay($aModelLines, "Models found on model.txt")
	If $ModelMissing Then
		MsgBox(0, "Script Player v0.53", " No file model found for " & $aNotes[2] & ".")
	Else
		Local $sFileModelList = GetModelList($aModelLines, $ObjectMatch) ; format array as a string
		;MsgBox(0,"model list output", $sFileModelList)
		ObjectMenu($sFileModelList, "File Map") ; pass formated string to the menu
	EndIf


EndFunc   ;==>ModelFileMenu

Func ModelFolderMenu($item)

	Local $MaxModelFileLine, $hModelFileOpen, $ModelMissing = True
	Local $aModelLines[1], $ObjectMatch = 0, $ii = 0 ; count for match model objects

	$aValue = Open_Count_Model_File()
	$hModelFileOpen = $aValue[0]
	$MaxModelFileLine = $aValue[1]

	$aNotes = _StringExplode($aFaultLog[$item], "#", 7) ; get fault log string item from the Fault.txt
	$ObjectSelected = $aNotes[2] ; Temp variable for the object error selected used for the mapdrive function
	$ObjectHeader = $aNotes[0] ; object header
	ReDim $aModelLines[$MaxModelFileLine]
	For $i = 1 To $MaxModelFileLine
		$sFileRead = FileReadLine($hModelFileOpen, $i) ; read model file
		If StringInStr($sFileRead, $ObjectHeader & "#Y#" & $aNotes[2], 2) Then ; if string in fault log matches object in the model.txt, do statement
			$ModelMissing = False ; Error maintenance
			$aModelLines[$ii] = $sFileRead ; assign all math objects from model to an array
			$ii += 1
			$ObjectMatch += 1
		EndIf
	Next
	FileClose($hModelFileOpen)
	If $ModelMissing Then
		MsgBox(0, "Script Player v0.53", " No folder model found for " & $aNotes[2] & ".")
	Else
		Local $sFileModelList = GetModelList($aModelLines, $ObjectMatch)
		ObjectMenu($sFileModelList, "Folder Map")
	EndIf

EndFunc   ;==>ModelFolderMenu

Func ServiceStartMenu($item)
	Local $processStatus
	$ProcessNotes = _StringExplode($aFaultLog[$item], "#", 6)
	$serviceScript = "Get-Service -name '" & $ProcessNotes[2] & "' -ComputerName " & $remotecomputer & " | Set-Service -Status 'Running'"
	;$VerifyScript = " if (( Get-Service -name '" & $ProcessNotes[2] & "' -ComputerName " &$remotecomputer & " ).Status -eq 'Running')  {write-host '1'} Else { write-host '0'}"
	$VerifyScript = " $ProcessStatus = Get-Service -name '" & $ProcessNotes[2] & "' -ComputerName " & $remotecomputer & @CR & _
			" if  ($ProcessStatus.Status -eq 'running') { write-host 1} Else {write-host 0} "
	MsgBox(0, "Script Player v0.53",  " Starting process, please be patient.", 2)
	RunWait("powershell.exe -executionpolicy unrestricted -NoProfile  " & $serviceScript, @SystemDir, @SW_HIDE, 3)
	Sleep(300)
	$status = Run("powershell.exe -executionpolicy unrestricted  -NoProfile  " & $VerifyScript, @SystemDir, @SW_HIDE, 3)
	StdinWrite($status, @CRLF)
	StdinWrite($status)
	While 1
		$processStatus &= StdoutRead($status)
		If @error Then ExitLoop
	WEnd

	If $processStatus = 1 Then
		MsgBox(0, "Script Player v0.53", $ProcessNotes[2] & " running")
	Else
		MsgBox(16, "Script Player v0.53", "Not able to start " & $ProcessNotes[2] & " process. Please check service dependencies. ")
	EndIf
EndFunc   ;==>ServiceStartMenu

Func ServiceStopMenu($item)
	Local $processStatus
	$ProcessNotes = _StringExplode($aFaultLog[$item], "#", 6)
	$serviceScript = "Get-Service -name '" & $ProcessNotes[2] & "' -ComputerName " & $remotecomputer & " | Stop-Service  -force "
	;$VerifyScript = " if (( Get-Service -name '" & $ProcessNotes[2] & "' -ComputerName " &$remotecomputer & " ).Status -eq 'Running')  {write-host '1'} Else { write-host '0'}"
	$VerifyScript = " $ProcessStatus = Get-Service -name '" & $ProcessNotes[2] & "' -ComputerName " & $remotecomputer & @CR & _
			" if  ($ProcessStatus.Status -eq 'stopped') { write-host 1} Else {write-host 0} "
	RunWait("powershell.exe -executionpolicy unrestricted -NoProfile  " & $serviceScript, @SystemDir, @SW_HIDE, 3)
	Sleep(500)
	$status = Run("powershell.exe -executionpolicy unrestricted  -NoProfile  " & $VerifyScript, @SystemDir, @SW_HIDE, 3)
	StdinWrite($status, @CRLF)
	StdinWrite($status)
	While 1
		$processStatus &= StdoutRead($status)
		If @error Then ExitLoop
	WEnd
	If $processStatus = 1 Then
		MsgBox(0, "Script Player v0.53", $ProcessNotes[2] & " stopped")
	Else
		MsgBox(16, "Script Player v0.53", "not able to stop " & $ProcessNotes[2] & " process. Please check service dependencies. ")
	EndIf
EndFunc   ;==>ServiceStartMenu

Func SearchWebMenu($item) ; for searching patches
	;MsgBox(0,"Search Menu ", "Search Web for Patch  " & $item)
	$aNotes = _StringExplode($aFaultLog[$item], "#", 6)
	;_ArrayDisplay($aNotes) ; #2
	$PatchName = StringStripWS($aNotes[2], 3) ; remove leading and trailing white spaces
	If StringInStr($PatchName, " ") Then
		$PatchName = StringReplace($PatchName, " ", "%")
		; $PatchName = $PatchName & "+site%3Amicrosoft.com"
	EndIf
	$Browser = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice", "Progid") ; Read default browser settings
	If $Browser = "ChromeHTML" Then
		RunWait(@ComSpec & ' /c start chrome.exe https://support.microsoft.com/en-us/search?query="' & $PatchName & '"', @SystemDir, @SW_HIDE, 3)
	ElseIf $Browser = "IE.HTTP" Then
		ShellExecute("iexplore.exe", "about:blank")
		WinWait("Blank Page")
		$oIE = _IEAttach("about:blank", "url")
		_IELoadWait($oIE)
		_IENavigate($oIE, 'https://support.microsoft.com/en-us/search?query="' & $PatchName & '"')
	ElseIf $Browser = "FirefoxURL" Then
		RunWait(@ComSpec & ' /c start firefox  -new-tab  https://support.microsoft.com/en-us/search?query="' & $PatchName & '"', @SystemDir, @SW_HIDE, 3)
	EndIf
EndFunc   ;==>SearchWebMenu

Func Msg($text)
	MsgBox(0, $text, $text)
EndFunc   ;==>Msg

Func ArrayD($varArray)
	_ArrayDisplay($varArray, " Array")
EndFunc   ;==>ArrayD

Func MapDrive($TargetMachine, $location)
	;MsgBox(0, "Map function", "Machine : " & $TargetMachine & "  Location: " & $location)
	RunWait(@ComSpec & " /c net use \\" & $TargetMachine & "\c$", @SystemDir, @SW_HIDE,3)
	Sleep(1000)
	TrayTip(" ", "Opening Explorer",4)
	Sleep(1000)
	;msg(@WindowsDir & "\Explorer.exe /n, \\" & $TargetMachine & "\c$")
	;Run( @WindowsDir & "\Explorer.exe /n, \\" & $TargetMachine & "\c$\'" & $location & "'")
	Run( @WindowsDir & "\Explorer.exe /n, \\" & $TargetMachine & "\c$")
	;TrayTip(" ", "Mapping TARGET machine",3)
	;Run( @WindowsDir & "Explorer.exe /n, \\" &$remotecomputer & "\c$\") RemoteMachine
EndFunc   ;==>MapDrive

Func CleanModel() ; delete duplicate and sort array
	Local $aLocal[1]

	$aModelMaxLine = _FileCountLines($modelfile)
	If $aModelMaxLine = 0 Then ; if model.txt has no entries -there is noting to clean thus exiting
		Return
	Endif
	ReDim $aLocal[$aModelMaxLine]

	$hModelFileOpen = FileOpen($modelfile, 0) ; Read Mode

	If $hModelFileOpen = -1 Then
		MsgBox(0, "Script Player v0.53", "Not able to open Model.txt.")
		Return
	EndIf

	For $i = 1 To $aModelMaxLine ; convert txt file to array
		$sFileRead = FileReadLine($hModelFileOpen, $i)
		$aLocal[$i - 1] = $sFileRead
	Next
	FileClose($hModelFileOpen)
	Local $aUnique = _ArrayUnique($aLocal) ; delete duplicate entries

	_ArrayRemoveBlanks($aUnique)


	_ArraySort($aUnique, 1, 1) ;  Sort Array---  descending order, start at array 1

	$hNewModelFileOpen = FileOpen($modelfile, 2) ; Open model txt again in overwrite Mode

	For $i = 1 To UBound($aUnique) - 1
		FileWriteLine($hNewModelFileOpen, $aUnique[$i]) ; write to array
	Next
	FileClose($hNewModelFileOpen)

EndFunc   ;==>CleanModel

Func ModelToArray($aFaulItem, $count)
	Local $aModelObject[1][4]
	Local $aTemp[5], $machinename

	ReDim $aModelObject[$count][4]
	For $i = 0 To UBound($aFaulItem) - 1
		$aTemp = _StringExplode($aFaulItem[$i], "#")
		$aModelObject[$i][0] = $aTemp[2] ; assign filename, machine name, file location and date
		$machinename = _StringBetween($aTemp[3], "\\", "\c$") ; extract machine name on array[3]
		$aModelObject[$i][1] = $machinename[0] ; assign machiname array 0 to array model object 1
		$aModelObject[$i][2] = $aTemp[3] ; assign file location
		$aModelObject[$i][3] = $aTemp[4] ;  assign date
	Next
	Return $aModelObject
EndFunc   ;==>ModelToArray

Func GetModelList($aFaultItem, $count)
	Local $aModelObject[1]
	Local $aTemp[6], $machinename
	_ArrayRemoveBlanks($aFaultItem)
	;_ArrayDisplay($aFaultItem," Whole Item array");+++++++++++++++++++++++++++ show array coming into function
	ReDim $aModelObject[$count]
	For $i = 0 To UBound($aFaultItem) - 1
		$aTemp = _StringExplode($aFaultItem[$i], "#")
		; _ArrayDisplay($aTemp,"machine") ;++++++++++++++++++++++++++ show  line being read
		$machinename = _StringBetween($aTemp[3], "\\", "\c$") ; extract machine name on array[3]
		$aModelObject[$i] = $machinename[0] & "  (" & $aTemp[4] & ")" ; assign machiname array 0 and date the date
	Next
	;_arrayDisplay($aModelObject,"ModelList")
	Local $sModelList = _ArrayToString($aModelObject, "|")
	$DefaultListValue = $aModelObject[0]
	;MsgBox(0,"", $DefaultListValue)  ; default highlighted value for combo list
	Return $sModelList
EndFunc   ;==>GetModelList

Func ObjectMenu($sList, $header)
	$ObjectMapMenu = GUICreate($header, 245, 109, -1, -1, -1, $WS_EX_TOOLWINDOW, $GUIFaultMenu)
	$ModelChoice = GUICtrlCreateCombo("Model Machine (Date Logged)", 24, 24, 197, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
	;(0,"model list insde the object menu", $sList)
	GUICtrlSetData(-1, $sList, "1")
	$MapButton = GUICtrlCreateButton("Map To Model", 84, 64, 91, 25)
	GUISetState(@SW_DISABLE, $GUIFaultMenu)
	;GUISetState(@SW_MINIMIZE, $GUIFaultMenu)
	GUISetState(@SW_SHOW, $ObjectMapMenu)

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				;GUISetState(@SW_RESTORE, $GUIFaultMenu)
				GUISetState(@SW_ENABLE, $GUIFaultMenu)
				GUIDelete($ObjectMapMenu)
				ExitLoop
			Case $MapButton
				GetNametoMap()
		EndSwitch
	WEnd
EndFunc   ;==>ObjectMenu

Func GetNametoMap() ; get information from the model "machine and date" menu choice
	Local $ItemSelected = (GUICtrlRead($ModelChoice))
	$aCompname = _StringBetween($ItemSelected, "", "(")
	$aDate = _StringBetween($ItemSelected, "(", ")")

	If ModelMachineOnline($aCompname[0]) Then
		; msgbox(0, "Object Selected", $ObjectSelected & " machinename: " & $aCompname[0] & " Date: " & $aDate[0] )
		GetObjectLoc($ObjectSelected, $aCompname[0], $aDate[0])
	EndIf
EndFunc   ;==>GetNametoMap

Func ModelMachineOnline($compname)
	Local $online, $status, $TestOnlineScriptBase
	$TestOnlineScriptBase = "$status = Test-Connection -computername " & $compname & " -quiet" & @CR & _
			"if ( $status  -eq $true) {write-host '1'} Else { write-host '0'} "
	$status = Run("powershell.exe -executionpolicy unrestricted -NoProfile  " & $TestOnlineScriptBase, @SystemDir, @SW_HIDE, 3)
	MsgBox(0, "Script Player v0.53", "Checking if machine is online", 1)
	StdinWrite($status, @CRLF)
	StdinWrite($status)
	While 1
		$online &= StdoutRead($status)
		If @error Then ExitLoop
	WEnd
	If $online = 1 Then
		Return True
	Else
		MsgBox(16, "Script Player v0.53", "Machine not online")
		Return False
	EndIf
EndFunc   ;==>ModelMachineOnline

Func GetObjectLoc($ObjectName, $CompItem, $dateItem)

	$aValue = Open_Count_Model_File()
	$hModelFileOpen = $aValue[0]
	$MaxModelFileLine = $aValue[1]

	$CompItem = StringStripWS($CompItem, 3)
	$dateItem = StringStripWS($dateItem, 3)
	$ObjectName = StringStripWS($ObjectName, 3)

	For $i = 1 To $MaxModelFileLine
		$sFileRead = FileReadLine($hModelFileOpen, $i) ; read model file
		If StringInStr($sFileRead, $ObjectHeader & "#Y#" & $ObjectName & "#\\" & $CompItem, 2) Then ; if string has the right object header, object name and machine name in fault log matches object in the model.txt, do statement
			If StringInStr($sFileRead, $dateItem, 2) Then ; has the right date
				$fileLoc = _StringBetween($sFileRead, "c$\", "#")
				FileClose($hModelFileOpen)
				MapDrive($CompItem, $fileLoc[0])
				Return
			EndIf
		EndIf
	Next
	FileClose($hModelFileOpen)
EndFunc   ;==>GetObjectLoc

Func NoteGui($note, $OnlineSelected)
	Local $NoteGUI

	If NoteshasLink($note) Then
		;msg($note)
		$alink = _StringBetween($note, "http", " ")
		; _ArrayDisplay($aLink,"notes link")
		$notelink = "http" & $alink[0]
		$newnote = StringReplace($note, $notelink, " ( Click Open Link button )")
	Else
		$newnote = $note
	EndIf

	If $OnlineSelected Then
		$menuHeader = "Contact Menu"
	Else
		$menuHeader = "Script Notes"
	EndIf


	$NoteGUI = GUICreate($menuHeader, 300, 75, -1, -1, -1, $WS_EX_TOOLWINDOW)
	$g_hEdit = _GUICtrlEdit_Create($NoteGUI, $newnote, 2, 3, 294, 67, BitOR($ES_READONLY, $ES_MULTILINE))
	GUISetState(@SW_SHOW)
	;GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
	; _GUICtrlEdit_AppendText($g_hEdit, @CRLF & "Append to the end?")
	; Loop until the user exits.
	Do
	Until GUIGetMsg() = $GUI_EVENT_CLOSE
	GUIDelete($NoteGUI)
EndFunc   ;==>NoteGui

Func AutoClip()
	$MemValue = ClipGet()
	If Not StringIsSpace($remotecomputer) Then
		Return $remotecomputer
	ElseIf StringInStr($MemValue, "  ") Then
		Return ""
	Else
		Return $MemValue
	EndIf
EndFunc   ;==>AutoClip

Func NoteshasLink($NoteString)
	If StringInStr($NoteString, "https://", 0, -1) Then
		Return True
	ElseIf StringInStr($NoteString, "http://", 0, -1) Then
		Return True
	Else
		Return False
	EndIf
EndFunc   ;==>NoteshasLink

Func OpenLink($item) ; for searching patches
	$alink = _StringBetween($aFaultLog[$item], "http", " ")
	$notelink = "http" & $alink[0]
	OpenBrowser($notelink)
EndFunc   ;==>OpenLink

Func ObjectHasModel($ObjectName)
	Local $MaxModelFileLine, $modelObject = 0
	; msg($objectname)

	$aValue = Open_Count_Model_File()
	$hModelFileOpen = $aValue[0]
	$MaxModelFileLine = $aValue[1]

	For $i = 1 To $MaxModelFileLine
		$sFileRead = FileReadLine($hModelFileOpen, $i) ; read model file
		If StringInStr($sFileRead, $ObjectName) Then ;
			$modelObject = 1
		EndIf
	Next
	FileClose($hModelFileOpen)

	If $modelObject Then
		;msg("true model")
		Return True
	Else
		;msg("false model")
		Return False
	EndIf
EndFunc   ;==>ObjectHasModel

Func Open_Count_Model_File()
	Local $returnvalue[2]

	$hModelFileOpen = FileOpen($modelfile, 0) ; Open model file to be read
	If $hModelFileOpen = -1 Then ; trap error
		MsgBox($MB_SYSTEMMODAL, "Script Player v0.53", "Error reading Model File.")
		Return
	EndIf
	$returnvalue[0] = $hModelFileOpen
	$returnvalue[1] = _FileCountLines($modelfile) ; get max lines in model file

	Return $returnvalue
EndFunc   ;==>Open_Count_Model_File

Func OpenBrowser($link)
	$Browser = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice", "Progid") ; Read default browser settings
	If $Browser = "ChromeHTML" Then
		RunWait(@ComSpec & ' /c start chrome.exe ' & $link, @SystemDir, @SW_HIDE, 3)
	ElseIf $Browser = "IE.HTTP" Then
		ShellExecute("iexplore.exe", "about:blank")
		WinWait("Blank Page")
		$oIE = _IEAttach("about:blank", "url")
		_IELoadWait($oIE)
		_IENavigate($oIE, $link)
	ElseIf $Browser = "FirefoxURL" Then
		RunWait(@ComSpec & ' /c start firefox  -new-tab ' & $link, @SystemDir, @SW_HIDE, 3)
	Else
		ShellExecute("iexplore.exe", "about:blank")
		WinWait("Blank Page")
		$oIE = _IEAttach("about:blank", "url")
		_IELoadWait($oIE)
		_IENavigate($oIE, $link)

	EndIf
EndFunc   ;==>OpenBrowser

Func UpdateFaultCount($count)

	GUICtrlSetData($FaultLabel, $count)
	Sleep(50)
EndFunc   ;==>UpdateFaultCount

#Region Timer
Func MeasureTime($mode)
	If $mode = "Start" Then
		$htimer = TimerInit()
		Return
	EndIf

	If $mode = "Stop" Then
  		$DiffTime = TimerDiff($htimer)
		$atime = StringSplit($DiffTime, ".")
		$timevalue = Round($atime[1])
		Return $timevalue ; Round($atime[1]/1000)
	Endif
EndFunc

Func TimeFile($mode) ; open, close, read or write time file

	If Not FileExists($timelog) Then ; check for time.txt
		_FileCreate($timelog)
	EndIf

	If $mode = "Read" Then
		$hTimeFileOpen = FileOpen($timelog, 0) ; Read fault.txt on Read Mode

	ElseIf $mode = "Write" Then
		$hTimeFileOpen = FileOpen($timelog, 1) ; Read fault.txt on Read Append

	ElseIf $mode = "Close" Then
		FileClose($hTimeFileOpen)
	EndIf

	If $hTimeFileOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "Script Player v0.53", "Error processing Time File.")
		Return
	EndIf

EndFunc ; TimeFile

Func WriteTimeLog($aFileTime)
	Local $ScriptMissing = 1

	$MaxFileLine = _FileCountLines($timelog)

	For $ii = 0 To UBound($aFileTime) - 1 ; get row count minus 1
		TimeFile("Read")
		For $i = 1 To $MaxFileLine
			$sFileRead = FileReadLine($hTimeFileOpen, $i)
			;msgbox(0,"", " Line being read: " & $sFileRead & ". i= " & $i )
			If StringInStr($sFileRead,$aFileTime[$ii][0], 2) Then ; if filescript is found in the log
				$aLogdata = StringSplit($sFileRead,"#") ; separate the filename and time logged
				If $aFileTime[$ii][1]> $aLogdata[2]  Then                                  ;($aFileTime[0][1] > $aLogdata[2]) Then
					_FileWriteToLine($timelog, $i, $aFileTime[$ii][0] & "#" & $aFileTime[$ii][1],True)
					;msg("Overwrite Time ")
				ElseIf $aFileTime[$ii][1] < $aLogdata[2] Then  ;($aFileTime[0][1] < $aLogdata[2]) Then
					;msg("Do Nothing ")
				Endif
				$ScriptMissing = 0
				ExitLoop
			EndIf
		Next
		TimeFile("Close")
		If $ScriptMissing Then
			TimeFile("Write")
			FileWrite($hTimeFileOpen, $aFileTime[$ii][0] & "#" & $aFileTime[$ii][1] & @CRLF)
			TimeFile("Close")
		Endif
		$ScriptMissing = 1 ; initialize value again
	Next

EndFunc

Func ReadTimeLog()
	Local $TotalTime
	Local $NewScript = 1
	Local $postfix = ""
	Local $MaxRow
	$MaxRow = _GUICtrlListBox_GetListBoxInfo($RunList) ; count the scripts being run
	$MaxFileLine = _FileCountLines($timelog) ; count lines on the time log
    TimeFile("Read")
		GUICtrlSetData($ScriptLabel, $MaxRow)
		For $i = 0 To $MaxRow - 1 Step 1
			$FileScript = _GUICtrlListBox_GetText($RunList, $i)
			$NewScript = 1
			For $ii = 1 To $MaxFileLine
				$sFileRead = FileReadLine($hTimeFileOpen, $ii)
				If StringInStr($sFileRead,$FileScript, 2,1) Then ; if filescript is found in the log
					$aLogdata = StringSplit($sFileRead,"#") ; separate the filename and time logged
					$TotalTime += $aLogdata[2]
					$NewScript = 0 ; script being ran is not new
					ExitLoop(1)
				Endif
			Next
			If $NewScript Then
				$postfix = "+"
			Endif

		Next
	TimeFile("Close")
	If $TotalTime = "" Then
		GUICtrlSetData($EstTimeLabel, "99:99")
	ElseIf $postfix = "+" Then
		GUICtrlSetData($EstTimeLabel, ConvertTime($TotalTime) & $postfix )
	Else
		GUICtrlSetData($EstTimeLabel, ConvertTime($TotalTime) )
	EndIf
EndFunc

Func ConvertTime($milliseconds)
	$seconds = ( 3 * $ScriptCount) + Round($milliseconds/1000) ; add 5 seconds per script being run
   ;msg("seconds = " & $seconds)
	If $seconds >=  60 Then
		$scripttime = Int($seconds / 60)
		$remainder = Mod($seconds , 60)
		$RunTime = $scripttime & ":" & PadTime($remainder)
		Return $RunTime
	ElseIf $seconds = 0 Then
		$Runtime = "." & $milliseconds & "s"
		Return $RunTime
	Else
		$RunTime = $seconds & "s"
		Return $RunTime
	Endif
EndFunc

Func PadTime($num)
	If $num < 10 Then
			Return "0" & $num
	Else
			Return $num
	EndIf
EndFunc

#EndRegion Timer

Func CreateNote() ; activite note window in scriptbuilder
			If Not FileExists($notefile) Then ; check for fault.txt
				_FileCreate($notefile)
			MsgBox(0,"Script Player v.051", "Note Feature Activated")
			Endif
		Return
EndFunc   ;==>HotKeyPressed

Func NotExpired()

	If $TrialEndDate >_NowCalc() Then
		;MsgBox(0,"","Not Expired")
		Return True
	Else
		MsgBox(64,"Script Builder v0.53" , "Trial Date of " &  _DateTimeFormat($TrialEndDate, 2) & " already expired.")
		Return False
	EndIf
EndFunc

Func ProcessSampleScript()  ;==>Process Sample Scripts to the right path

	Local  $MaxFileLine
	Local $outfile =  @ScriptDir & "\Util\fault.txt"
	Local $SentinelFile =  @ScriptDir & "\Util\SampleProcessed.txt"
	Local	$aPathFile[8]
    $aPathFile[0] = @ScriptDir & "\App\Google Drive Win7 64 bit.ps1" ; 64 bit
	$aPathFile[1] = @ScriptDir & "\App\IE11 MSFT Win7 64 bit.ps1"; 64 bit
	$aPathFile[2] = @ScriptDir & "\Task\MozillaFirefox v42 AutoUpdate Win7 64 bit.ps1"; 64 bit
	$aPathFile[3] = @ScriptDir & "\Task\GoogleDrive Synch Task Win7 64 bit.ps1"; 64 bit
	$aPathFile[4] = @ScriptDir & "\Symp\IE_CHROME_Firefox Default Browser issues.ps1" ; 64 bit
	$aPathFile[5] = @ScriptDir & "\Symp\SlowWin7 Check Services Not Needed.ps1"; 64 bit and 32 bit
	$aPathFile[6] = @ScriptDir & "\Task\Sample Printing Script.ps1";  32 bit
	$aPathFile[7] = @ScriptDir & "\App\Sample Microsoft Office 365 script.ps1";  32 bit


   If FileExists($SentinelFile) Then
	  ; MsgBox(0,"","Sample scripts already processed")
	   Return
	EndIf
    ; Open the file for read/write access.
	For $i = 0 To 7
		;MsgBox(0,"", "File being opened is: " & $aPathFile[$i])
		If FileExists($aPathFile[$i]) Then
			;MsgBox(0,"", "File being opened is: " & $aPathFile[$i])
			$hFileOpen = FileOpen($aPathFile[$i],  0 ) ; 0 = read mode , 1 = append mode
			If $hFileOpen = -1 Then
				MsgBox($MB_SYSTEMMODAL, "", "An error occurred opening sample script: " & $aPathFile[$i]  )
				Return False
			EndIf
			$MaxFileLine = _FileCountLines($aPathFile[$i])
			For $ii = 1 To $MaxFileLine
				$sFileRead = FileReadLine($hFileOpen, $ii)
				If StringInStr($sFileRead, "spkg", 2) Then
					;MsgBox(0,""," Found instance of 123.txt")
					;MsgBox(0,"",$sFileRead)
					$UpdatedPath = StringReplace($sFileRead,"spkg",$outfile, 2)  ;C:\Program Files\ITTM\Util\fault.txt
					_FileWriteToLine($aPathFile[$i], $ii,$UpdatedPath, True ) ; true = overwrite old line
					;MsgBox(0,"",$UpdatedPath)
				Endif
			Next
			FileClose($hFileOpen)
		Endif
	Next
_FileCreate($SentinelFile)

EndFunc   ;==>Process Sample Scripts to the right path
