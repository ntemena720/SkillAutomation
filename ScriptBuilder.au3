#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <GuiEdit.au3>
#include <ButtonConstants.au3>
#include <DateTimeConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <ComboConstants.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <Date.au3>
#include <MsgBoxConstants.au3>
#include <Constants.au3>
#include <WinAPI.au3>
#include <ListBoxConstants.au3>
#include <GuiListBox.au3>
#include <IE.au3>
#RequireAdmin ; needed for registry and hotfix scripts to run

#Region Variables  ; Heather-Think 192.168.1.94
Global $TrialEndDate = "2017/07/04" ; 00:00:00" ; YYYY/MM/DD Trial Copy
Global $remotecomputer, $TempScript
Global $outfile = '"' & @ScriptDir & '\Util\fault.txt"'
Global $AppFolder = @ScriptDir & "\App\"
Global $AppFiles =  @ScriptDir & "\App\*.ps1" ; for append or overwrite mode
Global $TaskFolder =  @ScriptDir & "\Task\"
Global $TaskFiles = @ScriptDir & "\Task\*.ps1" ; for append or overwrite mode
Global $SympFolder =  @ScriptDir & "\Symp\"
Global $SympFiles = @ScriptDir & "\Symp\*.ps1" ; for append or overwrite mode
Global $faultfile = @ScriptDir & "\Util\fault.txt"
Global $modelfile =  @ScriptDir & "\Util\model.txt"
Global $notefile =  @ScriptDir & "\Util\note.txt"
Global $Tempfile = @ScriptDir & "\Util\tempfile.ps1"

Global $hGUI, $h_EditGUI, $idADD, $ReadList, $SaveClose, $FileRadio, $FolderRadio, $ManualLink, $BuyLink
Global $FileGUI, $FileSize, $Filename, $ModifiedDate, $FileAddButton = 999, $FileTestButton = 999, $FilePath, $CDate, $SWord, $MdateBox, $CDateBox, $FSizeBox, $SWordBox, $FileScriptBase
Global $FolderGUI, $FolderName, $FolderAddButton = 999, $FolderTestButton = 999, $FolderCDate, $FolderMDate, $FolderNumFiles, $FolderLoc, $FolderCDateBox, $FolderMDateBox, $NumFilesBox, $FolderScriptBase
Global $RegistryGUI, $RegAddButton = 999, $RegTestButton = 999, $RegKey, $RegValue, $RegHive, $RegData, $RegistryScriptBase
Global $ProcessGUI, $ProcessAddButton = 999, $ProcessTestButton = 999, $ProcessName, $ProcessScriptBase
Global $ServiceGUI, $ServiceAddButton = 999, $ServiceTestButton = 999, $ServiceName, $ServiceScriptBase, $ServiceState
Global $OnlineGUI, $OnlineAddButton = 999, $OnlineTestButton = 999, $OnlineName, $OnlineScriptBase
Global $HotfixGUI, $HotfixAddButton = 999, $HotfixTestButton = 999, $HotfixName, $HotfixScriptBase
Global $EventGUI, $EventDone = 999, $LogName, $EntryName, $SearchWord
Global $SavetoFileGUI, $ProfileType, $ScriptFileName, $SavetoFilebutton = 999
Global $FileDirGUI, $FileDirButton = 999, $ProfileCombo, $FileList, $SelDir, $filenumber
Global $RunningScript = " ", $UniqueFileNameScript, $FirstTimeRun = 1 ; for testing temp script
Global $NoteText, $StringNumber ; Note Menu
Global $objectnotes = "None"

#EndRegion Variables  ; Heather-Think 192.168.1.94

#Region Menu
CheckProfileFolder()

$hGUI = GUICreate("Script Builder v0.53", 409, 315, -1, -1, -1, $WS_EX_TOOLWINDOW)
GUICtrlSetBkColor(-1, 0xFF0000)
CreateWindowMenu()
$FileRadio = GUICtrlCreateRadio("File", 24, 24, 57, 17)
$FolderRadio = GUICtrlCreateRadio("Folder", 24, 48, 57, 17)
$RegistryRadio = GUICtrlCreateRadio("Registry", 24, 72, 57, 17)
$ProcessRadio = GUICtrlCreateRadio("Process", 24, 96, 57, 17)
$ServiceRadio = GUICtrlCreateRadio("Service", 24, 120, 57, 17)
$OnlineRadio = GUICtrlCreateRadio("Connectivity", 24, 144, 78, 17)
$HotfixRadio = GUICtrlCreateRadio("HotFix", 24, 168, 57, 17)
$TestScript = GUICtrlCreateButton("Test Script", 24, 200, 100, 35)
$Close = GUICtrlCreateButton("Exit", 278, 200, 100, 35)
$ClearScript = GUICtrlCreateButton("Clear Script", 152, 200, 100, 35)
$SaveScript = GUICtrlCreateButton("Save as New File", 152, 245, 100, 35)
GUICtrlSetState($SaveScript, $GUI_DEFBUTTON)
$AppendB = GUICtrlCreateButton("Append to File", 24, 245, 100, 35)
$OverwriteB = GUICtrlCreateButton("Overwrite File", 278, 245, 100, 35)

$h_EditGUI = _GUICtrlEdit_Create($hGUI, "", 147, 24, 229, 160, BitOR($ES_READONLY, $ES_MULTILINE, $WS_VSCROLL, $WS_HSCROLL))
GUISetState(@SW_SHOW, $hGUI)
CheckExpiration()


#EndRegion Menu

While 1
	$aMsg = GUIGetMsg(1)
	Switch $aMsg[1]
		Case $hGUI
			Switch $aMsg[0]
				Case $GUI_EVENT_CLOSE, $Close
					EraseTempFile() ; delete temp file created when testing scripts
					Exit
				Case $FileRadio
					GUISetState(@SW_DISABLE, $hGUI)
					FileMenu()
				Case $FolderRadio
					GUISetState(@SW_DISABLE, $hGUI)
					FolderMenu()
				Case $RegistryRadio
					GUISetState(@SW_DISABLE, $hGUI)
					RegistryMenu()
				Case $ProcessRadio
					GUISetState(@SW_DISABLE, $hGUI)
					ProcessMenu()
				Case $ServiceRadio
					GUISetState(@SW_DISABLE, $hGUI)
					ServiceMenu()
				Case $OnlineRadio
					GUISetState(@SW_DISABLE, $hGUI)
					OnLineMenu()
				Case $HotfixRadio
					GUISetState(@SW_DISABLE, $hGUI)
					HotfixMenu()
				Case $SaveScript
					GUISetState(@SW_DISABLE, $hGUI)
					SavetoFileMenu()
				Case $TestScript
					If RunScriptLoaded() Then
						If CheckMachineOnline() Then
							RunFilePS()
						EndIf
					Endif
				Case $AppendB
					FileDirMenu("Append")
				Case $OverwriteB
					FileDirMenu("Overwrite")
				Case $ClearScript
					$RunningScript = " " ; clear running list
					_GUICtrlEdit_SetText($h_EditGUI, "") ; clear edit box
				Case $ManualLink
					OpenBrowser("https://www.i-ttm.com/script-builder-manual.html") ;
				Case $BuyLink
					OpenBrowser("https://www.i-ttm.com/store/c1/Featured_Products.html") ;
			EndSwitch
		Case $FileGUI
			Switch $aMsg[0]
				Case $GUI_EVENT_CLOSE
					GUIDelete($FileGUI)
					GUISetState(@SW_ENABLE, $hGUI)
				Case $FileAddButton
					If GetNote("File") Then
						GUISetState(@SW_ENABLE, $hGUI);
						FileScript(1)
						$RunningScript &= $FileScriptBase
						GUIDelete($FileGUI)
					Endif
				Case $FileTestButton
					If CheckMachineOnline() Then
						FileScript(0)
						RunPS($FileScriptBase)
					EndIf
			EndSwitch
		Case $FolderGUI
			Switch $aMsg[0]
				Case $GUI_EVENT_CLOSE
					GUIDelete($FolderGUI)
					GUISetState(@SW_ENABLE, $hGUI)
				Case $FolderAddButton
					If GetNote("Folder") Then
						GUISetState(@SW_ENABLE, $hGUI);
						FolderScript(1)
						$RunningScript &= $FolderScriptBase
						GUIDelete($FolderGUI)
					EndIf
				Case $FolderTestButton
					If CheckMachineOnline() Then
						FolderScript(0)
						RunPS($FolderScriptBase)
					EndIf
			EndSwitch
		Case $RegistryGUI
			Switch $aMsg[0]
				Case $GUI_EVENT_CLOSE
					GUIDelete($RegistryGUI)
					GUISetState(@SW_ENABLE, $hGUI)
				Case $RegAddButton
					If GetNote("Registry") Then
						GUISetState(@SW_ENABLE, $hGUI);
						RegistryScript(1)
						;MsgBox(0, "add", $RegistryScriptBase)
						$RunningScript &= $RegistryScriptBase
						GUIDelete($RegistryGUI)
					EndIf
				Case $RegTestButton
					If CheckMachineOnline() Then
						RegistryScript(0)
						RunPS($RegistryScriptBase)
					EndIf
			EndSwitch
		Case $ProcessGUI
			Switch $aMsg[0]
				Case $GUI_EVENT_CLOSE
					GUIDelete($ProcessGUI)
					GUISetState(@SW_ENABLE, $hGUI)
				Case $ProcessAddButton
					If GetNote("Process") Then
						GUISetState(@SW_ENABLE, $hGUI)
						ProcessScript(1)
						$RunningScript &= $ProcessScriptBase
						GUIDelete($ProcessGUI)
					EndIf
				Case $ProcessTestButton
					If CheckMachineOnline() Then
						ProcessScript(0)
						RunPS($ProcessScriptBase)
					EndIf
			EndSwitch
		Case $ServiceGUI
			Switch $aMsg[0]
				Case $GUI_EVENT_CLOSE
					GUIDelete($ServiceGUI)
					GUISetState(@SW_ENABLE, $hGUI)
				Case $ServiceAddButton
					If GetNote("Service") Then
						GUISetState(@SW_ENABLE, $hGUI)
						ServiceScript(1)
						$RunningScript &= $ServiceScriptBase
						GUIDelete($ServiceGUI)
					EndIf
				Case $ServiceTestButton
					If CheckMachineOnline() Then
						ServiceScript(0)
						RunPS($ServiceScriptBase)
					EndIf
			EndSwitch
		Case $OnlineGUI
			Switch $aMsg[0]
				Case $GUI_EVENT_CLOSE
					GUIDelete($OnlineGUI)
					GUISetState(@SW_ENABLE, $hGUI)
				Case $OnlineAddButton
					If 	GetNote("Online") Then
						GUISetState(@SW_ENABLE, $hGUI)
						OnlineScript(1)
						$RunningScript &= $OnlineScriptBase
						GUIDelete($OnlineGUI)
					Endif
				Case $OnlineTestButton
					OnlineScript(0)
					RunPS($OnlineScriptBase)
			EndSwitch
		Case $HotfixGUI
			Switch $aMsg[0]
				Case $GUI_EVENT_CLOSE
					GUIDelete($HotfixGUI)
					GUISetState(@SW_ENABLE, $hGUI)
				Case $HotfixAddButton
					If GetNote("Hotfix") Then
						GUISetState(@SW_ENABLE, $hGUI)
						HotfixScript(1)
						$RunningScript &= $HotfixScriptBase
						GUIDelete($HotfixGUI)
					EndIf
				Case $HotfixTestButton
					If CheckMachineOnline() Then
						HotfixScript(0)
						RunPS($HotfixScriptBase)
					EndIf
			EndSwitch
		Case $SavetoFileGUI
			Switch $aMsg[0]
				Case $GUI_EVENT_CLOSE
					GUISetState(@SW_ENABLE, $hGUI)
					GUIDelete($SavetoFileGUI)
				Case $SavetoFilebutton
					CheckSaveFilename()
					SavetoModel()
			EndSwitch
	EndSwitch
WEnd

#Region Object Menu

Func FileMenu()
	$FileGUI = GUICreate("File Script Builder", 343, 279, -1, -1, -1, $WS_EX_TOOLWINDOW, $hGUI)
	GUICtrlCreateLabel("File Path", 16, 16, 118, 17, $SS_RIGHT)
	$FilePath = GUICtrlCreateInput("c:\root\foldername", 152, 16, 169, 21)
	GUICtrlCreateLabel("FileName", 96, 48, 48, 17, $SS_RIGHT)
	$Filename = GUICtrlCreateInput("Filename", 152, 48, 169, 21)
	GUICtrlCreateLabel("Last Modified Date", 40, 80, 99, 17, $SS_RIGHT)
	$ModifiedDate = GUICtrlCreateDate("", 152, 80, 90, 21, $DTS_SHORTDATEFORMAT)
	GUICtrlCreateLabel("Create Date", 80, 112, 61, 17, $SS_RIGHT)
	$CDate = GUICtrlCreateDate("", 152, 112, 90, 21, $DTS_SHORTDATEFORMAT)

	GUICtrlCreateLabel("File Size in KB", 56, 144, 81, 17, $SS_RIGHT)
	$FileSize = GUICtrlCreateInput("1,000 ", 152, 144, 89, 21)
	GUICtrlCreateLabel("Search for the word:", 48, 176, 100, 17)
	$SWord = GUICtrlCreateInput("Error", 152, 176, 89, 21)
	$FileAddButton = GUICtrlCreateButton("Add Script", 205, 210, 91, 33)
	$FileTestButton = GUICtrlCreateButton("Test Run Script", 59, 210, 91, 33)
	GUICtrlSetState($FileTestButton, $GUI_DEFBUTTON)
	$MdateBox = GUICtrlCreateCheckbox("", 24, 80, 17, 17)
	$CDateBox = GUICtrlCreateCheckbox("", 24, 112, 17, 17)
	$FSizeBox = GUICtrlCreateCheckbox("", 24, 144, 17, 17)
	$SWordBox = GUICtrlCreateCheckbox("", 24, 176, 17, 17)

	GUISetState(@SW_SHOW, $FileGUI)
EndFunc   ;==>FileMenu

Func FolderMenu()
	$FolderGUI = GUICreate("Folder Script Builder", 385, 211, -1, -1, -1, $WS_EX_TOOLWINDOW, $hGUI)
	GUICtrlCreateLabel("Folder Location", 94, 16, 78, 17, $SS_RIGHT)
	$FolderLoc = GUICtrlCreateInput("c:\", 184, 8, 185, 21)
	GUICtrlCreateLabel("Folder Name", 102, 40, 70, 17, $SS_RIGHT)
	$FolderName = GUICtrlCreateInput("foldername", 184, 40, 185, 21)
	GUICtrlCreateLabel("Create Date", 104, 72, 68, 17, $SS_RIGHT)
	$FolderCDate = GUICtrlCreateDate("", 184, 72, 90, 21, $DTS_SHORTDATEFORMAT)
	GUICtrlCreateLabel("Modified Date", 68, 104, 105, 17, $SS_RIGHT)
	$FolderMDate = GUICtrlCreateDate("", 184, 104, 90, 21, $DTS_SHORTDATEFORMAT)
	GUICtrlCreateLabel("# Subfolders/Files", 60, 136, 109, 17, $SS_RIGHT)
	$FolderNumFiles = GUICtrlCreateInput("00", 184, 136, 89, 21)
	$FolderAddButton = GUICtrlCreateButton("Add Script", 232, 166, 91, 33)
	$FolderTestButton = GUICtrlCreateButton("Test Run Script", 83, 166, 91, 33)
	GUICtrlSetState($FolderTestButton, $GUI_DEFBUTTON)
	$FolderCDateBox = GUICtrlCreateCheckbox("", 48, 72, 17, 17)
	$FolderMDateBox = GUICtrlCreateCheckbox("", 48, 104, 17, 17)
	$NumFilesBox = GUICtrlCreateCheckbox("", 48, 136, 17, 17)

	GUISetState(@SW_SHOW, $FolderGUI)
EndFunc   ;==>FolderMenu

Func RegistryMenu()
	$RegistryGUI = GUICreate("Registry Script Builder", 438, 213, -1, -1, -1, $WS_EX_TOOLWINDOW, $hGUI)
	GUICtrlCreateLabel("Hive", 64, 16, 26, 17)
	$RegHive = GUICtrlCreateCombo("LocalMachine", 104, 16, 121, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
	GUICtrlSetData($RegHive, "CurrentUser", "LocalMachine")
	GUICtrlCreateLabel("Key Name", 32, 48, 59, 17)
	$RegKey = GUICtrlCreateInput("SOFTWARE\Microsoft\Windows\CurrentVersion", 104, 48, 321, 21)
	GUICtrlCreateLabel("Value", 56, 80, 31, 17)
	$RegValue = GUICtrlCreateInput("ProgramFilesDir", 104, 80, 121, 21)
	GUICtrlCreateLabel("Expected Data", 8, 112, 85, 17)
	$RegData = GUICtrlCreateInput("C:\Program Files", 104, 112, 121, 21)
	$RegAddButton = GUICtrlCreateButton("Add Script", 264, 144, 91, 33)
	$RegTestButton = GUICtrlCreateButton("Test Run Script", 96, 144, 91, 33)
	GUICtrlSetState($RegTestButton, $GUI_DEFBUTTON)
	GUISetState(@SW_SHOW, $RegistryGUI)
EndFunc   ;==>RegistryMenu

Func ProcessMenu()
	$ProcessGUI = GUICreate("Process Script Builder", 240, 145, -1, -1, -1,$WS_EX_TOOLWINDOW, $hGUI)
	GUICtrlCreateLabel("Process Name", 88, 20, 73, 17)
	$ProcessName = GUICtrlCreateInput("Explorer", 64, 40, 121, 21)
	$ProcessAddButton = GUICtrlCreateButton("Add Script", 136, 80, 91, 33)
	$ProcessTestButton = GUICtrlCreateButton("Test Run Script", 16, 80, 91, 33)
	GUICtrlSetState($ProcessTestButton, $GUI_DEFBUTTON)
	GUISetState(@SW_SHOW, $ProcessGUI)
EndFunc   ;==>ProcessMenu

Func ServiceMenu()
	$ServiceGUI = GUICreate("Service Script Builder", 241, 175, -1, -1, -1, $WS_EX_TOOLWINDOW, $hGUI)
	GUICtrlCreateLabel("Service Name", 93, 20, 73, 17)
	$ServiceName = GUICtrlCreateInput("Print Spooler", 67, 40, 121, 21)
	GUICtrlCreateLabel("Expected State", 90, 72, 80, 17)
	$ServiceState = GUICtrlCreateCombo("Running", 88, 96, 75, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
	GUICtrlSetData($ServiceState, "Stopped", "Running")
	$ServiceAddButton = GUICtrlCreateButton("Add Script", 136, 128, 91, 33)
	$ServiceTestButton = GUICtrlCreateButton("Test Run Script", 24, 128, 91, 33)
	GUICtrlSetState($ServiceTestButton, $GUI_DEFBUTTON)
	GUISetState(@SW_SHOW, $ServiceGUI)

EndFunc   ;==>ServiceMenu

Func OnLineMenu()
	$OnlineGUI = GUICreate("Connectivity Script Builder", 241, 146, -1, -1, -1, $WS_EX_TOOLWINDOW, $hGUI)
	GUICtrlCreateLabel("Web Link", 100, 20, 97, 17)
	$OnlineName = GUICtrlCreateInput("http://", 64, 40, 121, 21)
	$OnlineAddButton = GUICtrlCreateButton("Add Script", 136, 80, 91, 33)
	$OnlineTestButton = GUICtrlCreateButton("Test Run Script", 16, 80, 91, 33)
	GUICtrlSetState($OnlineTestButton, $GUI_DEFBUTTON)
	GUISetState(@SW_SHOW, $OnlineGUI)
EndFunc   ;==>OnLineMenu

Func HotfixMenu()
	$HotfixGUI = GUICreate("HotFix Script Builder", 241, 146, -1, -1, -1, $WS_EX_TOOLWINDOW, $hGUI)
	GUICtrlCreateLabel("Patch Name", 93, 20, 97, 17)
	$HotfixName = GUICtrlCreateInput("KBXXXXXXX", 58, 40, 121, 21)
	$HotfixAddButton = GUICtrlCreateButton("Add Script", 136, 80, 91, 33)
	$HotfixTestButton = GUICtrlCreateButton("Test Run Script", 16, 80, 91, 33)
	GUICtrlSetState($HotfixTestButton, $GUI_DEFBUTTON)
	GUISetState(@SW_SHOW, $HotfixGUI)
EndFunc   ;==>HotfixMenu

#EndRegion Object Menu

#Region Object Scripts
Func FileScript($show)

	Local $filepathv1 = StringReplace(GUICtrlRead($FilePath), "C:\", "", 0, 2)
	Local $filenamev = StringReplace(GUICtrlRead($Filename), "C:\", "", 0, 2)
	Local $filesizev = GUICtrlRead($FileSize)
	Local $filemdatev = GUICtrlRead($ModifiedDate)
	Local $filecdatev = GUICtrlRead($CDate)
	Local $filewordv = GUICtrlRead($SWord)
	Local $FileSizeLine

	Local $ShowFilePath = StringReplace(GUICtrlRead($FilePath) & "\" & $filenamev, "\\", "\", 0, 2)

	$filepathv1_ = $filepathv1 ; if no [user] detected - assign them as the same



		If 	StringInStr($filepathv1,"[user]", 0) Then ; if userprofile dependent
			$filepathv1_ = StringReplace($filepathv1, "[user]","'+ $userfoldername + '") ; syntax for ps comments new syntax
			;$filepathv1out = StringReplace($filepathv1, "[user]","$userfoldername") ; syntax for outfile assignments
			$filepathv1 = StringReplace($filepathv1, "[user]","'$userfoldername'") ; syntax for ps commands
		Endif

		If ($show) Then
		_GUICtrlEdit_AppendText($h_EditGUI, "Filename to test: " & $ShowFilePath & @CRLF)
	EndIf

	If _IsChecked($FSizeBox) Then

		$FileSizeLine = "if ( $filelength -eq " & $filesizev & " ) " & @CR & _
				"  { " & @CR & _
				"   $fileD = 'FileSize#Y#" & $filenamev & "#\\" & $remotecomputer & "\c$\" & $filepathv1_ & "#'+  $datestamp +'" & "#' + $filelength + '#" & $filesizev & "'" & @CR & _    ; FileName,Size,[Y,N],Found Value, Model Value,time stamp
				"   write-host '" & $filenamev & " file has the correct file size of " & $filesizev & " ' " & @CR & _
				"    }  " & @CR & _
				"Else " & @CR & _
				"   { " & @CR & _
				"   $fileD = 'FileSize#N#" & $filenamev & "#\\" & $remotecomputer & "\c$\" & $filepathv1_ & "#' + $filelength + '#" & $filesizev & "#'+  $datestamp +'#" & $objectnotes & " '" & @CR & _ ; FileName,Size,[Y,N],Found Value, Model Value,time stamp
				"   write-host '" & $filenamev & " file has the incorrect file size of' $filelength'. Expected size is " & $filesizev & " '  -ForegroundColor red " & @CR & _
				"    } " & @CR & _
				" Out-file -filepath " & $outfile & " -append -inputObject $fileD "

		If ($show) Then
			_GUICtrlEdit_AppendText($h_EditGUI, "File size: " & $filesizev & @CRLF)
		EndIf
	Else

		$FileSizeLine = ""
	EndIf

	If _IsChecked($SWordBox) Then
		$FileSearchLine = "if ($content -like '*" & $filewordv & "*' ) " & @CR & _
				"  { " & @CR & _
				"   $fileA = 'FileWord#Y#" & $filenamev & "#\\" & $remotecomputer & "\c$\" & $filepathv1_ & "#'+  $datestamp +'" & "#" & $filewordv & "'" & @CR & _  ;FileWord,[Y,N]  filename ,Model Word,  $datestamp
				"   write-host '" & $filenamev & " file has the word " & $filewordv & " ' " & @CR & _
				"    } " & @CR & _
				"Else " & @CR & _
				"   { " & @CR & _
				"   $fileA = 'FileWord#N#" & $filenamev & "#\\" & $remotecomputer & "\c$\" & $filepathv1_ & "#" & $filewordv & "#'+  $datestamp +'#" & $objectnotes & " '" & @CR & _  ;FileWord,[Y,N]  filename ,Model Word,  $datestamp
				"   write-host '" & $filenamev & " file does not contain the word " & $filewordv & " ' -ForegroundColor red " & @CR & _
				"    } " & @CR & _
				"Out-file -filepath " & $outfile & " -append -inputObject $fileA  "
		If ($show) Then
			_GUICtrlEdit_AppendText($h_EditGUI, "File Word Search: " & $filewordv & @CRLF)
		EndIf
	Else
		$FileSearchLine = ""
	EndIf

	If _IsChecked($CDateBox) Then
		$FileCLine = "if ( $Cdate -eq '" & $filecdatev & "' ) " & @CR & _
				"  { " & @CR & _
				"   $fileC = 'FileCreate#Y#" & $filenamev & "#\\" & $remotecomputer & "\c$\" & $filepathv1_ & "#'+  $datestamp +'" & "#' + $Cdate +'#" & $filecdatev & "'" & @CR & _ ;FileCreate,[Y,N],& $filenamev &,Found Value, Model Value, Datestamp
				"   write-host '" & $filenamev & " file has the correct create date of " & $filecdatev & " ' } " & @CR & _
				"Else " & @CR & _
				"   { " & @CR & _
				"   $fileC = 'FileCreate#N#" & $filenamev & "#\\" & $remotecomputer & "\c$\" & $filepathv1_ & "#' + $Cdate +'#" & $filecdatev & "#'+  $datestamp +'#" & $objectnotes & " '" & @CR & _ ;FileCreate,[Y,N],& $filenamev &,Found Value, Model Value, Datestamp
				"	write-host '" & $filenamev & " file has the wrong create date of' $Cdate'. Expected date " & $filecdatev & " ' -ForegroundColor red " & @CR & _
				"    } " & @CR & _
				"Out-file -filepath " & $outfile & " -append -inputObject $fileC "
		If ($show) Then
			_GUICtrlEdit_AppendText($h_EditGUI, "File Create Date: " & $filecdatev & @CRLF)
		EndIf
	Else
		$FileCLine = ""
	EndIf

	If _IsChecked($MdateBox) Then ;$MdateBox, $CDateBox,$FSizeBox,$SWordBox
		$FileMLine = "if ( $Mdate -eq '" & $filemdatev & "' ) " & @CR & _
				"  { " & @CR & _
				"   $fileB = 'FileModified#Y#" & $filenamev & "#\\" & $remotecomputer & "\c$\" & $filepathv1_ & "#'+  $datestamp +'" & "#' + $Mdate + '#" & $filemdatev & "'" & @CR & _ ; FileModified,[Y,N],& $filenamev &,Found Value, Model Value
				"   write-host '" & $filenamev & " file has the correct modified date of " & $filemdatev & " ' }" & @CR & _
				"Else " & @CR & _
				"  { " & @CR & _
				"   $fileB = 'FileModified#N#" & $filenamev & "#\\" & $remotecomputer & "\c$\" & $filepathv1_ & "# ' + $Mdate + '#" & $filemdatev & "#'+  $datestamp +'#" & $objectnotes & " '" & @CR & _ ; FileModified,[Y,N],& $filenamev &,Found Value, Model Value
				"   write-host '" & $filenamev & " file has the wrong modified date of' $Mdate'. Expected value = " & $filemdatev & " ' -ForegroundColor red  " & @CR & _
				"    } " & @CR & _
				"Out-file -filepath " & $outfile & " -append -inputObject $fileB  "
		If ($show) Then
			_GUICtrlEdit_AppendText($h_EditGUI, "File Modified Date: " & $filemdatev & @CRLF)
		EndIf
	Else
		$FileMLine = ""
	EndIf

	$FileScriptBase = "" ; initialize scriptbase

	If ($show) Then
		$remotecomputer = "' + $computerName + '" ; for outfile variable
		$remotemachine = "$computerName" ; for set location variable
	EndIf

	If Not ($show) Then
		$remotemachine = $remotecomputer ; for set location variable
		$FileScriptBase = "$ErrorActionPreference = 'silentlycontinue'"  & @CR & _
									"$user = (gwmi win32_computersystem -computername " & $remotemachine & " ).Username" & @CR & _
									"$userfoldername = $user.Remove(0,$user.IndexOf('\') + 1)" & @CRLF
	EndIf

	$FileScriptBase &= "Set-Location \\" & $remotemachine & "\c$\'" & $filepathv1 & "'" & @CR & _ ; "\\c$\'"
			"$file = Get-Childitem | where-Object { $_.name -eq '" & $filenamev & "'} " & @CR & _
			"write-host '      ***************' -ForegroundColor yellow " & @CR & _
			"$datestamp = Get-Date -Format MM-dd-yyyy " & @CR & _  ; outfile date stamp
			"If ($file.Exists) " & @CR & _
			"   { " & @CR & _
			"   $filefound = 'File#Y#" & $filenamev & "#\\" & $remotecomputer & "\c$\" & $filepathv1_ & "#'+  $datestamp +''" & @CR & _  ;File,[Y,N],& $filenamev &,FullFilepath, +  $datestamp
			"   write-host '" & $filenamev & " file Exist'" & @CR & _
			"   $content = get-content '" & $filenamev & "'" & @CR & _
			"   $Mdate = $file.LastWritetime.ToShortDateString() " & @CR & _
			"   $CDate = $file.CreationTime.ToShortDateString()" & @CR & _
			"   $filelength = $file.Length" & @CR & _
			"" & $FileSizeLine & @CR & _
			"" & $FileSearchLine & @CR & _
			"" & $FileCLine & @CR & _
			"" & $FileMLine & @CR & _
			"    }" & @CR & _
			"Else " & @CR & _
			"   {" & @CR & _
			"   $filefound = 'File#N#" & $filenamev & "#\\" & $remotecomputer & "\c$\" & $filepathv1_ & "#'+  $datestamp +'#" & $objectnotes & " '" & @CR & _  ;File,[Y,N],& $filenamev &,FullFilepath, +  $datestamp
			"   write-host '" & $filenamev & " not found ' -ForegroundColor red " & @CR & _
			"   }" & @CR & _
			"write-host '      ***************' -ForegroundColor yellow " & @CR & _
			"Out-file -filepath " & $outfile & " -append -inputObject $filefound " & @CR & _
			"Set-Location c:\ " & @CRLF

	If ($show) Then
		_GUICtrlEdit_AppendText($h_EditGUI, "     ***** " & @CRLF)
	EndIf
;MsgBox(0,"file script", $FileScriptBase)
EndFunc   ;==>FileScript

Func FolderScript($show)

	Local $FolderLocv = StringReplace(GUICtrlRead($FolderLoc), "C:\", "", 0, 2)
	Local $Foldernamev = StringReplace(GUICtrlRead($FolderName), "C:\", "", 0, 2)
	Local $FolderCDatev = GUICtrlRead($FolderCDate)
	Local $FolderMDatev = GUICtrlRead($FolderMDate)
	Local $FolderNumFilesv = GUICtrlRead($FolderNumFiles)
	Local $ShowFolderPath = StringReplace(GUICtrlRead($FolderLoc) & "\" & $Foldernamev, "\\", "\", 0, 2)


	$FolderLocv_ = $FolderLocv ; if no [user] detected - assign them as the same
	$FolderLocvC = $FolderLocv  ; if no [user] detected - assign them as the same
	$Foldernamev_ = $Foldernamev
	$Foldernamev = "'"& $Foldernamev &"'"

		If 	StringInStr($FolderLocv,"[user]", 0) Then
			;MsgBox(0, "before string replace", $FolderLocv) ; if folder location is userprofile
			;$FolderLocv_ = StringReplace($FolderLocv, "[user]", "$userfoldername") ; syntax for ps comments old syntax
			$FolderLocvC = StringReplace($FolderLocv, "[user]","' + $userfoldername + '") ; syntax for ps comments new syntax
			$FolderLocv = StringReplace($FolderLocv, "[user]","'$userfoldername'") ; syntax for ps commands
		Endif

	If $FolderLocv = "" Then ;  if folder location is root and folderloc variable is empty run this
		$FolderLocvCount = ""
	Else
		$FolderLocvCount = $FolderLocv & "\"
	EndIf

;MsgBox(0," before folder loc", $FolderLocvCount & $Foldernamev)
$FullFolderPath = StringReplace($FolderLocvCount & $Foldernamev,"'","") ; used for counting files and folders
;MsgBox(0," after folder loc", $FullFolderPath)


	If  StringInStr($Foldernamev,"[user]", 0) Then
	    $Foldernamev_ = StringReplace($Foldernamev, "[user]","+ $userfoldername +")
		$Foldernamev = StringReplace($Foldernamev, "'[user]'","$userfoldername")
	Endif

;MsgBox(0, "beging function",  $Foldernamev )
;MsgBox(0,"Foldername value is ", $Foldernamev)
;MsgBox(0,"other Foldername value is ", $Foldernamev)

	If ($show) Then
		_GUICtrlEdit_AppendText($h_EditGUI, "Folder Object to test: " & $ShowFolderPath & @CRLF)
	EndIf

	If _IsChecked($FolderCDateBox) Then
		$FolderCDateLine = " if ($folderCDate_ -eq '" & $FolderCDatev & "' ) " & @CR & _
				"   { " & @CR & _
				"   $folderB =  'FolderCreate#Y#" & $Foldernamev_ & "#\\" & $remotecomputer & "\c$\" & $FolderLocvC & "#'+  $datestamp +'#' + $folderCDate_ +'#" & $FolderCDatev & "'" & @CR & _ ;FolderCreate,[Y,N],FolderName,Found Value, Model Value datestamp
				'   write-host "' & $Foldernamev & ' folder has the expected create date: ' & $FolderCDatev & '"' & @CR & _
				"    } " & @CR & _
				"Else " & @CR & _
				"   { " & @CR & _
				"   $folderB =  'FolderCreate#N#" & $Foldernamev_ & "#\\" & $remotecomputer & "\c$\" & $FolderLocvC & "#' + $folderCDate_ +'#" & $FolderCDatev & "#'+  $datestamp +'#" & $objectnotes & " '" & @CR & _ ;FolderCreate,[Y,N],FolderName,Found Value, Model Value datestamp
				'   write-host "' & $Foldernamev & ' folder has the wrong create date of $folderCDate_. Expected Value: ' & $FolderCDatev & '" -ForegroundColor red ' & @CR & _
				"    } " & @CR & _  ;;;;;;;;;
				" Out-file -filepath " & $outfile & " -append -inputObject $folderB "

		If ($show) Then
			_GUICtrlEdit_AppendText($h_EditGUI, "Folder Create Date: " & $FolderCDatev & @CRLF)
		EndIf
	Else
		$FolderCDateLine = ""
	EndIf

	If _IsChecked($FolderMDateBox) Then
		$FolderMDateLine = "if ($folderMDate_ -eq '" & $FolderMDatev & "' ) " & @CR & _
				"  { " & @CR & _
				"   $folderA =  'FolderModified#Y#" & $Foldernamev_ & "#\\" & $remotecomputer & "\c$\" & $FolderLocvC & "#'+  $datestamp +'#' + $folderMDate_ +'#" & $FolderMDatev & "'" & @CR & _ ;FolderModified,[Y,N],FolderName,Found Value, Model Value datestamp
				'   write-host "' & $Foldernamev & ' folder has the expected modified date: ' & $FolderMDatev & ' " ' & @CR & _ ; single qoutes
				"    } " & @CR & _
				"Else " & @CR & _
				"  { " & @CR & _
				"   $folderA =  'FolderModified#N#" & $Foldernamev_ & "#\\" & $remotecomputer & "\c$\" & $FolderLocvC & "#' + $folderMDate_ +'#" & $FolderMDatev & "#'+  $datestamp +'#" & $objectnotes & " '" & @CR & _ ;FolderModified,[Y,N],FolderName,Found Value, Model Value datestamp
				'   write-host "' & $Foldernamev & ' folder has the wrong modified date of $folderMDate_. Expected value: ' & $FolderMDatev & '" -ForegroundColor red  ' & @CR & _ ; single qoutes
				"    } " & @CR & _
				" Out-file -filepath " & $outfile & " -append -inputObject $folderA "

		If ($show) Then
			_GUICtrlEdit_AppendText($h_EditGUI, "Folder Modified Date: " & $FolderMDatev & @CRLF)
		EndIf
	Else
		$FolderMDateLine = ""
	EndIf

	If _IsChecked($NumFilesBox) Then
		$FolderFilesLine = " if ($folderfileCount  -eq " & $FolderNumFilesv & " ) " & @CR & _
				"    { " & @CR & _
				"   $folderC = 'FolderNumber#Y#" & $Foldernamev_ & "#\\" & $remotecomputer & "\c$\" & $FolderLocvC & "#'+  $datestamp +' #' + $folderfileCount +'#" & $FolderNumFilesv & "'" & @CR & _ ;FolderNumber,[Y,N],FolderName,Found Value, Model Value"# ' +  $datestamp
				'   write-host "'& $Foldernamev & ' folder has the  correct number of files:  ' & $FolderNumFilesv & ' " ' & @CR & _ ; single qoutes
				"    } " & @CR & _
				"Else " & @CR & _
				"  { " & @CR & _
				"   $folderC = 'FolderNumber#N#" & $Foldernamev_ & "#\\" & $remotecomputer & "\c$\" & $FolderLocvC & "#' + $folderfileCount +'#" & $FolderNumFilesv & "#'+  $datestamp +'#" & $objectnotes & " '" & @CR & _ ;FolderNumber,[Y,N],FolderName,Found Value, Model Value"# ' +  $datestamp
				'   write-host "' & $Foldernamev & ' folder has incorrect number of files/folders of $folderfileCount. Expected number of files: ' & $FolderNumFilesv & ' " -ForegroundColor red ' & @CR & _ ; single qoutes
				"    } " & @CR & _
				" Out-file -filepath " & $outfile & " -append -inputObject $folderC "
		If ($show) Then
			_GUICtrlEdit_AppendText($h_EditGUI, "Number Files in the Folder: " & $FolderNumFilesv & @CRLF)
		EndIf
	Else
		$FolderFilesLine = ""
	EndIf

	$FolderScriptBase = "" ; initialize scriptbase

	If ($show) Then
		$remotecomputer = "' + $computerName + '" ; for outfile variable
		$remotemachine = "$computerName" ; for set location variable
	EndIf

	If Not ($show) Then
		$remotemachine = $remotecomputer ; for set location variable
		$FolderScriptBase = "$ErrorActionPreference = 'silentlycontinue'" & @CR & _
			"$user = (gwmi win32_computersystem -computername " & $remotemachine & ").Username" & @CR & _
			"$userfoldername = $user.Remove(0,$user.IndexOf('\') + 1)" & @CRLF
	EndIf
    ;MsgBox(0, "namev",  $Foldernamev )
	;MsgBox(0, "locv_", $FolderLocv_ )
	$FolderScriptBase &= "Set-Location \\" & $remotemachine & "\c$\'" & $FolderLocv & "'" & @CR & _
			"$folder = Get-Childitem -directory | where-object {$_.name -eq " & $Foldernamev & "}" & @CR & _
			"write-host '      ***************' -ForegroundColor yellow " & @CR & _
			"$datestamp = Get-Date -Format MM-dd-yyyy " & @CR & _  ; outfile date stamp
			"if ($folder.Exists) " & @CR & _
			"   { " & @CR & _
			"   $folderMDate_ = $folder.LastWriteTime.ToShortDateString() " & @CR & _
			"   $folderCDate_ = $folder.CreationTime.ToShortDateString() "  & @CR & _           ;;;;;;;;;;;;;;
			"   $folderfileCount = (Get-ChildItem -force  \\" & $remotemachine & "\c$\'" & $FullFolderPath & "' | Measure-object).Count " & @CR & _ ; good
			"   $FolderResult = 'Folder#Y#"& $Foldernamev_ &"#\\" & $remotecomputer & "\c$\" & $FolderLocvC & "#'+ $datestamp +'' " & @CR & _  ;Folder,[Y,N],FolderName,FullFolderpath datestamp
			'   write-host  "'& $Foldernamev & ' folder EXIST"' & @CR & _
			"" & $FolderCDateLine & @CR & _
			"" & $FolderMDateLine & @CR & _
			"" & $FolderFilesLine & @CR & _
			"   } " & @CR & _
			"Else " & @CR & _
			"   {" & @CR & _
			"   $FolderResult = 'Folder#N#" & $Foldernamev_ & "#\\" & $remotecomputer & "\c$\" & $FolderLocvC & "#'+  $datestamp +'#" & $objectnotes & " '" & @CR & _ ;Folder,[Y,N],FolderName,FullFolderpath datestamp
			'   write-host "' & $Foldernamev & ' folder not found "  -ForegroundColor red ' & @CR & _
			"   }" & @CR & _
			"write-host '      ***************' -ForegroundColor yellow " & @CR & _
			"Out-file -filepath " & $outfile & " -append -inputObject  $FolderResult  " & @CR & _
			"Set-Location c:\ " & @CRLF

;	"   $FolderResult = 'Folder#Y#" & $Foldernamev_ & "#\\" & $remotecomputer & "\c$\" & $FolderLocv_ & "#'+  $datestamp +''" & @CR & _  ;Folder,[Y,N],FolderName,FullFolderpath datestamp


	If ($show) Then
		_GUICtrlEdit_AppendText($h_EditGUI, "     ***** " & @CRLF)
	EndIf
	;msg("end folderscript func")

EndFunc   ;==>FolderScript

Func RegistryScript($show)

	Local $RegHivev = GUICtrlRead($RegHive)
	Local $RegKeyv = GUICtrlRead($RegKey)
	Local $RegValuev = StringReplace(GUICtrlRead($RegValue), "\", "", 0, 2)
	Local $RegDatav = GUICtrlRead($RegData)
	Local $HiveValue

	If $RegHivev = "LocalMachine" Then
		$HiveValue = "HKLM"
	Else
		$HiveValue = "HKCU"
	EndIf

	$RegistryScriptBase = "" ; initialize scriptbase

	If ($show) Then ; if adding script on reg menu
		$remotecomputer = "$computerName" ;
		$RegistryScriptBase = "write-host '      ***************' -ForegroundColor yellow " & @CR & _
				"$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey( '" & $RegHivev & "', $computerName  ) " & @CRLF
	EndIf

	If Not ($show) Then ; if testing script on reg menu
		$RegistryScriptBase = " $ErrorActionPreference = 'silentlycontinue'" & @CR & _
				"write-host '      ***************' -ForegroundColor yellow " & @CR & _
				"$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey( '" & $RegHivev & "','" & $remotecomputer & "' ) " & @CRLF
	EndIf

	; check remoteregistry get-service
	$RegistryScriptBase &= "$datestamp = Get-Date -Format MM-dd-yyyy " & @CR & _  ; outfile date stamp
			"$RegKey = $Reg.OpenSubkey('" & $RegKeyv & "' ) " & @CR & _
			"if ($Regkey ) #regkey works  " & @CR & _ ; open subkey found, subkey = $Regkey
			"        {  " & @CR & _
			"                $RegKeyResult = 'RegSubKey#Y#' + $RegKey +'#" & $RegHivev & "\" & $RegKeyv & "#'+  $datestamp +''" & @CR & _ ;RegSubKey,[Y,N],RegKey,Model RegKey datestamp
			"             write-host  '" & $RegHivev & "\" & $RegKeyv & " subkey found ' " & @CR & _
			"                $Regdata = $RegKey.GetValue( '" & $RegValuev & "' ) " & @CR & _
			"                if ($Regdata) #regdata accessed " & @CR & _
			"                           {  " & @CR & _
			"                           $RegSubKeyValueResult = 'RegSubKeyValue#Y#" & $RegValuev & "#" & $RegHivev & "\" & $RegKeyv & " " & $RegValuev & "#'+  $datestamp +''" & @CR & _ ;RegSubKeyValue,[Y,N],RegSubkeyValue,Model RegSubKeyValue datestamp
			"                           write-host  '" & $RegHivev & "\" & $RegKeyv & " >> " & $RegValuev & " subkey value  found  '" & @CR & _
			"                            if ($Regdata.ToString().Equals('" & $RegDatav & "')) # RegData value  is correct  " & @CR & _ ; Registry Data Correct  ; data = $RegData
			"                                      { " & @CR & _
			"	                              	     $RegDataResult = 'RegData#Y#' + $Regdata +'#" & $RegHivev & "\" & $RegKeyv & " " & $RegValuev & " -> " & $RegDatav & "#'+  $datestamp +''" & @CR & _ ;RegData,[Y,N],RegData,Model Value datestamp
			"		                            	 write-host '" & $RegHivev & "\" & $RegKeyv & " " & $RegValuev & " contains " & $RegDatav & " data.  '" & @CR & _
			"                                       } " & @CR & _
			"                               Else    # Regdata value is incorrect " & @CR & _  ; Registry Data Incorrect  ; data = $RegData
			"                                        { " & @CR & _
			"	                              	     $RegDataResult = 'RegData#N#' + $Regdata +'#" & $RegHivev & "\" & $RegKeyv & " " & $RegValuev & " -> " & $RegDatav & "#'+  $datestamp  +'#" & $objectnotes & " '" & @CR & _ ;RegData,[Y,N],RegData,Model Value datestamp
			"                                        write-host  '" & $RegHivev & "\" & $RegKeyv & " >> " & $RegValuev & " with data : " & $RegDatav & " not found '  -ForegroundColor red " & @CR & _
			"                                         } " & @CR & _
			"                              }  " & @CR & _
			"                  Else #regdata not accessed" & @CR & _
			"                         { " & @CR & _
			"                           $RegSubKeyValueResult = 'RegSubKeyValue#N#" & $RegValuev & "#" & $RegHivev & "\" & $RegKeyv & " " & $RegValuev & "#'+  $datestamp  +'#" & $objectnotes & " '" & @CR & _ ;RegSubKeyValue,[Y,N],RegSubkeyValue,Model RegSubKeyValue datestamp
			"                         write-host  '" & $RegHivev & "\" & $RegKeyv & " >> " & $RegValuev & " subkey value  not found  '  -ForegroundColor red " & @CR & _
			"                         } " & @CR & _
			"          } " & @CR & _
			"Else #regkey does not work" & @CR & _ ; Open subkey not found subkey = $Regkey
			"         { " & @CR & _
			"            $RegKeyResult = 'RegSubKey#N#Not accessible#" & $RegHivev & "\" & $RegKeyv & "#'+  $datestamp  +'#" & $objectnotes & " '" & @CR & _ ;RegsubKey,[Y,N],RegKey,Model RegKey datestamp
			"             write-host  '" & $RegHivev & "\" & $RegKeyv & " subkey not found '  -ForegroundColor red " & @CR & _
			"          } " & @CR & _
			"write-host '      ***************' -ForegroundColor yellow " & @CR & _
			"Out-File -filepath " & $outfile & " -Append -InputObject  $RegKeyResult,$RegSubKeyValueResult,$RegDataResult " & @CRLF



	If ($show) Then
		_GUICtrlEdit_AppendText($h_EditGUI, "Registry value to test: " & @CRLF)
		_GUICtrlEdit_AppendText($h_EditGUI, $HiveValue & "\" & $RegKeyv & "\" & $RegValuev & " containing " & $RegDatav & @CRLF)
		_GUICtrlEdit_AppendText($h_EditGUI, "     ***** " & @CRLF)
	EndIf
EndFunc   ;==>RegistryScript

Func ProcessScript($show)

	Local $ProcessNamev = StringReplace(GUICtrlRead($ProcessName), ".exe", "") ; delete process name with exe at the end

	$ProcessScriptBase = "" ; initialize scriptbase

	If ($show) Then
		$remotecomputer = "$computerName" ;
	EndIf

	If Not ($show) Then
		$ProcessScriptBase = "$ErrorActionPreference = 'silentlycontinue'" & @CRLF
	EndIf


	$ProcessScriptBase &= "write-host '      ***************' -ForegroundColor yellow " & @CR & _
			"$processstatus = Get-Process -name '" & $ProcessNamev & "' -computername " & $remotecomputer & @CR & _
			"$datestamp = Get-Date -Format MM-dd-yyyy " & @CR & _  ; outfile date stamp
			" if ($processstatus.ProcessName -eq '" & $ProcessNamev & "')" & @CR & _
			"   { " & @CR & _
			"   $processResult = 'Process#Y#" & $ProcessNamev & "#'+ $computerName +'#'+  $datestamp +''" & @CR & _ ;Process, [Y,N],ProcessName,$remotecomputer, datestamp
			"   write-host  '" & $ProcessNamev & " process found'" & @CR & _
			"   } " & @CR & _
			"Else " & @CR & _
			"   { " & @CR & _
			"   $processResult = 'Process#N#" & $ProcessNamev & "#'+ $computerName +'#'+  $datestamp  +'#" & $objectnotes & " '" & @CR & _  ;Process, [Y,N],ProcessName,$remotecomputer, datestamp
			"   write-host  '" & $ProcessNamev & " process not  found'   -ForegroundColor red " & @CR & _
			"   } " & @CR & _
			"write-host '      ***************' -ForegroundColor yellow " & @CR & _
			"Out-file -filepath " & $outfile & " -append -inputObject $processResult " & @CRLF

	If ($show) Then
		_GUICtrlEdit_AppendText($h_EditGUI, "Process name to check: " & $ProcessNamev & @CRLF)
		_GUICtrlEdit_AppendText($h_EditGUI, "     ***** " & @CRLF)
	EndIf
EndFunc   ;==>ProcessScript

Func ServiceScript($show)

	Local $ServiceNamev = GUICtrlRead($ServiceName)
	Local $ServiceStatev = GUICtrlRead($ServiceState)

	$ServiceScriptBase = "" ; initialize scriptbase

	If ($show) Then
		$remotecomputer = "$computerName" ;
	EndIf

	If Not ($show) Then
		$ServiceScriptBase = "$ErrorActionPreference = 'silentlycontinue'" & @CRLF
	EndIf


	If $ServiceStatev = "Running" Then
		;msg("running")
		$ServiceScriptBase &= " $servicestatus = Get-Service -name '" & $ServiceNamev & "' -computername " & $remotecomputer & @CR & _
				"write-host '      ***************' -ForegroundColor yellow " & @CR & _
				"$datestamp = Get-Date -Format MM-dd-yyyy " & @CR & _  ; outfile date stamp
				"if ( $servicestatus.Status -eq 'Running') " & @CR & _
				"   { " & @CR & _
				"   $serviceResult = 'Service#Y#" & $ServiceNamev & "#running#' + $computerName +'#'+  $datestamp +''" & @CR & _;ServiceName,Found,[Y,N],Found Valuetime stamp  +  $datestamp " & @CR & _
				"   write-host '" & $ServiceNamev & " is Running' " & @CR & _
				"   } " & @CR & _
				"Elseif ( $servicestatus.Status -eq 'Stopped') " & @CR & _
				"   { " & @CR & _
				"   $serviceResult = 'Service#N#" & $ServiceNamev & "#stopped#' + $computerName +'#'+  $datestamp +'#" & $objectnotes & " '" & @CR & _ ;ServiceName,Found,[Y,N],Found Valuetime stamp  +  $datestamp " & @CR & _
				"   write-host '" & $ServiceNamev & " process is stopped ' -ForegroundColor red " & @CR & _
				"   } " & @CR & _
				"Else " & @CR & _
				"   { " & @CR & _
				"   $serviceResult = 'Service#N#" & $ServiceNamev & "#not found#' + $computerName +'#'+  $datestamp +'#" & $objectnotes & " '" & @CR & _ ;ServiceName,Found,[Y,N],Found Valuetime stamp  +  $datestamp " & @CR & _
				"   write-host '" & $ServiceNamev & " process not found'    -ForegroundColor red " & @CR & _
				"   } " & @CR & _
				"write-host '      ***************' -ForegroundColor yellow " & @CR & _
				"Out-file -filepath " & $outfile & " -append -inputObject $serviceResult  " & @CRLF

	Else
		;msg("Stopped")
		$ServiceScriptBase &= " $servicestatus = Get-Service -name '" & $ServiceNamev & "' -computername " & $remotecomputer & @CR & _
				"write-host '      ***************' -ForegroundColor yellow " & @CR & _
				"$datestamp = Get-Date -Format MM-dd-yyyy " & @CR & _  ; outfile date stamp
				"if ( $servicestatus.Status -eq 'Stopped') " & @CR & _
				"   { " & @CR & _
				"   $serviceResult = 'Service#Y#" & $ServiceNamev & "#stopped#' + $computerName +'#'+  $datestamp +''" & @CR & _;ServiceName,Found,[Y,N],Found Valuetime stamp  +  $datestamp " & @CR & _
				"   write-host '" & $ServiceNamev & " is Stopped' " & @CR & _
				"   } " & @CR & _
				"Elseif ( $servicestatus.Status -eq 'Running') " & @CR & _
				"   { " & @CR & _
				"   $serviceResult = 'Service#N#" & $ServiceNamev & "#running#' + $computerName +'#'+  $datestamp +'#" & $objectnotes & " '" & @CR & _ ;ServiceName,Found,[Y,N],Found Valuetime stamp  +  $datestamp " & @CR & _
				"   write-host '" & $ServiceNamev & " process is running ' -ForegroundColor red " & @CR & _
				"   } " & @CR & _
				"Else " & @CR & _
				"   { " & @CR & _
				"   $serviceResult = 'Service#N#" & $ServiceNamev & "#not found#' + $computerName +'#'+  $datestamp +'#" & $objectnotes & " '" & @CR & _ ;ServiceName,Found,[Y,N],Found Valuetime stamp  +  $datestamp " & @CR & _
				"   write-host '" & $ServiceNamev & " process not found'    -ForegroundColor red " & @CR & _
				"   } " & @CR & _
				"write-host '      ***************' -ForegroundColor yellow " & @CR & _
				"Out-file -filepath " & $outfile & " -append -inputObject $serviceResult  " & @CRLF
	Endif





	If ($show) Then
		_GUICtrlEdit_AppendText($h_EditGUI, "Service name to check: " & $ServiceNamev &  " should be " & $ServiceStatev & @CRLF)
		_GUICtrlEdit_AppendText($h_EditGUI, "     ***** " & @CRLF)
	EndIf

EndFunc   ;==>ServiceScript

Func OnlineScript($show)

	Local $OnlineNamev = GUICtrlRead($OnlineName)

	$OnlineScriptBase = "" ; initialize scriptbase

	If Not ($show) Then
		$OnlineScriptBase = " $ErrorActionPreference = 'silentlycontinue'" & @CRLF
	EndIf

	$OnlineScriptBase &= "write-host '      ***************' -ForegroundColor yellow " & @CR & _
			"$datestamp = Get-Date -Format MM-dd-yyyy " & @CR & _  ; outfile date stamp
			"$HTTP_Response = $null " & @CR & _
			"$HTTP_Request = [System.Net.WebRequest]::Create('" & $OnlineNamev & "') " & @CR & _
			"try { " & @CR & _
			"	$HTTP_Response = $HTTP_Request.GetResponse() " & @CR & _
			"	$HTTP_Status = [int]$HTTP_Response.StatusCode " & @CR & _
			"	If ($HTTP_Status -eq 200) { " & @CR & _
			"   $OnlineResult = 'Online#Y#" & $OnlineNamev & "#'+  $datestamp +''" & @CR & _    ;Online,[Y,N],OnlineMachineName, datestamp
			"   write-host '" & $OnlineNamev & " is Online' " & @CR & _
			"		} " & @CR & _
			"	else{ " & @CR & _
			"   $OnlineResult = 'Online#N#" & $OnlineNamev & "#'+  $datestamp +'#" & $objectnotes & " '" & @CR & _    ;Online,[Y,N],OnlineMachineName, datestamp
			"   write-host '" & $OnlineNamev & " is not online'  -ForegroundColor red " & @CR & _
			"		} " & @CR & _
			"	$HTTP_Response.Close() " & @CR & _
			"		} " & @CR & _
			"catch{ " & @CR & _
			"	$HTTP_Status = [regex]::matches($_.exception.message, '(?<=\()[\d]{3}').Value " & @CR & _
			"   $OnlineResult = 'Online#N#" & $OnlineNamev & "#'+  $datestamp +'#" & $objectnotes & " '" & @CR & _    ;Online,[Y,N],OnlineMachineName, datestamp
			"   write-host '" & $OnlineNamev & " is not online'  -ForegroundColor red " & @CR & _
			"	}" & @CR & _
			"write-host '      ***************' -ForegroundColor yellow " & @CR & _
			"Out-file -filepath " & $outfile & " -append -inputObject $OnlineResult" & @CRLF

	If ($show) Then ; if add button clicked
		_GUICtrlEdit_AppendText($h_EditGUI, "Server name to check: " & $OnlineNamev & @CRLF)
		_GUICtrlEdit_AppendText($h_EditGUI, "     ***** " & @CRLF)
	EndIf

EndFunc   ;==>OnlineScript

Func HotfixScript($show)

	Local $HotfixNamev = GUICtrlRead($HotfixName)

	$HotfixScriptBase = "" ; initialize scriptbase
	If ($show) Then
		$remotecomputer = "$computerName" ;
		$HotfixScriptBase &= "write-host '      ***************' -ForegroundColor yellow " & @CR & _
				"$datestamp = Get-Date -Format MM-dd-yyyy " & @CR & _  ; outfile date stamp
				"$hotfixsearch = get-hotfix -id " & $HotfixNamev & " -computername " & $remotecomputer & @CRLF

	EndIf

	If Not ($show) Then
		$HotfixScriptBase = "$ErrorActionPreference = 'silentlycontinue'" & @CR & _
				"write-host '      ***************' -ForegroundColor yellow " & @CR & _
				"$datestamp = Get-Date -Format MM-dd-yyyy " & @CR & _  ; outfile date stamp
				"$hotfixsearch = get-hotfix -id " & $HotfixNamev & " -computername '" & $remotecomputer & "'" & @CRLF
	EndIf


	$HotfixScriptBase &= "if ( $hotfixsearch ) " & @CR & _
			"   { " & @CR & _
			"   $HotFixResult = 'HotFix#Y#" & $HotfixNamev & "#' + $computerName +'#' +  $datestamp +''" & @CR & _ ;HotFix,[Y,N],HotFixName, $remotecomputer, datestamp
			"   write-host   '" & $HotfixNamev & " found in machine' " & @CR & _
			"   } " & @CR & _
			"Else " & @CR & _
			"   { " & @CR & _
			"   $HotFixResult = 'HotFix#N#" & $HotfixNamev & "#' + $computerName +'#' +  $datestamp +'#" & $objectnotes & " '" & @CR & _ ;HotFix,[Y,N],HotFixName, $remotecomputer, datestamp
			"   write-host   '" & $HotfixNamev & " not installed on machine'   -ForegroundColor red " & @CR & _
			"   } " & @CR & _
			"write-host '      ***************' -ForegroundColor yellow " & @CR & _
			"Out-file -filepath " & $outfile & " -append -inputObject $HotFixResult  " & @CRLF

	If ($show) Then ; if add button clicked
		_GUICtrlEdit_AppendText($h_EditGUI, "Hot fix to search: " & $HotfixNamev & @CRLF)
		_GUICtrlEdit_AppendText($h_EditGUI, "     ***** " & @CRLF)
	EndIf

EndFunc   ;==>HotfixScript

#EndRegion Object Scripts

#Region Other functions

Func _IsChecked($idControlID)
	Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc   ;==>_IsChecked

Func UniqueFilename()

	Local $sFilePath, $filenamev
	$ProfileTypev = GUICtrlRead($ProfileType)
	$ScriptFileNametemp = GUICtrlRead($ScriptFileName)
	$ScriptFileNamev = StringReplace($ScriptFileNametemp, ".ps1", "")
	If $ScriptFileNametemp = "" Or $ScriptFileNametemp = " " Then
		MsgBox(16, "FiDi", "Filename Error", 2)
		Return False
	EndIf

	If $ProfileTypev = "Application Profile" Then ; assign profile extension with ps1
		$sFilePath = $AppFolder & "\" & $ScriptFileNamev & ".ps1"
	ElseIf $ProfileTypev = "Task Profile" Then
		$sFilePath = $TaskFolder & "\" & $ScriptFileNamev & ".ps1"
	ElseIf $ProfileTypev = "Symptom Profile" Then
		$sFilePath = $SympFolder & "\" & $ScriptFileNamev & ".ps1"
	Else
		MsgBox(16, "Profile Script Builder", "Filename Error", 2)
		Return False
	EndIf

	If FileExists($sFilePath) Then
		MsgBox(0, "Profile Script Builder", "Filename already exist.", 2)
		Return False
	Else
		$UniqueFileNameScript = $sFilePath
		Return True
	EndIf
EndFunc   ;==>UniqueFilename

Func SavetoFile($ScriptBase)
	If Not FileExists($UniqueFileNameScript) Then _FileCreate($UniqueFileNameScript)

	Local $hFileOpen = FileOpen($UniqueFileNameScript, $FO_OVERWRITE)
	If $hFileOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "", "Error occured saving the script.", 2)
		Return False
	EndIf
	$TempScript = ""
	$TempScript = "Param( [parameter(Mandatory=$true)] " & @CR & _
			"[String] $computerName ) " & @CR & _
			"$ErrorActionPreference = 'silentlycontinue'" & @CR & _
			"$user = (gwmi win32_computersystem -computername $computerName).Username" & @CR & _
			"$userfoldername = $user.Remove(0,$user.IndexOf('\') + 1)"  & @CRLF

	$TempScript &= $ScriptBase

	FileWrite($hFileOpen, $TempScript) ;
	FileClose($hFileOpen)
	$RunningScript = "" ; clear running list
	_GUICtrlEdit_SetText($h_EditGUI, "") ; clear edit box
	MsgBox(0, "Profile Script Builder", "Profile Saved", 2)
EndFunc   ;==>SavetoFile

Func SaveRunTempFile()
	Local $TempfilePs1 =  @ScriptDir & "\Util\" & "tempfile.ps1"
	If Not FileExists($TempfilePs1) Then
		_FileCreate($TempfilePs1)
	Endif
	Local $hFileOpen = FileOpen($TempfilePs1, $FO_OVERWRITE)
	If $hFileOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "", "Error processing Scripts.", 2)
		Return False
	EndIf
	$TempScript = ""
	$TempScript = "Param( [parameter(Mandatory=$true)] " & @CR & _
			"[String] $computerName ) " & @CR & _
			"$ErrorActionPreference = 'silentlycontinue'" & @CR & _
			"$user = (gwmi win32_computersystem -computername $computerName).Username" & @CR & _
			"$userfoldername = $user.Remove(0,$user.IndexOf('\') + 1)"  & @CRLF

	$TempScript &= $RunningScript
    ;MsgBox(0,"tempscript for user folder", $TempScript)
	FileWrite($hFileOpen, $TempScript) ;
	FileClose($hFileOpen)

EndFunc   ;==>SaveRunTempFile

Func RunPS($script)
	;MsgBox(0, "run ps", $script) ;;;+++++++++++++++++++++++++++++++++++++++++++++delete it to before compiling
	;$compname = InputBox("Service Desk Tool", "Please enter Machine Name or IP address
	Run("powershell.exe -executionpolicy unrestricted -NoExit -NoProfile " & $script, @SystemDir, @SW_MAXIMIZE, 7)
EndFunc   ;==>RunPS

Func RunFilePS()
	SaveRunTempFile() ; check tempfile; run with sample command & filepath computername

	If StringInStr($Tempfile," ") Then
		If $FirstTimeRun Then ; if this function already ran skip the replacement
			$tempfile = StringReplace($Tempfile,":\",":\'")
			;MsgBox(0,"has space",$tempfile) Execute if scriptdir has spaces
		Endif
		Run("powershell.exe -executionpolicy unrestricted -NoExit -NoProfile " & $tempfile & "' " & $remotecomputer , @SystemDir, @SW_MAXIMIZE, 7)
	Else
		;MsgBox(0,"has NO space",$tempfile) Execute if scriptdir has NO spaces
		Run("powershell.exe -executionpolicy unrestricted -NoExit -NoProfile " & $tempfile & " " & $remotecomputer , @SystemDir, @SW_MAXIMIZE, 7)
	Endif
	$FirstTimeRun = 0
EndFunc   ;==>RunFilePS

Func SavetoFileMenu()

	If RunScriptLoaded() Then
		$SavetoFileGUI = GUICreate("Save to File", 445, 160, -1, -1, -1, $WS_EX_TOOLWINDOW, $hGUI)
		GUICtrlCreateLabel("Profile Type", 24, 16, 66, 17)
		$ProfileType = GUICtrlCreateCombo("Application Profile", 104, 16, 121, 25)
		GUICtrlSetData($ProfileType, "Task Profile|Symptom Profile")
		GUICtrlCreateLabel("File Name", 32, 48, 59, 17)
		$ScriptFileName = GUICtrlCreateInput("AppName_Version_Description", 104, 48, 321, 21)
		$SavetoFilebutton = GUICtrlCreateButton("Save Script to File", 160, 88, 115, 33)
		GUICtrlSetState($SavetoFilebutton, $GUI_DEFBUTTON)
		GUISetState(@SW_SHOW, $SavetoFileGUI)
	EndIf
EndFunc   ;==>SavetoFileMenu

Func CheckMachineOnline()
	Local $TestOnlineScriptBase, $online, $status

	Local $compname = "" ; initialize comptername
	Local $tempname = ""
	$tempname = InputBox("Script Builder v0.53", "Enter Machine Name", AutoClip() )
	If ($tempname = "") Or (@error > 0) Then
		Return False
	EndIf

	$compname = StringStripWS($tempname, 8)
	$TestOnlineScriptBase = "$status = Test-Connection -computername " & $compname & " -quiet" & @CR & _
			"if ( $status  -eq $true) {write-host '1'} Else { write-host '0'} "
	$status = Run("powershell.exe -executionpolicy unrestricted -NoProfile  " & $TestOnlineScriptBase, @SystemDir, @SW_HIDE, 3)
	MsgBox(0, "Profile Script Builder", "Checking if machine is online", 1)
	StdinWrite($status, @CRLF)
	StdinWrite($status)
	While 1
		$online &= StdoutRead($status)
		If @error Then ExitLoop
	WEnd

	If $online = 1 Then
		$remotecomputer = $compname
		;MsgBox(0,"online status"," remote comp is: " & $remotecomputer  )
		Return True
	Else
		MsgBox(16, "Profile Script Builder", "Machine not online", 2)
		Return False
	EndIf
EndFunc   ;==>CheckMachineOnline

Func InitializeRunningScript()
	$RunningScript = "Param( [parameter(Mandatory=$true)] " & @CR & _
			"[String] $computerName ) " & @CR & _
			"$ErrorActionPreference = 'silentlycontinue'" & @CR
EndFunc   ;==>InitializeRunningScript

Func CheckSaveFilename()


	If UniqueFilename() Then
		SavetoFile($RunningScript)
		GUIDelete($SavetoFileGUI)
		GUISetState(@SW_ENABLE, $hGUI)
	Else
		GUIDelete($SavetoFileGUI)
		SavetoFileMenu()
	EndIf
EndFunc   ;==>CheckSaveFilename

Func CheckProfileFolder()

	Local $ScriptFolder[3] = [$AppFolder, $TaskFolder, $SympFolder]

	For $i = 0 To 2

		If DirGetSize($ScriptFolder[$i]) = -1 Then
			DirCreate($ScriptFolder[$i])
		EndIf
	Next

EndFunc   ;==>CheckProfileFolder

Func SavetoModel()
	Local $sFileRead, $sFileRead

	If Not FileExists($faultfile) Then ; check for fault.txt
		Return
	EndIf

	If Not FileExists($modelfile) Then ; check model .txt
		_FileCreate($modelfile)
	EndIf

	$MaxFileLine = _FileCountLines($faultfile)

	$hFaultFileOpen = FileOpen($faultfile, 0) ; Read fault.txt on Read Mode
	If $hFaultFileOpen = -1 Then
		Return
	EndIf

	$hModelFileOpen = FileOpen($modelfile, 1) ; Open model.txt in Append Mode
	If $hModelFileOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "FiDi", "Error opening Model.txt File.", 2)
		Return
	EndIf

	For $i = 1 To $MaxFileLine
		$sFileRead = FileReadLine($hFaultFileOpen, $i)
		If StringInStr($sFileRead, "#Y#", 2) Then
			FileWriteLine($hModelFileOpen, $sFileRead) ; assign working objects to model.txt
		EndIf
	Next

	FileClose($hFaultFileOpen)
	FileClose($hModelFileOpen)

EndFunc   ;==>SavetoModel


#EndRegion Other functions

#Region NoteMenu
;:********************* NOTE Menu***********************

Func GetNote($object)
	Local $noteDefault

	;If Not FileExists($notefile) Then ; skip asking for notes if note.txt is missing
	;	Return
	;Endif

	Switch $object
		Case "Online"
			$noteDefault = " Enter contact information or instructions if website is offline."
		Case "File"
			$noteDefault = " Enter instructions if file is a suspect fault."
		Case "Folder"
			$noteDefault = " Enter instructions if folder is a suspect fault."
		Case Else
			$noteDefault = "Enter instructions if object is a suspect fault"
	EndSwitch
	$ShortNoteGUI = GUICreate("Script Builder Notes v0.53", 378, 155, -1, -1,-1, $WS_EX_TOOLWINDOW)
	GUICtrlCreateLabel($noteDefault, 14, 14, 350, 17)
	$NoteText = GUICtrlCreateEdit("", 14, 39, 339, 61, BitOR($ES_WANTRETURN, $ES_MULTILINE), $WS_EX_CLIENTEDGE)
	$LimitLabel = GUICtrlSetLimit($NoteText, 165) ; to limit the entry to 165
	GUICtrlSetData(-1, "165 char limit")
	$NoteButton = GUICtrlCreateButton("Done", 160, 108, 60, 20)
	GUICtrlSetState($NoteButton, $GUI_DEFBUTTON)
	$StringNumber = GUICtrlCreateLabel("165", 332, 108, 47, 17)
	GUISetState(@SW_SHOW, $ShortNoteGUI)


	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				GUIDelete($ShortNoteGUI)
				Return False ; if note was canceled go back to object menu
				ExitLoop
			Case $NoteButton
				$objectnotes = FormatNote()
				GUIDelete($ShortNoteGUI)
				Return True
				ExitLoop
		EndSwitch
		ShowCharLimit()
	WEnd

EndFunc   ;==>GetNote

Func ShowCharLimit()
	Local $char, $StringLimit
	$char = StringLen(GUICtrlRead($NoteText))
	$StringLimit = 165 - $char
	GUICtrlSetData($StringNumber, $StringLimit)
	Sleep(50)
EndFunc   ;==>ShowCharLimit

Func FormatNote()
	$textnote = GUICtrlRead($NoteText)
	Return $textnote
EndFunc   ;==>FormatNote

#EndRegion NoteMenu

Func FileDirMenu($appendMode)
	GUISetState(@SW_DISABLE, $hGUI)
	If RunScriptLoaded() Then
		$FileDirGUI = GUICreate("Script Builder v0.53", 232, 209, -1, -1, -1, $WS_EX_TOOLWINDOW, $hGUI) ; 540 ;-1, $WS_EX_TOOLWINDOW
		$ProfileCombo = GUICtrlCreateCombo("Application Profile", 56, 24, 121, 25, BitOR($CBS_DROPDOWN, $CBS_AUTOHSCROLL))
		GUICtrlSetData(-1, "Task Profile|Symptom Profile")
		$FileList = GUICtrlCreateList("", 24, 56, 183, 96, BitOR($GUI_SS_DEFAULT_LIST, $LBS_EXTENDEDSEL))
		GUICtrlSetData(-1, "")
		_GUICtrlListBox_Dir($FileList, $AppFiles)
		$SelDir = $AppFolder ; initialize selected directory
		$FileDirButton = GUICtrlCreateButton($appendMode, 78, 161, 81, 25)
		GUICtrlSetState($FileDirButton, $GUI_DEFBUTTON)
		GUISetState(@SW_SHOW, $FileDirGUI)
	Else
		;GUISetState(@SW_ENABLE, $hGUI)
		Return
	EndIf

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
					GUISetState(@SW_ENABLE, $hGUI)
					GUIDelete($FileDirGUI)
					ExitLoop
			Case $FileDirButton
					$filenumber = _GUICtrlListBox_GetCurSel($FileList)
					If $filenumber > 0 Then
						If $appendMode = "Append" Then
							AppendScriptToFile()
						ElseIf $appendMode = "Overwrite" Then
							OverwriteScriptToFile()
						EndIf
						GUISetState(@SW_ENABLE, $hGUI)
						GUIDelete($FileDirGUI)
						ExitLoop
					ElseIf $filenumber = 0 then
						MsgBox(0,"Profile Script Builder", "Please select a script")
					Endif
			Case $ProfileCombo
					ShowFileList(GUICtrlRead($ProfileCombo)) ; show sletected profile file list
		EndSwitch
	WEnd

EndFunc   ;==>FileDirMenu

Func ShowFileList($profile)
	$SelDir = ""
	If $profile = "Application Profile" Then
		_GUICtrlListBox_ResetContent($FileList)
		_GUICtrlListBox_Dir($FileList, $AppFiles)
		$SelDir = $AppFolder
	ElseIf $profile = "Task Profile" Then
		_GUICtrlListBox_ResetContent($FileList)
		_GUICtrlListBox_Dir($FileList, $TaskFiles)
		$SelDir = $TaskFolder
	ElseIf $profile = "Symptom Profile" Then
		_GUICtrlListBox_ResetContent($FileList)
		_GUICtrlListBox_Dir($FileList, $SympFiles)
		$SelDir = $SympFolder
	EndIf
EndFunc   ;==>ShowFileList

Func Msg($text)
	MsgBox(0, $text, $text)
EndFunc   ;==>Msg

Func AppendScriptToFile()
	;msg("append button")
	$SelScripts = _GUICtrlListBox_GetText($FileList, $filenumber)
	Selectedprofile($RunningScript, $SelDir & $SelScripts, "Append")

EndFunc   ;==>AppendScriptToFile

Func OverwriteScriptToFile()
	;msg("Overwrite Button")
	$SelScripts = _GUICtrlListBox_GetText($FileList, $filenumber)
	Selectedprofile($RunningScript, $SelDir & $SelScripts, "Overwrite")
EndFunc   ;==>OverwriteScriptToFile

Func Selectedprofile($ScriptBase, $fileSelected, $mode)

	If $mode = "Overwrite" Then
		;msg($mode & " ****  ")
		$filemode = 2 ;$FO_OVERWRITE
		$TempScript = "Param( [parameter(Mandatory=$true)] " & @CR & _
				"[String] $computerName ) " & @CR & _
				"$ErrorActionPreference = 'silentlycontinue'" & @CR & _
				"$user = (gwmi win32_computersystem -computername $computerName).Username" & @CR & _
				"$userfoldername = $user.Remove(0,$user.IndexOf('\') + 1)"  & @CRLF
	ElseIf $mode = "Append" Then
		;msg($mode & "  **** ")
		$filemode = 1 ; $FO_APPEND
		$TempScript = ""
	EndIf

	Local $hFileOpen = FileOpen($fileSelected, $filemode)
	If $hFileOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "", "Error occured writing to script.", 2)
		Return False
	EndIf
	$TempScript &= $ScriptBase
	;msg($TempScript)
	FileWrite($hFileOpen, $TempScript) ;
	FileClose($hFileOpen)
	$RunningScript = " " ; clear running list
	_GUICtrlEdit_SetText($h_EditGUI, "") ; clear edit box
	MsgBox(0, "Script Builder", "Script Saved", 1)

EndFunc   ;==>Selectedprofile

Func CreateWindowMenu() ; menuv
	Local $ManualInfoMenu = GUICtrlCreateMenu("Help")
	Local $idInfoMenu = GUICtrlCreateMenu("About")
	Local $BuyMenu = GUICtrlCreateMenu("Purchase")
	GUICtrlCreateMenuItem("Report bugs to  bugreport@i-ttm.com", $idInfoMenu)
	GUICtrlCreateMenuItem("Trial copy expires : " & _DateTimeFormat($TrialEndDate, 2), $idInfoMenu)
	$ManualLink = GUICtrlCreateMenuItem("Open Script Builder Manual", $ManualInfoMenu)
	$BuyLink = GUICtrlCreateMenuItem("Get Script Builder Software", $BuyMenu)
EndFunc   ;==>CreateWindowMenu

Func RunScriptLoaded() ; check if running script has data
	If $RunningScript = " " Then
		MsgBox(16, "Script Builder v0.53", "No scripts to process", 2)
		GUISetState(@SW_ENABLE, $hGUI)
		Return False
	Else
		Return True
	EndIf
EndFunc

Func AutoClip()
	$MemValue = ClipGet()
	If StringInStr($MemValue, " ") Then
		Return ""
	ElseIf StringInStr($MemValue,"computer") Then
		Return "localhost"
	Else
		Return $MemValue
	EndIf
EndFunc   ;==>AutoClip

Func CheckExpiration()

	If $TrialEndDate >_NowCalc() Then
		Return True
	Else
		DisableButton()
		MsgBox(64,"Script Builder v0.53" , "Trial Date of " &  _DateTimeFormat($TrialEndDate, 2) & " already expired.")
		Return False
	EndIf
EndFunc

Func DisableButton() ; disable file script button
	GUICtrlSetState($SaveScript, $GUI_DISABLE)
	GUICtrlSetState($AppendB, $GUI_DISABLE)
	GUICtrlSetState($OverwriteB, $GUI_DISABLE)
EndFunc


Func EraseTempFile()
	Local $TempfilePs2 = @ScriptDir & "\Util\tempfile.ps1"
	If FileExists($TempfilePs2) Then
		;MsgBox(0,"", "Tempfile exist - deleting")
		FileDelete($TempfilePs2)
	EndIf
EndFunc

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
