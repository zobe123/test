#include <file.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Date.au3>
#Include <FF.au3>
#Include <Chrome.au3>
#Include <Array.au3>
#include <Constants.au3>
#include <Guiconstants.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <ProgressConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>


;### Variablen
$sLongDayName = _DateDayOfWeek(@WDAY)
$stamp = _NowTime()
$stamp2 = _NowDate()
; ConsoleWrite(@UserName & ": " & $sLongDayName & " " & $stamp2 & ", " & $stamp & ": ")
$sFileSource = ""
;$Quelle = ""
$sFileOutput = @WorkingDir & "\output_" & StringReplace($stamp2, ".", "") & StringReplace($stamp, ":", "") & ".txt"
$IP = "127.0.0.1"
$iPort = 4242
$iTimeOut = 3000
Local $iPID = 0
;$sComboRead = "Filme 1080p"
;Global $ForumID = "forumchoice[]28"
;Global $ForumID = "forumchoice[]=61&forumchoice[]=68&forumchoice[]=71&forumchoice[]=145&forumchoice[]=80&forumchoice[]=101&forumchoice[]=168&forumchoice[]=96&forumchoice[]=103&forumchoice[]=99&forumchoice[]=175&forumchoice[]=176&forumchoice[]=177&forumchoice[]=181"
Global $ForumID = "&forumchoicer=1&forumchoice%5B%5D=61&forumchoice%5B%5D=68&forumchoice%5B%5D=71&forumchoice%5B%5D=145&forumchoice%5B%5D=80&forumchoice%5B%5D=101&forumchoice%5B%5D=168&forumchoice%5B%5D=96&forumchoice%5B%5D=103&forumchoice%5B%5D=99&forumchoice%5B%5D=175&forumchoice%5B%5D=176&forumchoice%5B%5D=177&forumchoice%5B%5D=181&showposts=0"
Global $noResult = 1
;Global $prelink="https://usenet-4all.pw/forum/search.php?do=process&" & $ForumID & "&titleonly=0&query="
Global $prelink="https://house-of-usenet.com/search.php?action=do_search&postthread=1&forums=all&showresults=threads&keywords=" ;& $ForumID & "&titleonly=0&query="
$AnzahlTitelText = ""
Local $i = 0
Local $gui_perc = 0
Global $text[0]
; Declare the flags
$Interrupt = 0
$EventCheck = 0



Opt("GUIOnEventMode", 1)
;GUI()

;Func GUI()
$Form1 = GUICreate("HoU batch-Search by ZOBe123", 381, 281, 192, 114)
GUISetIcon("C:\Windows\System32\SyncCenter.dll", -50)
GUISetOnEvent($GUI_EVENT_CLOSE, "ThatExit")

$Input = GUICtrlCreateGroup("Input", 10, 8, 361, 110)
$Quelle = GUICtrlCreateLabel("Quelldatei: " & $sFileSource, 15, 24, 275, 25)
$open_Button_input = GUICtrlCreateButton("öffnen", 305, 24, 57, 25)
GUICtrlSetOnEvent($open_Button_input, "OpenFile")
$Forum = GUICtrlCreateCombo("Filme 1080p", 15, 56, 193, 25, BitOR($CBS_DROPDOWNLIST,$CBS_AUTOHSCROLL))
GUICtrlSetData($Forum, "Filme 720p|3D HDTV 1080/720 Deutsch|Serien 1080p|Serien 720p", "Filme 1080p")
GUICtrlSetState($Forum, $GUI_DISABLE)
$AnzahlTitelText = GUICtrlCreateLabel("", 15, 95, 275, 20)

$Output = GUICtrlCreateGroup("Output", 10, 128, 361, 110)
;GUICtrlSetState($Output, $GUI_DISABLE)
$Checkbox = GUICtrlCreateCheckbox("'kein Ergebnis' auch protokollieren", 16, 180, 345, 25)
GUICtrlSetState($Checkbox, $GUI_DISABLE)
$Checkbox2 = GUICtrlCreateCheckbox("Suche auch nach:", 16, 205, 145, 25)
GUICtrlSetState($Checkbox2, $GUI_DISABLE)
Global $InputBox = GUICtrlCreateInput("", 160, 205, 200, 20)
GUICtrlSetState($InputBox, $GUI_DISABLE)
$Ziel = GUICtrlCreateLabel("Zieldatei: " & $sFileOutput, 15, 144, 275, 30)
$open_Button_output = GUICtrlCreateButton("öffnen", 305, 144, 57, 25)
GUICtrlSetOnEvent($open_Button_output, "OpenOutputFile")
;GUICtrlSetState($open_Button_output, $GUI_DISABLE)
;GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $gui_progressbar = GUICtrlCreateProgress(10, 248, 245, 25)
Global $gui_progressbarLabel = GUICtrlCreateLabel("", 10, 253, 245, 25, 0x01)
Guictrlsetbkcolor($gui_progressbarLabel, $GUI_BKCOLOR_TRANSPARENT)
$StartButton = GUICtrlCreateButton("Start", 265, 248, 105, 25)
GUICtrlSetOnEvent($StartButton, "CreateOutput")
GUICtrlSetState($StartButton, $GUI_DISABLE)
Global $StopButton = GUICtrlCreateButton("Stop", 265, 248, 105, 25)
GUICtrlSetOnEvent($StopButton, "StopFunc")
GUICtrlSetState($StopButton, $GUI_HIDE)
;$Label1 = GUICtrlCreateLabel("(c) 2017 by ZOBe123", 264, 256, 106, 17)

GUISetState(@SW_SHOW)
GUIRegisterMsg($WM_COMMAND, "_WM_COMMAND")

While 1
  Sleep(10)
  If $EventCheck = 1 Then
    ; This temporary return to the main loop, allows AutoIt to quickly handle system events
    ; such as GUI_EVENT_CLOSE. If we reached this far, then its safe to assume that
    ; there was no system events, and we can return to the RunnerFunc.
    ;GUI()
  EndIf
WEnd

;Func GUI()

;While 1
;	$nMsg = GUIGetMsg()
;	Switch $nMsg
;		 Case $GUI_EVENT_CLOSE
;			ExitLoop
;			Exit
;		 Case - 3
;			ExitLoop
;			Exit
;		 Case $open_Button_input
;			$sFileSource = OpenFile()
;			ConsoleWrite($sFileSource)
;			If $sFileSource <> "" Then
;			   GUICtrlSetState($Forum, $GUI_ENABLE)
;			   GUICtrlSetState($Output, $GUI_ENABLE)
;			   GUICtrlSetData($Quelle,"Quelldatei: " & $sFileSource)
;			   GUICtrlSetState($Checkbox, $GUI_ENABLE)
;			   GUICtrlSetState($Checkbox2, $GUI_ENABLE)
;			   GUICtrlSetState($StartButton, $GUI_ENABLE)
;			Else
;			   GUICtrlSetState($Forum, $GUI_DISABLE)
;			   GUICtrlSetState($Output, $GUI_DISABLE)
;			   GUICtrlSetState($Checkbox, $GUI_DISABLE)
;			   GUICtrlSetState($Checkbox2, $GUI_DISABLE)
;			   GUICtrlSetState($StartButton, $GUI_DISABLE)
;			EndIf
;		 Case $open_Button_output
;			$iPID = Run("notepad.exe", $sFileOutput, @SW_SHOWMAXIMIZED)
;		 Case $Checkbox
;			If _IsChecked($Checkbox) Then
;			   ConsoleWrite("Checkbox checked " & @CRLF)
;			   $noResult = 1
;			Else
;			   ConsoleWrite("Checkbox unchecked " & @CRLF)
;			   $noResult = 0
;			EndIf
;		 Case $Checkbox2
;			If _IsChecked($Checkbox2) Then
;			   ConsoleWrite("Checkbox2 checked " & @CRLF)
;			   GUICtrlSetState($InputBox, $GUI_ENABLE)
;			Else
;			   ConsoleWrite("Checkbox2 unchecked " & @CRLF)
;			   GUICtrlSetState($InputBox, $GUI_DISABLE)
;			   GUICtrlSetData($InputBox, "")
;			EndIf
;		 Case $Forum
;			;$sComboRead = GUICtrlRead($Forum)
;			;ConsoleWrite("The combobox is currently displaying: " & $sComboRead & @CRLF)
;			If GUICtrlRead($Forum) = "Filme 1080p" or "" Then
;			   $ForumID = 96
;			ElseIf GUICtrlRead($Forum) = "Filme 720p" Then
;			   $ForumID = 103
;			ElseIf GUICtrlRead($Forum) = "3D HDTV 1080/720 Deutsch" Then
;			   $ForumID = 25
;			ElseIf GUICtrlRead($Forum) = "Serien 1080p" Then
;			   $ForumID = 161
;			ElseIf GUICtrlRead($Forum) = "Serien 720p" Then
;			   $ForumID = 114
;			EndIf
;			ConsoleWrite("The combobox is currently displaying: " & $ForumID & @CRLF)
;		 Case $StartButton
;			;If GUICtrlRead($StartButton) = "Start" Then
;			   GUICtrlSetState($StartButton, $GUI_HIDE)
;			   GUICtrlSetState($StopButton, $GUI_SHOW)
;			   CreateOutput()
;			;Else
;		 Case $StopButton
;			GUICtrlSetState($StopButton, $GUI_HIDE)
;			GUICtrlSetState($StartButton, $GUI_SHOW)
;			   ;GUICtrlSetData($StartButton, "Start")
;			   ;ExitLoop
;			   StopFunc()
;			;EndIf
;	EndSwitch
;WEnd
;EndFunc

Func OpenFile()
   Local Const $sMessage = "Quelldatei auswählen"
   Local $sFileOpenDialog = FileOpenDialog($sMessage, @WorkingDir & "\", "Text-File (*.txt)", $FD_FILEMUSTEXIST)
   If @error Then
	  ;Display error message when no File is selected.
	  If $sFileOpenDialog <> "" Then
		 MsgBox($MB_SYSTEMMODAL, "", "Keine Datei ausgewählt.")
	  EndIf
   Else
	  ;Replace instances of "|" with @CRLF in the string returned by FileOpenDialog.
	  $sFileOpenDialog = StringReplace($sFileOpenDialog, "|", @CRLF)
	  ;$sFileSource = $sFileOpenDialog
	  Dim $text
	  If Not _FileReadToArray($sFileOpenDialog, $text) Then
		 MsgBox(4096, "Error", " Error reading text file (" & $sFileOpenDialog & ") to Array error:" & @error)
		 Exit
	  EndIf
	  ConsoleWrite(UBound($text) -1 & " Titel gefunden in: " & $sFileOpenDialog & @CRLF)
	  GUICtrlSetData($AnzahlTitelText, UBound($text) -1 & " Titel gefunden")
	  GUICtrlSetState($Forum, $GUI_ENABLE)
	  GUICtrlSetState($Output, $GUI_ENABLE)
	  GUICtrlSetData($Quelle,"Quelldatei: " & $sFileOpenDialog)
	  GUICtrlSetState($Checkbox, $GUI_ENABLE)
	  GUICtrlSetState($Checkbox2, $GUI_ENABLE)
	  GUICtrlSetState($StartButton, $GUI_ENABLE)

	  Return $sFileOpenDialog
   EndIf
EndFunc


Func _IsChecked($idControlID)
   Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc   ;==>_IsChecked


Func CreateOutput()
  ; This check avoids GUI flicker
  If $EventCheck = 0 Then
	  GUICtrlSetState($StartButton, $GUI_HIDE)
	  GUICtrlSetState($StopButton, $GUI_SHOW)
  EndIf

  $M = UBound($text) -1

  $Interrupt = 0
  $EventCheck = 0


; Create a constant variable in Local scope of the filepath that will be read/written to.
;Local Const $sFilePath = $sFileOutput ;_WinAPI_GetTempFileName(@TempDir)
ConsoleWrite("Erzeuge Datei: " & $sFileOutput & @CRLF)

; Create a temporary file to write data to.
;If Not FileWrite($sFileOutput, "###### created by " & @UserName & " on " & $sLongDayName & " " & $stamp2 & ", " & $stamp & " ######" & @CRLF) Then
;   MsgBox($MB_SYSTEMMODAL, "", "An error occurred whilst writing the temporary file.")
;EndIf

; Open the file for writing (append to the end of a file) and store the handle to a variable.
Local $hFileOpen = FileOpen($sFileOutput, $FO_APPEND)
If $hFileOpen = -1 Then
   MsgBox($MB_SYSTEMMODAL, "", "An error occurred whilst writing the temporary file.")
   ; Return False
Else
   FileWrite($sFileOutput, "###### created by " & @UserName & " on " & $sLongDayName & " " & $stamp2 & ", " & $stamp & " ######" & @CRLF)
EndIf

   ;now filling File, here starts the Grab from U4A
;Global $prelink="https://usenet-4all.pw/forum/search.php?do=process&forumchoice[]=" & $ForumID & "&titleonly=1&query="



For $item = 1 To UBound($text) -1

	  ConsoleWrite(@CRLF & $text[$item] & " (" & $item & "/" & UBound($text) -1 & ")")

   ; GUI draw Progressbar
        $gui_perc = $gui_perc + (1/(UBound($text) -1)*100)
        GUICtrlSetData($gui_progressbar, $gui_perc)
        GUICtrlSetData($gui_progressbarLabel, Round($gui_perc, 2) & "%" & " (" & $item & "/" & UBound($text) -1 & ")")

      ;  Sleep(5)

   If GUICtrlRead($InputBox) <> "" Then
	  $linkinput = $text[$item] & " " & GUICtrlRead($InputBox)
   Else
	  $linkinput = $text[$item]
   EndIf
   $linkinput = StringReplace($linkinput, ",", "")
   $linkinput = StringReplace($linkinput, "&", "")
   $linkinput = StringReplace($linkinput, "-", " ")
   $linkinput = StringReplace($linkinput, "–", " ")
   $linkinput = StringReplace($linkinput, "ü", "ue")
   $linkinput = StringReplace($linkinput, "ä", "ae")
   $linkinput = StringReplace($linkinput, "!", "")
   $linkinput = StringReplace($linkinput, "ö", "oe")
   $linkinput = StringReplace($linkinput, "?", "")
   $linkinput = StringReplace($linkinput, "'", "")
   $linkinput = StringReplace($linkinput, "ß", "ss")
   $linkinput = StringReplace($linkinput, ".", "")
   $linkinput = StringReplace($linkinput, "/", " ")
   $linkinput = StringReplace($linkinput, "(", "")
   $linkinput = StringReplace($linkinput, ")", "")
   $linkinput = StringReplace($linkinput, '"', "")
   $linkinput = StringReplace($linkinput, "’", "")
   $linkinput = StringRegExpReplace($linkinput, "\[.*?\]", "")

    $link = StringRegExpReplace($linkinput, "\s+", "+")
   ConsoleWrite($link)
   $tries = 1
   $maxTries = 3
  While Not _FFConnect($IP, $iPort, $iTimeOut)
	  If $tries > $maxTries Then
		 MsgBox(64,"Error","Can't connect to FireFox")
		 Exit
	  EndIf
	  ;ShellExecute("chrome.exe", "--start-maximized")
	  Run('"C:\Program Files (x86)\Mozilla Firefox\firefox.exe" /repl /repl 4242 ')
	  $tries += 1
	  Sleep(90)
   If $Interrupt <> 0 Then
	  $EventCheck = 0
     Return
   EndIf
  WEnd
;GUICtrlSetData($gui_text1,$gui_text1 & ".")
;ConsoleWrite($gui_text1 & ".")
   Sleep(60)
       ; Check for Interruption
    If $Interrupt <> 0 Then
	  $EventCheck = 0
      Return
    EndIf
   If Not _FFOpenURL($prelink & $link) Then
	  ;MsgBox(64,"","Can't open link")
	  FileWriteLine($hFileOpen, $text[$item] & ": ERROR : Link lässt sich nicht öffnen, Sonderzeichen?")
   ; Exit
   EndIf
   Sleep(3000)
   ;$success = 0
   $aLinks = _FFLinksGetAll()
   ;_ArrayDisplay($aLinks, "Alle Elemente")
   Local $vValue = ".*tid_.*"
   Local $sColumn = 4
   Local $aResult = _ArrayFindAll($aLinks, $vValue, Default, Default, Default, 3, $sColumn)

   If @error Then
	  ; MsgBox($MB_SYSTEMMODAL, "Not Found", '"' & $vValue & '" was not found on column ' & $sColumn & '.')
	  If $noResult = 1 Then
		 FileWriteLine($hFileOpen, $text[$item] & ": keine Suchergebnis")
	  EndIf
	  ConsoleWrite(" 0 Suchergebnisse")
	  ;$gui_text2 = $text[$item] & " 0 Suchergebnisse"
   Else


   FileWriteLine($hFileOpen, @CRLF & $text[$item] & ": ")

   ;ConsoleWrite($aLinks[$aResult[0]], "Suchergebnisse")
   ;_ArrayDisplay($aResult, "Elemente mit Links")
   ConsoleWrite(Ubound($aResult) & " Suchergebnisse:" & @CRLF)

   For $item_found = 0 To UBound($aResult) -1
	  ;ConsoleWrite($item_found & "/" & UBound($aResult) -1 & @CRLF)
	  $Line = $aResult[$item_found]
	  ;$gui_text2 = $text[$item] & " " & $Line &" Suchergebnisse"
	  ConsoleWrite($aLinks[$Line][5] & " - ")
	  ConsoleWrite($aLinks[$Line][0] & @CRLF)
	  FileWriteLine($hFileOpen, $aLinks[$Line][5] & " - " & $aLinks[$Line][0])
	  If $Item_found = UBound($aResult) -1 Then
		 FileWriteLine($hFileOpen, " ") ; Leerzeile fuer Uebersichtlichkeit nach dem letzen Eintrag
	  EndIf
   Next

   EndIf

Next
EndFunc

Func StopFunc()
  GUICtrlSetState($StopButton, $GUI_HIDE)
  GUICtrlSetState($StartButton, $GUI_SHOW)
  ConsoleWrite(" >Stopped" & @CRLF)
EndFunc

Func _WM_COMMAND($hWnd, $Msg, $wParam, $lParam)
  If BitAND($wParam, 0x0000FFFF) =  $StopButton Then $Interrupt = 1
  Return $GUI_RUNDEFMSG
EndFunc

Func OpenOutputFile()
   ShellExecute($sFileOutput)
EndFunc

Func ThatExit()
   Exit
EndFunc
