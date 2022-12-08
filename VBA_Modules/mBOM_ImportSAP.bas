Attribute VB_Name = "mBOM_ImportSAP"

' Autors: Roger Fankhauser & Fadri Pestalozzi
' Last Update: 20.08.2020

' ====================================================================================
' =============== declarations to check if NUMLOCK and CAPSLOCK are on ===============
' ====================================================================================

Option Explicit

Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VK_CAPITAL = &H14
Private Const VK_NUMLOCK = &H90
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
           

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

' API declarations:

Private Declare Function GetVersionEx Lib "kernel32" _
   Alias "GetVersionExA" _
   (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Sub keybd_event Lib "user32" _
   (ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare Function GetKeyboardState Lib "user32" _
   (pbKeyState As Byte) As Long

Private Declare Function SetKeyboardState Lib "user32" _
   (lppbKeyState As Byte) As Long
   

' ===========================================================================================
' ============== check initial values for num and capslock to reset at the end ==============
' ===========================================================================================

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const kCapital = 20
Private Const kNumlock = 144
      
Public Function CapsLock() As Boolean
CapsLock = KeyState(kCapital)
End Function

Public Function NumLock() As Boolean
NumLock = KeyState(kNumlock)
End Function

Private Function KeyState(lKey As Long) As Boolean
KeyState = CBool(GetKeyState(lKey))
End Function

' ======= ToDo block keyboard and mouse input and popup windows while SAP interactions are running =======
'Declare Function BlockInput Lib "USER32.dll" (ByVal fBlockIt As Long) As Long
'Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub BOMImport_SAP()

Dim iniNumLock As Boolean, iniCapsLock As Boolean
iniNumLock = NumLock()
iniCapsLock = CapsLock()

' ensure NumLock is on and CapsLock is off
ToggleNumLock (True)
ToggleCapsLock (False)

' choose SRO# to export SAP BOM
' same as previously exportet CAD
Dim ws As Worksheet
Dim wsNames_CAD() As String, wsNames_SAP() As String, chosenSRO As String, chosenSRO_SAP As String, chosenSROSpaces As String, filenameSAP As String
Dim i As Integer, j As Integer, k As Integer, N_CAD As Integer
i = 0
j = 0
N_CAD = 0

' get worksheet names containing "CAD" and "SAP"
For Each ws In Worksheets
    If InStr(ws.Name, "CAD") <> 0 Then
        i = i + 1
        ReDim Preserve wsNames_CAD(1 To i)
        wsNames_CAD(i) = ws.Name
    ElseIf InStr(ws.Name, "SAP") <> 0 Then
        j = j + 1
        ReDim Preserve wsNames_SAP(1 To j)
        wsNames_SAP(j) = ws.Name
    End If
Next ws

N_CAD = i

' based on existing CAD BOM choose SAP BOM
Dim answer As Integer, answer2 As Integer

If N_CAD = 0 Then
    MsgBox "Noch keine CAD Stückliste vorhanden!" & vbNewLine & vbNewLine & "Zuerst Import CAD Stückliste", 16, "CAD Stückliste fehlt"
    Exit Sub
Else
    For i = 1 To N_CAD
        answer = MsgBox("Soll SAP Stückliste zu " & vbNewLine & wsNames_CAD(i) & vbNewLine & "importiert werden?", vbYesNo, "SAP BOM Import")
        If answer = vbYes Then
            
            ' get SRO# of chosen worksheet
            chosenSRO = Left(wsNames_CAD(i), 9)
            
            ' generate SAP input
            chosenSROSpaces = Left(chosenSRO, 3) & " " & Mid(chosenSRO, 4, 3) & " " & Right(chosenSRO, 3)
            
            ' generate names for SAP worksheet and temporary output file
            chosenSRO_SAP = chosenSRO & "_SAP"
            filenameSAP = chosenSRO & "_SAP.xls"
            
            ' check if that SRO# already has SAP export
            For k = 1 To j
                If InStr(wsNames_SAP(k), chosenSRO) <> 0 Then
                    answer2 = MsgBox("Ausgewählte SAP Stückliste bereits vorhanden." & vbNewLine & vbNewLine & "Soll diese durch neuen Import ersetzt werden?" & vbNewLine & vbNewLine & "Tabellenblatt " & chosenSRO_SAP & " wird dabei gelöscht.", vbYesNo, "SAP Stückliste bereits vorhanden")
                    If answer2 = vbYes Then
                        GoTo replaceSAP
                    Else
                        Exit Sub
                    End If
                End If
            Next k
            
            ' exit loop if a BOM was chosen
            GoTo BOM_chosen
        ElseIf i = N_CAD Then ' no worksheet name chosen
            MsgBox "Keine weitere CAD Stückliste zur Auswahl vorhanden!", 16, "Keine weitere CAD Stückliste zu Auswahl"
            Exit Sub
        End If
    Next i

' to replace, the sheet with old data has to be deleted
replaceSAP:

' check if sheet exists before deleting
Dim sheet As Worksheet
For Each sheet In ActiveWorkbook.Worksheets
     If sheet.Name = chosenSRO_SAP Then
        ' delete sheet with name chosenSRO_SAP
        ' disable warning
        Application.DisplayAlerts = False
        Sheets(chosenSRO_SAP).Delete
        Application.DisplayAlerts = True
     End If
Next sheet


BOM_chosen:
End If


' ========================================================
' ================== before opening SAP ==================
' ============== get user name and password ==============
' ========================================================
Dim SAP_UserName, SAP_PW As String

SAP_UserName = ThisWorkbook.Sheets("StLiVergleich").Range("A12")
SAP_PW = ThisWorkbook.Sheets("StLiVergleich").Range("A14")

' ========================================================
' =================== Part 0, open SAP ===================
' ========================================================

Dim SapGui As Object
Dim saplogon As Object
Dim connection
Dim Wshell As Object

' Defer error trapping.
On Error Resume Next

' Test to see if there is a copy of Inventor already running.
Dim SAPWasNotRunning As Boolean    ' Boolen Check.

' Variable to hold reference to SAP.
' Check if SAP is running, on Error SAP is not running
Set SapGui = GetObject("SAPGUI")
    If Err.Number <> 0 Then SAPWasNotRunning = True
Err.Clear    ' Clear Err object in case error occurred.

' when SAP is not running--> start SAP with shell script and reference it
If SAPWasNotRunning Then
    ' Set SAP Status back to False
    SAPWasNotRunning = False

    Dim Wshshell As Object
    ' Create Shell object
    Set Wshshell = CreateObject("Wscript.Shell")

    'Execute Shell Script to start SAP
    Wshshell.Run Chr(34) & ("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe") & Chr(34) & " " & "/INI_FILE" & "=" & Chr(34) & "\\longpathtoini\appl\Sap\saplogon\int\saplogon.ini" & Chr(34)

    ' Wait for Programm to start up
    Do Until Wshshell.AppActivate("SAP Logon")
        Application.Wait Now + TimeValue("0:00:01")
    Loop

    'clear object
    Set Wshell = Nothing

    ' Try reference again
    Set SapGui = GetObject("SAPGUI")
    If Err.Number <> 0 Then SAPWasNotRunning = True
    Err.Clear    ' Clear Err object in case error occurred.

    ' On Error it's not running and we couldn't start it--> End Programm
    If SAPWasNotRunning = True Then
        MsgBox "SAP konnte nicht gestartet werden oder keine Verbindung möglich! ", 16, "SAP BOM Import"
        Exit Sub
    End If
Else
    MsgBox "Bitte laufende SAP Sitzung schliessen(speichern)... ", 16, "SAP BOM Import"
    Exit Sub
End If

' ======================================================================================
' ======================= PART 1, generate temporary SAP BOM file ======================
' ====================== Reason: SRO IT does not allow scripting =======================
' ====================== use SendKeys(Keys, Wait) as a workaround ======================
' ======================================================================================
' list of key codes, https://www.contextures.com/excelvbasendkeys.html#keycombo
' ESC is {ESC}
' ENTER is ~
' SHIFT is +
' CONTROL is ^
' ALT is %

' temporary directory
Dim directory As String, CurrentFilename As String
directory = "C:\kt\WorkSpace\"
CurrentFilename = ThisWorkbook.Name

' wait for window to load
Application.Wait Now + TimeValue("0:00:01")

' Anmelden
''' SendKeys "%A"
''' even after waiting ALT+A does not work, reason unknown, use ENTER instead, SHIFT+ENTER for those who have own SAP-account
SendKeys "+~"

' wait for next window to load
Application.Wait Now + TimeValue("0:00:01")

' type user name and password
' check for optional user input
If StrComp(SAP_UserName, "") = 0 Then
    SendKeys "srochse"
    SendKeys "{TAB}"     ' go to next field
    SendKeys "srochs"
Else
    SendKeys SAP_UserName
    SendKeys "{TAB}"     ' go to next field
    SendKeys SAP_PW
End If
SendKeys "~" ' ENTER password

' wait for next window to load
Application.Wait Now + TimeValue("0:00:01")

' ensure SAP search field is open
SendKeys "~"
Application.Wait Now + TimeValue("0:00:01")
SendKeys "{ESC}" ' jump into search field if it was closed, remain inside if already there

' wait for next window to load
Application.Wait Now + TimeValue("0:00:01")

' SAP BOM + ENTER
SendKeys "cs12"
SendKeys "~"

' wait for next window to load
Application.Wait Now + TimeValue("0:00:01")

' SAP Material, input SRO# corresponding to WorksheetName_CAD
' SendKeys "565 097 018"
SendKeys chosenSROSpaces

' SAP Anwendung, use TAB to navigate to target field
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "1"
SendKeys "{TAB}"
SendKeys "pp01"

' wait for next window to load
Application.Wait Now + TimeValue("0:00:01")

' Ausführen, generate BOM
SendKeys "{F8}"

' wait for next window to load
Application.Wait Now + TimeValue("0:00:02")

' Lokale Datei, export to local file
SendKeys "^+{F9}" 'CTRL+SHIFT+F9

' wait for next window to load
Application.Wait Now + TimeValue("0:00:01")

' Text mit Tabulatoren
SendKeys "{DOWN}"
SendKeys "{TAB}"
SendKeys "~"

' wait for next window to load
Application.Wait Now + TimeValue("0:00:01")

' Dateiname
' SendKeys "565097018_SAP.xls"
SendKeys "^a" 'Select All (ctrl + A)
Application.Wait Now + TimeValue("0:00:01")
SendKeys "{DEL}" ' Delete Text
Application.Wait Now + TimeValue("0:00:01")
SendKeys filenameSAP

' Verzeichnis
SendKeys "+{TAB}"
Application.Wait Now + TimeValue("0:00:01")
SendKeys "^a" 'Select All (ctrl + A)
Application.Wait Now + TimeValue("0:00:01")
SendKeys "{DEL}" ' Delete Text
Application.Wait Now + TimeValue("0:00:01")
SendKeys directory
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "~"

' wait for next window to load
Application.Wait Now + TimeValue("0:00:01")

' SAP-GUI-Sicherheit, Zulassen
SendKeys "%Z"

' wait for next window to load
Application.Wait Now + TimeValue("0:00:01")

' close SAP export, 1/2
SendKeys "%{F4}"

' wait for next window to load
Application.Wait Now + TimeValue("0:00:02")  ' #12:00:02 AM#

' close SAP export, 2/2
SendKeys "{TAB}"
SendKeys "~"

' SAP Logon 740, reactivate first SAP window
SendKeys "%{TAB}"

' wait for next window to load
Application.Wait Now + TimeValue("0:00:01")

' SAP Logon 740, close
SendKeys "%{F4}"

' switch to active message window
SendKeys "%{ESC}"
SendKeys "%{TAB}"

' reset NumLock to original state, turn back on if it was on at the beginning
If iniNumLock = False Then
    ToggleNumLock (False)
End If

' reset CapsLock to original state, turn back on if it was on at the beginning
If iniCapsLock = True Then
    ToggleCapsLock (True)
End If

' =====================================================================================
' ================ Part 2, import the new SAP BOM into this macro file ================
' =====================================================================================

Dim checkObjektId As String
If Dir(directory & filenameSAP) = "" Then
    Debug.Print ("File does not exist")
    MsgBox "Import SAP Stückliste nicht erfolgreich!", 16, "Import nicht erfolgreich"
Else
    Dim dummyOutput As String
    dummyOutput = importSheets(directory, CurrentFilename, filenameSAP)
    
    ' ====================================================================================
    ' ============= Part 3, delete temporary SAP output file after importing =============
    ' ====================================================================================
    
    ' no need to check again if file exists
    ' If Dir(directory) <> "" Then
        Kill (directory & filenameSAP)
    ' End If
    
    
    ' ====================================================================================
    ' ================ Part 4, correct SAP raw data if columns misaligned ================
    ' ====================================================================================
    
    'Dim checkObjektId As String
    checkObjektId = Sheets(chosenSRO_SAP).Range("D10").Value ' can be at D10 instead of E10, e.g. Anatas 565097002
    
    Worksheets(chosenSRO_SAP).Activate
    Set ws = ActiveSheet
    
    If StrComp(checkObjektId, "ObjektId") = 0 Then
        Range("D:D").Insert
    End If
    
    Worksheets("StLiVergleich").Activate
    
    MsgBox "Import SAP Stückliste erfolgreich.", vbInformation, "Import erfolgreich"

End If

End Sub




Private Function importSheets(directory As String, CurrentFilename As String, SourceFilename As String)

' 1. declare variables and Worksheet object
Dim fileName As String, total As Integer, sheet As Worksheet

' 2. Turn off screen updating and displaying alerts.
Application.ScreenUpdating = False
Application.DisplayAlerts = False

' 3. Initialize the variable directory. We use the Dir function to find the first *.xl?? file stored in this directory.
fileName = Dir(directory & "*.xl??")
' Note: The Dir function supports the use of multiple character (*) and single character (?) wildcards to search for all different types of Excel files.

' 4. The variable fileName now holds the name of the first Excel file found in the directory. Add a Do While Loop.
Do While fileName <> ""
    
    ' import only from SourceFilename
    If StrComp(fileName, SourceFilename) = 0 Then
        ' 5. There is no simple way to copy worksheets from closed Excel files. Therefore we open the Excel file.
        Workbooks.Open (directory & fileName)
        
        ' 6. Import the sheets from the Excel file into import-sheet.xlsm.
        For Each sheet In Workbooks(fileName).Worksheets
            total = Workbooks(CurrentFilename).Worksheets.Count
            Workbooks(fileName).Worksheets(sheet.Name).Copy _
            after:=Workbooks(CurrentFilename).Worksheets(total)
        Next sheet
        ' Explanation: the variable total holds track of the total number of worksheets of import-sheet.xlsm. We use the Copy method of the Worksheet object to copy each worksheet and paste it after the last worksheet of import-sheets.xlsm.
        
        ' 7. Close the Excel file.
        Workbooks(fileName).Close
    End If
    
    ' 8. The Dir function is a special function. To get the other Excel files, you can use the Dir function again with no arguments.
    fileName = Dir()
    ' Note: When no more file names match, the Dir function returns a zero-length string (""). As a result, Excel VBA will leave the Do While loop.

Loop

' 9. Turn screen updating and displaying alerts back on
Application.ScreenUpdating = True
Application.DisplayAlerts = True

' 10. function output
importSheets = "done"

End Function







Public Sub ToggleCapsLock(TurnOn As Boolean)

    'To turn capslock on, set turnon to true
    'To turn capslock off, set turnon to false
    
      Dim bytKeys(255) As Byte
      Dim bCapsLockOn As Boolean
      
'Get status of the 256 virtual keys
      GetKeyboardState bytKeys(0)
      
      bCapsLockOn = bytKeys(VK_CAPITAL)
      Dim typOS As OSVERSIONINFO
      
      If bCapsLockOn <> TurnOn Then 'if current state <>
                                     'requested stae
        
       If typOS.dwPlatformId = _
           VER_PLATFORM_WIN32_WINDOWS Then  '=== Win95/98

          bytKeys(VK_CAPITAL) = 1
          SetKeyboardState bytKeys(0)

        Else    '=== WinNT/2000

        'Simulate Key Press
          keybd_event VK_CAPITAL, &H45, _
             KEYEVENTF_EXTENDEDKEY Or 0, 0
        'Simulate Key Release
          keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY _
             Or KEYEVENTF_KEYUP, 0
        End If
      End If

     
End Sub




Public Sub ToggleNumLock(TurnOn As Boolean)

    'To turn numlock on, set turnon to true
    'To turn numlock off, set turnon to false
    
      Dim bytKeys(255) As Byte
      Dim bnumLockOn As Boolean
      
'Get status of the 256 virtual keys
      GetKeyboardState bytKeys(0)
      
      bnumLockOn = bytKeys(VK_NUMLOCK)
      Dim typOS As OSVERSIONINFO
      
      If bnumLockOn <> TurnOn Then 'if current state <>
                                     'requested stae
        
       If typOS.dwPlatformId = _
           VER_PLATFORM_WIN32_WINDOWS Then  '=== Win95/98

          bytKeys(VK_NUMLOCK) = 1
          SetKeyboardState bytKeys(0)

        Else    '=== WinNT/2000

        'Simulate Key Press
          keybd_event VK_NUMLOCK, &H45, _
             KEYEVENTF_EXTENDEDKEY Or 0, 0
        'Simulate Key Release
          keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY _
             Or KEYEVENTF_KEYUP, 0
        End If
      End If
     
End Sub


