Attribute VB_Name = "mBOM_ImportCAD_Inventor"

Sub BOMImport_CAD()

' Authors: Roger Fankhauser & Fadri Pestalozzi
' Last Update: 25.04.2020

Dim oApp As Object
Dim SheetName As String
Dim IsBomOpen
Dim CompareWorkbook As Workbook
Dim BOMWorkbook As Workbook

' Create AssmeblyDocument object
' When Autodesk Inventor Object Library is not set in References, then an error occurs.
Dim oAssyDoc As AssemblyDocument

' Defer error trapping.
On Error Resume Next

' make weak assumption that active workbook is the target
Set CompareWorkbook = Application.ActiveWorkbook

' Variable to hold reference to Inventor.
' Test to see if there is a copy of Inventor already running.
Dim InventorWasNotRunning As Boolean    ' Boolen Check.

' Getobject function called without the first argument returns a
' reference to an instance of the application. If the application isn't
' running, an error occurs.
Set oApp = GetObject(, "Inventor.Application")
If Err.Number <> 0 Then InventorWasNotRunning = True
Err.Clear    ' Clear Err object in case error occurred.

' Check if Inventor is running otherwise quit programm
If InventorWasNotRunning = True Then
    MsgBox "Bitte Inventor Starten und Baugruppe öffnen! ", 16, "Inventor BOM Import"
    Exit Sub
    
' Check if doucment is open otherwise quit programm
ElseIf oApp.Documents.VisibleDocuments.Count = 0 Then
    MsgBox "Es ist kein Dokument geoeffnet", 16, "Unerwarteter Fehler"
    Exit Sub
    
    ' Check if active document is not equal part document
ElseIf oApp.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
    MsgBox "Es ist keine Baugruppe geöffnet: " & oApp.ActiveDocument.DocumentType, 16, "Inventor BOM Import"
    Exit Sub

Else
    ' When the document is assembly document
    ' Set a reference to the active document.
    
    answer = MsgBox("Aktive Baugruppe: " & oApp.ActiveDocument.DisplayName & vbNewLine & vbNewLine & "Soll diese Importiert werden?", vbYesNo, "CAD Inventor BOM Import")
    If answer = vbYes Then
    
        ' ======= check if already imported =======
        Dim chosenCAD As String
        chosenCAD = oApp.ActiveDocument.DisplayName ' get SRO# of CAD assembly to be imported
        chosenCAD = Left(chosenCAD, 9)
        chosen_CAD = chosenCAD & "_CAD" ' corresponding excel worksheet name

        ' get worksheet names containing "CAD"
        Dim wsNames_CAD() As String
        Dim i As Integer
        For Each ws In Worksheets
            If InStr(ws.Name, "CAD") <> 0 Then
                i = i + 1
                ReDim Preserve wsNames_CAD(1 To i)
                wsNames_CAD(i) = ws.Name
            End If
        Next ws

        ' check if target SRO# already has CAD export
        Dim j As Integer
        For j = 1 To i
            If InStr(wsNames_CAD(j), chosenCAD) <> 0 Then
                answer = MsgBox("Ausgewählte CAD Stückliste bereits vorhanden." & vbNewLine & vbNewLine & "Soll diese durch neuen Import ersetzt werden?" & vbNewLine & vbNewLine & "Tabellenblatt " & chosen_CAD & " wird dabei gelöscht.", vbYesNo, "CAD Stückliste bereits vorhanden")
                If answer = vbNo Then
                    Exit Sub
                End If
            End If
        Next j
            
        ' ======= import CAD =======
        
        ' Set the active Docuemnt
        Set oAssyDoc = oApp.ActiveDocument
        ' Get the custom property set.
        Dim customPropSet As PropertySet
        Set customPropSet = oAssyDoc.PropertySets.Item("Inventor User Defined Properties")
        
        Dim PDB_Name As Property
        Set PDB_Name = customPropSet.Item("PDB_Name")
        
        SheetName = PDB_Name.Expression & "_CAD"
        Debug.Print "Sheet Name: " & SheetName
    Else
        Exit Sub
    End If
    
End If

' Change Coursor to loading
Application.Cursor = xlWait

' Set a reference to the BOM
Dim oBOM As BOM
Set oBOM = oAssyDoc.ComponentDefinition.BOM
 
' Path to the xml file
Dim oPathXML As String
oPathXML = "Z:\CHS\KT_Engineering\CAD\BOM\Vorlagen\SRO000067428.xml"
 
' Path to the export file
Dim oPathBOM As String
Path = "C:\kt\WorkSpace\"
fileName = SheetName & "_BOM_StructuredAllLevels.xls"
oPathBOM = Path & fileName

' Check if file exists
If Not Dir(oPathBOM, vbDirectory) = vbNullString Then
    ' The xml-file exists now check if xlsx-file is open
    Debug.Print ("File already Exists")
    IsBomOpen = IsWorkBookOpen(oPathBOM)
    ' When the file is open:
    If IsBomOpen = True Then
        'closes BOM Excel and discards any changes that have been made to it.
        Workbooks(fileName).Close SaveChanges:=False
        Debug.Print ("File Closed...")
    End If
End If

' Set the structured view to 'all levels'
oBOM.StructuredViewFirstLevelOnly = False

' Set the delimiter to point
oBOM.StructuredViewDelimiter = "."

' Make sure that the structured view is enabled.
oBOM.StructuredViewEnabled = True

' Import the xml file
oBOM.ImportBOMCustomization (oPathXML)

 ' Set a reference to the "Structured" BOMView
Dim oStructuredBOMView As BOMView
Set oStructuredBOMView = oBOM.BOMViews.Item("Strukturiert")

Debug.Print "Start BOM Export to: " & oPathBOM
' Export the BOM view to a Text File Tab Delimited file
oStructuredBOMView.Export oPathBOM, kMicrosoftExcelFormat
Debug.Print "Exported..."

' Set the reference to the exported bom
Set BOMWorkbook = Application.Workbooks.Open(oPathBOM)

With CompareWorkbook
' Check if Sheet exists and delete it
    For Each ws In .Worksheets
     If ws.Name = SheetName Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        Debug.Print "Sheet Deleted"
     End If
    Next ws
End With

' Add new Sheet
CompareWorkbook.Sheets.Add(after:=CompareWorkbook.Sheets(CompareWorkbook.Sheets.Count)).Name = SheetName

' Copy The Data
lastrow = BOMWorkbook.Sheets(1).UsedRange.Rows.Count

'Object
CompareWorkbook.Sheets(SheetName).Range("A1", "A" & lastrow).Value = BOMWorkbook.Sheets(1).Range("B1", "B" & lastrow).Value
'Stücklistenstruktur
CompareWorkbook.Sheets(SheetName).Range("B1", "B" & lastrow).Value = BOMWorkbook.Sheets(1).Range("A1", "A" & lastrow).Value
'Anzahl
CompareWorkbook.Sheets(SheetName).Range("C1", "C" & lastrow).Value = BOMWorkbook.Sheets(1).Range("I1", "I" & lastrow).Value
'PDB_Name
CompareWorkbook.Sheets(SheetName).Range("D1", "D" & lastrow).Value = BOMWorkbook.Sheets(1).Range("F1", "F" & lastrow).Value
'PDB_Ident
CompareWorkbook.Sheets(SheetName).Range("E1", "E" & lastrow).Value = BOMWorkbook.Sheets(1).Range("E1", "E" & lastrow).Value
'PDB_Version
CompareWorkbook.Sheets(SheetName).Range("F1", "F" & lastrow).Value = BOMWorkbook.Sheets(1).Range("G1", "G" & lastrow).Value
'Titel
CompareWorkbook.Sheets(SheetName).Range("G1", "G" & lastrow).Value = BOMWorkbook.Sheets(1).Range("H1", "H" & lastrow).Value
'Material
CompareWorkbook.Sheets(SheetName).Range("H1", "H" & lastrow).Value = BOMWorkbook.Sheets(1).Range("D1", "D" & lastrow).Value
'Masse
CompareWorkbook.Sheets(SheetName).Range("I1", "I" & lastrow).Value = BOMWorkbook.Sheets(1).Range("C1", "C" & lastrow).Value

' Close BOM workbook
BOMWorkbook.Close

'Select first Sheet
CompareWorkbook.Sheets(1).Activate

' Change Coursor to standard
Application.Cursor = xlDefault
' Message to user
MsgBox "Import CAD Stückliste erfolgreich.", vbInformation, "Import erfolgreich"

End Sub


' check if workbook is open
Function IsWorkBookOpen(fileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open fileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function
