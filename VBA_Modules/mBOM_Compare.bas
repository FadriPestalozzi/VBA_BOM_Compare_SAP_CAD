Attribute VB_Name = "mBOM_compare"

' Author: Fadri Pestalozzi
' Last Update: 02.09.2020

' clear output data
Sub btnClearData_Click() ' button caption = "Auswertung löschen"
    Sheets("StLiVergleich").Range("B3:BB10000").Clear
End Sub

' ##################################################
' ################### MAIN START ###################
' ##################################################
Sub btnStLiVgl_Click()

' before code start
Application.ScreenUpdating = False

' clear old data at new run
btnClearData_Click

' ------ global syntax ------
'c* = CAD
's* = SAP
'd* = difference

' S = BaugruppenStufe
' A = Anzahl
' U = Unit = Einheit
' I = ID = SRO# = 565010246
' V = Version
' T = Titel
    
' ------ global initialize ------
Dim sheetNameHome As String
sheetNameHome = "StLiVergleich"

' counters
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim m As Integer
Dim n As Integer
Dim t As Integer

' temporary containers
Dim ok As String
Dim tmp As Variant
Dim tmpID As String
Dim tmpID2 As String
Dim tmpStr As String
Dim tmpStr0 As String
Dim tmpStr1 As String
Dim tmpStr2 As String
Dim tmpStr3 As String
Dim tmpInt As Integer
Dim intResult As Integer
Dim tmpDB As Double


' ------------ prepare read data ------------

' initialize
Dim sheetNameCAD As String
Dim sheetNameSAP As String
Dim firstCAD As Integer
Dim firstSAP As Integer
Dim lastCAD As Integer
Dim lastSAP As Integer

' get chosen data sheet names
sheetNameCAD = Sheets(sheetNameHome).Range("A4")
sheetNameSAP = Sheets(sheetNameHome).Range("A6")

' check if input data available
If StrComp(sheetNameCAD, "") = 0 Then
    MsgBox "Import CAD Stückliste ausführen" & vbNewLine & vbNewLine & "UND" & vbNewLine & vbNewLine & "Datenauswahl aktualisieren", 16, "Keine CAD Rohdaten vorhanden"
    Exit Sub
ElseIf StrComp(sheetNameSAP, "") = 0 Then
    MsgBox "Import SAP Stückliste ausführen" & vbNewLine & vbNewLine & "UND" & vbNewLine & vbNewLine & "Datenauswahl aktualisieren", 16, "Keine SAP Rohdaten vorhanden"
    Exit Sub
End If

' define data range
firstCAD = 2 ' CAD data starting on row 2
firstSAP = 12 ' SAP data starting on row 12
lastCAD = lastrow(sheetNameCAD)
lastSAP = lastrow(sheetNameSAP)

' ensure array length 1 To L
Dim cL As Integer
Dim sL As Integer
cL = lastCAD - firstCAD + 1
sL = lastSAP - firstSAP + 1


' ------------ read CAD as Variant ------------

Dim cSVar As Variant
Dim cAVar As Variant
Dim cIVar As Variant
Dim cJVar As Variant
Dim cVVar As Variant
Dim cTVar As Variant
Dim cUVar() As Variant ' specify length for dummy input
ReDim cUVar(1 To cL)

cSVar = readData(sheetNameCAD, firstCAD, lastCAD, "A") ' Struktur (Objekt)
cAVar = readData(sheetNameCAD, firstCAD, lastCAD, "C") ' Anzahl + Unit
cIVar = readData(sheetNameCAD, firstCAD, lastCAD, "D") ' PDB_Name
cJVar = readData(sheetNameCAD, firstCAD, lastCAD, "E") ' PDB_Ident
cVVar = readData(sheetNameCAD, firstCAD, lastCAD, "F") ' PDB_Version
cTVar = readData(sheetNameCAD, firstCAD, lastCAD, "G") ' Titel

For i = 1 To cL                                        ' Placeholder for units
    ' initialize CAD units
    cUVar(i) = " "
    
    ' if no title, write this into output
    If IsEmpty(cTVar(i)) Then
        cTVar(i) = "KEIN TITEL!"
    End If
Next i


' ------------ read SAP as Variant ------------

Dim sSVar As Variant
Dim sIVar As Variant
Dim sVVar As Variant
Dim sTVar As Variant
Dim sAVar As Variant
Dim sUVar As Variant

sSVar = readData(sheetNameSAP, firstSAP, lastSAP, "B") ' Struktur (Stufe)
sIVar = readData(sheetNameSAP, firstSAP, lastSAP, "E") ' PDB_Ident (ObjektId)
sVVar = readData(sheetNameSAP, firstSAP, lastSAP, "G") ' Version (ÄndNr)
sTVar = readData(sheetNameSAP, firstSAP, lastSAP, "H") ' Titel (Objektkurztekt)
sAVar = readData(sheetNameSAP, firstSAP, lastSAP, "I") ' Anzahl (Menge)
sUVar = readData(sheetNameSAP, firstSAP, lastSAP, "J") ' Unit (ME)


' ------------ extract BG structure to enable type Integer ------------

' CAD starting at level 0 and SAP at level 1
' --> raise CAD to same level as SAP by adding +1
cSVar = Stufen(cSVar, 1)
sSVar = Stufen(sSVar, 0)


' ------------ SAP ObjektID, delete empty spaces within ------------
sIVar = delSpaces(sIVar)


' ------------ SAP version, ensure same emptyType as CAD ------------
For i = 1 To sL
    If IsEmpty(sVVar(i)) Then
        sVVar(i) = " "
    End If
Next i


' ------------ CAD units ------------
For i = 1 To cL
    If VarType(cAVar(i)) = vbString Then ' VarType(string) = 8 = vbString
        tmpStr = cAVar(i)
        ' check for units "mm" and "m"
        ' remove unit to get value
        If InStr(tmpStr, "mm") Then
            t = InStr(tmpStr, "mm")
            tmpStr = Left(tmpStr, t - 1) ' remove unit and space before unit
            cAVar(i) = CDbl(tmpStr) / 1000   ' convert to base unit "m"
            cUVar(i) = "m"
        ElseIf InStr(tmpStr, "m") Then
            t = InStr(tmpStr, "m")
            tmpStr = Left(tmpStr, t - 1) ' remove unit and space before unit
            cAVar(i) = CDbl(tmpStr)
            cUVar(i) = "m"
        Else
            t = InStr(tmpStr, " ") ' find index of unknown unit, separated from value by empty space
            tmpStr0 = Right(tmpStr, Len(tmpStr) - t) ' unknown unit
            tmpStr = Left(tmpStr, t - 1) ' remove unit and space before unit
            tmpStr1 = Left(cIVar(i), 3)    ' name part 1
            tmpStr2 = Mid(cIVar(i), 3 + 1, 3) ' name part 2
            tmpStr3 = Right(cIVar(i), 3)   ' name part 3
            MsgBox "Unbekannte Einheit [" & tmpStr0 & "]" & vbNewLine & "in CAD Rohdaten" & " bei " & vbNewLine & tmpStr1 & " " & tmpStr2 & " " & tmpStr3, 16, "Rohdaten CAD: Unbekannte Einheit"
            cAVar(i) = CDbl(tmpStr)
            cUVar(i) = tmpStr1 ' store unknown unit for CAD
        End If
    End If
Next i



' ------------ SAP units ------------
For i = 1 To sL
    If StrComp(sUVar(i), "ST") = 0 Then
        sUVar(i) = " "
    ElseIf StrComp(sUVar(i), "M") = 0 Then
        sUVar(i) = "m"
'    Else ' unknown unit already stored for SAP  ' output included, no error message necessary
'        tmpStr = sIVar(i)
'        tmpStr1 = Left(tmpStr, 3)
'        tmpStr2 = Mid(tmpStr, 3 + 1, 3)
'        tmpStr3 = Right(tmpStr, 3)
'        MsgBox "Unbekannte Einheit [" & sUVar(i) & "]" & vbNewLine & "in SAP Rohdaten" & " bei " & vbNewLine & tmpStr1 & " " & tmpStr2 & " " & tmpStr3, 16, "Rohdaten SAP: Unbekannte Einheit"
    End If
Next i


' ----------------------------------------------------------------------------
' -------------------- convert data to Integer and String --------------------
' ------------------------ S contains integer values ------------------------
' ------------------------ A contains double values ------------------------
' ------------------------ rest contains string values -----------------------
' --------------------------- cSStr = CStr(cSVar) ----------------------------
' ----------------------------------------------------------------------------

' ------------- CAD

Dim cS() As Integer
Dim cA() As Double
Dim cU() As String
Dim cI() As String
Dim cJ() As String
Dim cV() As String
Dim cT() As String

ReDim cS(1 To cL)
ReDim cA(1 To cL)
ReDim cU(1 To cL)
ReDim cI(1 To cL)
ReDim cJ(1 To cL)
ReDim cV(1 To cL)
ReDim cT(1 To cL)

For i = 1 To cL
    cS(i) = cSVar(i)
    cA(i) = cAVar(i)
    cU(i) = cUVar(i)
    cI(i) = cIVar(i)
    cJ(i) = cJVar(i)
    cV(i) = cVVar(i)
    cT(i) = cTVar(i)
    
    ' ----- if cI(i)="HLP*" and IsEmpty(cJ(i))=0 then replace cI(i) with cJ(i) xx2
'    If StrComp(Left(cI(i), 3), "HLP") = 0 And IsEmpty(cJ(i)) = False Then
'        cIVar(i) = cJVar(i)
'        cI(i) = cJ(i)
'    End If
Next i



' ------------- SAP

Dim sS() As Integer
Dim sA() As Double
Dim sU() As String
Dim sI() As String
Dim sV() As String
Dim sT() As String

ReDim sS(1 To sL)
ReDim sA(1 To sL)
ReDim sU(1 To sL)
ReDim sI(1 To sL)
ReDim sV(1 To sL)
ReDim sT(1 To sL)

For i = 1 To sL
    sS(i) = sSVar(i)
    sA(i) = sAVar(i)
    sU(i) = sUVar(i)
    sI(i) = sIVar(i)
    sV(i) = sVVar(i)
    sT(i) = sTVar(i)
Next i




' ------------------------------------------------------------------
' ------------- ensure non-empty version " " and not "" ------------
' ------------ otherwise removed by following exclusion ------------
' ------------------------------------------------------------------

For i = 1 To sL
    If StrComp(sV(i), "") = 0 Then
        sV(i) = " "
    End If
Next i



' --------------------------------------------------------------------
' -------------- exclude unnecessary data from analysis --------------
' --------------------------------------------------------------------

' get initial characters to be excluded
Dim excludeStartsVar As Variant
Dim excludeStarts() As String
ReDim excludeStarts(1 To 1)

excludeStartsVar = readExcludeStarts(sheetNameHome)

' check if 1 or multiple exclusions needed
Dim excludeDim As Integer
excludeDim = NumberOfDimensions(excludeStartsVar)

If excludeDim = 0 Then
    If StrComp(excludeStartsVar, "") <> 0 Then
        ReDim excludeStarts(1 To UBound(excludeStartsVar))
    Else
        ReDim excludeStartsVar(1 To 1)
        excludeStartsVar(1) = " "
    End If
Else
    ReDim excludeStarts(1 To UBound(excludeStartsVar))
End If



' excludeStartsVar 2 excludeStarts
' convert type variant to type string if necessary
For i = 1 To UBound(excludeStarts)
    tmp = excludeStartsVar(i)
    If VarType(tmp) = vbString Then
        excludeStarts(i) = tmp
    Else
        excludeStarts(i) = CStr(tmp)
    End If
Next i



' ------------ identify data rows to be excluded ------------

' store first indices for exclusion
Dim cExcludeFirsts() As Integer
Dim sExcludeFirsts() As Integer
ReDim cExcludeFirsts(1 To 1)
ReDim sExcludeFirsts(1 To 1)

' store last indices for exclusion
Dim cExcludeLasts() As Integer
Dim sExcludeLasts() As Integer
ReDim cExcludeLasts(1 To 1)
ReDim sExcludeLasts(1 To 1)

k = 0
t = 0

' ------------ prepare exclusion CAD ------------

For i = 1 To cL
    tmpID = cI(i)

    For j = 1 To UBound(excludeStarts)
        
        ' current ID_start to exclude
        tmpStr = excludeStarts(j)
                
        intResult = StrComp(Left(tmpID, Len(tmpStr)), tmpStr)

        If intResult = 0 Then
            k = k + 1
            ReDim Preserve cExcludeFirsts(1 To k)
            ReDim Preserve cExcludeLasts(1 To k)
            cExcludeFirsts(k) = i
            cExcludeLasts(k) = lastIdxOfLowerLvls(cS, i)
           
           ' "exit for" if "i" excluded to avoid overwriting cExclude(i) with non-exclusion
           Exit For
        
        End If
    Next j
Next i


' --- CAD assembly can contain parts of identical ID
' IF number at S+1 identical to number at S THEN exclude that subordinate part

' temporary structure containers
Dim tmpS1 As Integer
Dim tmpS2 As Integer

For i = 1 To cL - 1
    tmpS1 = cS(i)
    tmpS2 = cS(i + 1)
    
    ' If structure increases by +1, i.e. if code entered an assembly Then
    If tmpS2 - tmpS1 = 1 Then
        tmpStr1 = cI(i) ' get ID of that assembly to check if there are parts inside with identical ID
        
        j = i + 1 ' initialize counter for parts in assembly
        
        ' loop through all part IDs within assembly
        While cS(j) = tmpS2 And j < cL ' while on same assembly level And within data range
            
            tmpStr2 = cI(j) ' ID of current part
            
            ' flag part to be excluded if ID identical to assembly
            If StrComp(tmpStr1, tmpStr2) = 0 Then
                k = k + 1
                ReDim Preserve cExcludeFirsts(1 To k)
                ReDim Preserve cExcludeLasts(1 To k)
                cExcludeFirsts(k) = j
                cExcludeLasts(k) = j
                ' Debug.Print tmpStr1
            End If
            j = j + 1 ' increment counter
        Wend
    End If
Next i



' ------------ prepare exclusion SAP ------------

For i = 1 To sL
    tmpID = sI(i)

    For j = 1 To UBound(excludeStarts)
        
        ' current ID_start to exclude
        tmpStr = excludeStarts(j)
        
        intResult = StrComp(Left(tmpID, Len(tmpStr)), tmpStr)

        If intResult = 0 Then
            t = t + 1
            ReDim Preserve sExcludeFirsts(1 To t)
            ReDim Preserve sExcludeLasts(1 To t)
            sExcludeFirsts(t) = i
            sExcludeLasts(t) = lastIdxOfLowerLvls(sS, i)
           
           ' "exit for" if "i" excluded to avoid overwriting cExclude(i) with non-exclusion
           Exit For
        
        End If
    Next j
Next i


' ---------------------------------------------------------
' -------------------- apply exclusion --------------------
' -------- recombine in same order after exclusion --------
' ---------------------------------------------------------

cAVar = applyExclusionRanges(cA, cExcludeFirsts, cExcludeLasts)
cIVar = applyExclusionRanges(cI, cExcludeFirsts, cExcludeLasts)
cSVar = applyExclusionRanges(cS, cExcludeFirsts, cExcludeLasts)
cTVar = applyExclusionRanges(cT, cExcludeFirsts, cExcludeLasts)
cVVar = applyExclusionRanges(cV, cExcludeFirsts, cExcludeLasts)
cUVar = applyExclusionRanges(cU, cExcludeFirsts, cExcludeLasts)

sAVar = applyExclusionRanges(sA, sExcludeFirsts, sExcludeLasts)
sIVar = applyExclusionRanges(sI, sExcludeFirsts, sExcludeLasts)
sSVar = applyExclusionRanges(sS, sExcludeFirsts, sExcludeLasts)
sTVar = applyExclusionRanges(sT, sExcludeFirsts, sExcludeLasts)
sVVar = applyExclusionRanges(sV, sExcludeFirsts, sExcludeLasts)
sUVar = applyExclusionRanges(sU, sExcludeFirsts, sExcludeLasts)

' reset array lengths
cL = UBound(cAVar)
sL = UBound(sAVar)

' ------ reset data arrays ------
ReDim cS(1 To cL)
ReDim cA(1 To cL)
ReDim cU(1 To cL)
ReDim cI(1 To cL)
ReDim cV(1 To cL)
ReDim cT(1 To cL)

ReDim sS(1 To sL)
ReDim sA(1 To sL)
ReDim sU(1 To sL)
ReDim sI(1 To sL)
ReDim sV(1 To sL)
ReDim sT(1 To sL)



' ------ convert once again away from variant ------
For i = 1 To cL
    cS(i) = cSVar(i)
    cA(i) = cAVar(i)
    cU(i) = cUVar(i)
    cI(i) = cIVar(i)
    cV(i) = cVVar(i)
    cT(i) = cTVar(i)
Next i


For i = 1 To sL
    sS(i) = sSVar(i)
    sA(i) = sAVar(i)
    sU(i) = sUVar(i)
    sI(i) = sIVar(i)
    sV(i) = sVVar(i)
    sT(i) = sTVar(i)
Next i


' ------------ list all IDs, CAD + SAP ------------

' VBA has no built-in function to concatenate arrays.
' Dimension an array with total length and paste source arrays
Dim csL As Integer
csL = cL + sL

Dim csI() As String
ReDim csI(1 To csL)
k = 0

For i = 1 To csL
    If i <= cL Then
        csI(i) = cI(i)
    Else
        k = k + 1
        csI(i) = sI(k)
    End If
Next i

' --------------- get unique values ---------------

' sort IDs in ascending order
' Call sort2(csI, csT)  ' sort data before key
Call sort2(csI, csI)  ' sort key  after  data

Dim csIUniqueVar() As Variant
Dim csTUniqueVar() As Variant

csIUniqueVar = getUnique(csI)
cTUniqueVar = getTitles(csIUniqueVar, cI, cT)
sTUniqueVar = getTitles(csIUniqueVar, sI, sT)

Dim csIUnique() As String
Dim cTUnique() As String
Dim sTUnique() As String

ReDim csIUnique(1 To UBound(csIUniqueVar))
ReDim cTUnique(1 To UBound(csIUniqueVar))
ReDim sTUnique(1 To UBound(csIUniqueVar))

' no variant
For i = 1 To UBound(csIUniqueVar)
    csIUnique(i) = csIUniqueVar(i)
    cTUnique(i) = cTUniqueVar(i)
    sTUnique(i) = sTUniqueVar(i)
Next i


' ----------------------------------------------------------------------
' --------------------------- generate posID ---------------------------
' ---------------- BEFORE CALLING notInCAD and notInSAP ----------------
' ----------------------------------------------------------------------

Dim cPosID() As Variant
ReDim cPosID(1 To cL)

Dim sPosID() As Variant
ReDim sPosID(1 To sL)

For i = 1 To cL
    cPosID(i) = concatenatePosID(cS, cI, i)
Next i

For i = 1 To sL
    sPosID(i) = concatenatePosID(sS, sI, i)
Next i


' -------------------------------------------------------------------
' --------------------------- compare IDs ---------------------------
' ---------------- to generate notInCAD and notInSAP ----------------
' -------------------------------------------------------------------

Dim notInCAD() As Variant
Dim notInSAP() As Variant
ReDim notInSAP(1 To cL)
ReDim notInCAD(1 To sL) ' max length of "not in CAD" = length "in SAP"

m = 0
n = 0

For h = 1 To UBound(csIUnique) ' loop through all IDs from CAD + SAP
    For j = 1 To cL ' loop through IDs in CAD
        For i = 1 To sL ' loop through IDs in SAP
            ' StrComp = 0 if strings are equal
            If StrComp(csIUnique(h), cI(j)) = 0 And _
                StrComp(csIUnique(h), sI(i)) = 0 Then
                GoTo nextUnique
            End If
        Next i
    Next j


    For j = 1 To cL ' loop through IDs in CAD
        If StrComp(csIUnique(h), cI(j)) = 0 Then ' only CAD matches --> not in SAP
            m = m + 1
            notInSAP(m) = csIUnique(h)
            GoTo nextUnique
        End If
    Next j

    For i = 1 To sL ' loop through IDs in SAP
        If StrComp(csIUnique(h), sI(i)) = 0 Then ' only SAP matches --> not in CAD
            n = n + 1
            notInCAD(n) = csIUnique(h)
            GoTo nextUnique
        End If
    Next i

nextUnique: ' jump here if unique ID was checked is in Both CAD and SAP
Next h

' trim output
notInCAD = deleteEmpty(notInCAD)
notInSAP = deleteEmpty(notInSAP)



' -------------------------------------------------------------------
' ------------ multiple occurrences of same ID possible -------------
' ---------------- posID and A can be longer than ID ----------------
' ------------------ redim preserve var(1 to end) -------------------
' -------------------------------------------------------------------

Dim notInCAD_ID() As Variant
Dim notInSAP_ID() As Variant
ReDim notInSAP_ID(1 To 1)
ReDim notInCAD_ID(1 To 1)

Dim notInCAD_pos() As Variant
Dim notInSAP_pos() As Variant
ReDim notInSAP_pos(1 To 1)
ReDim notInCAD_pos(1 To 1)

Dim notInCAD_A() As Variant
Dim notInSAP_A() As Variant
ReDim notInSAP_A(1 To 1)
ReDim notInCAD_A(1 To 1)

Dim notInCAD_T() As Variant
Dim notInSAP_T() As Variant
ReDim notInSAP_T(1 To 1)
ReDim notInCAD_T(1 To 1)

' counters for *_posID and *_A
m = 0
n = 0

' CAD two nested loops since ID can be used several times
For k = 1 To UBound(notInSAP)
    tmpStr = notInSAP(k)
    For i = 1 To cL
        If StrComp(tmpStr, cI(i)) = 0 Then ' ID found
            m = m + 1
            ReDim Preserve notInSAP_A(1 To m)
            ReDim Preserve notInSAP_ID(1 To m)
            ReDim Preserve notInSAP_pos(1 To m)
            ReDim Preserve notInSAP_T(1 To m)
            notInSAP_A(m) = cA(i)
            notInSAP_ID(m) = cI(i)
            notInSAP_pos(m) = cPosID(i)
            notInSAP_T(m) = cT(i)
        End If
    Next i
Next k


' SAP two nested loops since ID can be used several times
For t = 1 To UBound(notInCAD)
    tmpStr = notInCAD(t)
    For j = 1 To sL
        If StrComp(tmpStr, sI(j)) = 0 Then ' ID found
            n = n + 1
            ReDim Preserve notInCAD_A(1 To n)
            ReDim Preserve notInCAD_ID(1 To n)
            ReDim Preserve notInCAD_pos(1 To n)
            ReDim Preserve notInCAD_T(1 To n)
            notInCAD_A(n) = sA(j)
            notInCAD_ID(n) = sI(j)
            notInCAD_pos(n) = sPosID(j)
            notInCAD_T(n) = sT(j)
        End If
    Next j
Next t


' --------------------------------------------------------------
' -------- get IDs that are present in both CAD and SAP --------
' --------- csIUnique without (notInCAD and notInSAP) ----------
' --------------------------------------------------------------

' initialize
t = 0
Dim inCADandSAP() As Variant
Dim cinCADandSAPT() As Variant
Dim sinCADandSAPT() As Variant

For i = 1 To UBound(csIUnique)
    
    ' check notInCAD
    For j = 1 To UBound(notInCAD)
        If StrComp(csIUnique(i), notInCAD(j)) = 0 Then
            GoTo notInCADandSAP
        End If
    Next j

    ' check notInSAP
    For k = 1 To UBound(notInSAP)
        If StrComp(csIUnique(i), notInSAP(k)) = 0 Then
            GoTo notInCADandSAP
        End If
    Next k

    t = t + 1
    ReDim Preserve inCADandSAP(1 To t)
    ReDim Preserve cinCADandSAPT(1 To t)
    ReDim Preserve sinCADandSAPT(1 To t)
    
    inCADandSAP(t) = csIUnique(i)
    cinCADandSAPT(t) = cTUnique(i)
    sinCADandSAPT(t) = sTUnique(i)
    
notInCADandSAP:
Next i

Dim LinCADandSAP As Integer
LinCADandSAP = UBound(inCADandSAP)

' ----------------------------------------------
' -------------- compare versions --------------
' ----------------------------------------------

' initialize
Dim versionID() As Variant
Dim versionCAD() As Variant
Dim versionSAP() As Variant
Dim versionT() As Variant
ReDim versionID(1 To LinCADandSAP) ' deleteEmpty later
ReDim versionCAD(1 To LinCADandSAP) ' deleteEmpty later
ReDim versionSAP(1 To LinCADandSAP) ' deleteEmpty later
ReDim versionT(1 To LinCADandSAP) ' deleteEmpty later

m = 0

For i = 1 To LinCADandSAP
    ' inCADandSAP(i) ' ID to compare versions from CAD and SAP
    For j = 1 To cL
        If StrComp(inCADandSAP(i), cI(j)) = 0 Then ' j = matching index
            For k = 1 To sL
                If StrComp(inCADandSAP(i), sI(k)) = 0 Then ' k = matching index
                    If StrComp(cV(j), sV(k)) <> 0 Then ' different versions detected
                        ' cV(j)  ' version CAD
                        ' sV(k)  ' version SAP
                        m = m + 1
                        versionID(m) = cI(j)
                        versionT(m) = sT(k) ' should be same title in c and s
                        versionCAD(m) = cV(j)
                        versionSAP(m) = sV(k)
                        GoTo versionNext
                    End If
                End If
            Next k
        End If
    Next j
versionNext:
Next i


versionID = deleteEmpty(versionID)
versionT = deleteEmpty(versionT)
versionCAD = deleteEmpty(versionCAD)
versionSAP = deleteEmpty(versionSAP)


' --------------------------------------------------------------------
' ---------------------- calculate total amount ----------------------
' --------------------- loop through inCADandSAP ---------------------
' -------- if same ID then sum amount until top level reached --------
' --------------------------------------------------------------------

' initialize
Dim dATotID() As String  ' ID
Dim cdATotT() As String   ' title
Dim sdATotT() As String   ' title
Dim dATot() As Double   ' value

ReDim dATotID(1 To LinCADandSAP)
ReDim cdATotT(1 To LinCADandSAP)
ReDim sdATotT(1 To LinCADandSAP)
ReDim dATot(1 To LinCADandSAP)

Dim cAtotAtSameID() As Double
Dim sAtotAtSameID() As Double
ReDim cAtotAtSameID(1 To LinCADandSAP)
ReDim sAtotAtSameID(1 To LinCADandSAP)

Dim cTmpSum As Double
Dim sTmpSum As Double
k = 0


Dim tmpT As String ' title


' sumup structureAmounts for each unique ID present in both CAD & SAP
For i = 1 To LinCADandSAP

    ' temporary data
    tmpStr = inCADandSAP(i)
    ctmpT = cinCADandSAPT(i)
    stmpT = sinCADandSAPT(i)

    ' reset sum counters
    cTmpSum = 0
    sTmpSum = 0

''' Bugfixing xxx
''' cA ggb. cI+cS verschoben
''' leere Anzahl in readData verhindern

'If i = 25 Then
'a = 1
'End If

    ' --- CAD
    For j = 1 To cL
        If StrComp(tmpStr, cI(j)) = 0 Then ' for every instance of CAD-ID
            cTmpSum = cTmpSum + sumUntilTopLvl(cS, cA, j)
        End If
    Next j
    cAtotAtSameID(i) = cTmpSum

    ' --- SAP
    For j = 1 To sL
        If StrComp(tmpStr, sI(j)) = 0 Then ' for every instance of SAP-ID    SAP already summed until top level
            sTmpSum = sTmpSum + sA(j)             ' sumUntilTopLvl(sS2sumTop, sA2sumTop, j)
        End If
    Next j
    sAtotAtSameID(i) = sTmpSum

    ' --------------------- generate dATot ---------------------
    ' output if difference in total amount of current ID (=tmpStr)
    If cAtotAtSameID(i) <> sAtotAtSameID(i) Then
        k = k + 1
        dATot(k) = cAtotAtSameID(i) - sAtotAtSameID(i) ' Delta = CAD - SAP
        dATotID(k) = tmpStr
        cdATotT(k) = ctmpT
        sdATotT(k) = stmpT
    End If

Next i


' --------------------- ensure unique and remove empty

' initialize
Dim dAIout() As String  ' tmp for dATotID
Dim cdATout() As String
Dim sdATout() As String
Dim dAAout() As Double ' tmp for dATot
m = 0

' --- get unique dATotID
Dim dATotIDunique() As String
dATotID = deleteEmptyStr(dATotID) ' first remove empty values
Call QuickSort(dATotID, 1, UBound(dATotID)) ' need sorted input for getUnique
dATotIDunique = getUniqueStr(dATotID)

' --- check for unique entries
For i = 1 To UBound(dATotIDunique)
    tmpID = dATotIDunique(i)
    
    ' check current output
    For j = 1 To UBound(dATotID)
        tmpID2 = dATotID(j)
        
        ' store matching output and go to next unique
        If StrComp(tmpID, tmpID2) = 0 Then
            ' dynamically extend output arrays
            m = m + 1
            ReDim Preserve dAIout(1 To m)
            ReDim Preserve cdATout(1 To m)
            ReDim Preserve sdATout(1 To m)
            ReDim Preserve dAAout(1 To m)
            
            dAIout(m) = dATotID(j)
            cdATout(m) = cdATotT(j)
            sdATout(m) = sdATotT(j)
            dAAout(m) = dATot(j)
            GoTo dATotIDnext
        End If
    Next j
dATotIDnext:
Next i

' store output
ReDim dATotID(1 To m)
ReDim cdATotT(1 To m)
ReDim sdATotT(1 To m)
ReDim dATot(1 To m)

dATotID = dAIout
cdATotT = cdATout
sdATotT = sdATout
dATot = dAAout


' ------------------------------------------------------
' ------------ generate csU = update output ------------
' ------------- apply position update "U" --------------
' ---------- and output data arrays {I,pos,A} ----------
' ------------------------------------------------------

Dim cTmp As String
Dim sTmp As String

Dim cUA As Variant
Dim sUA As Variant
ReDim cUA(1 To 1)
ReDim sUA(1 To 1)

Dim cUI As Variant
Dim sUI As Variant
ReDim cUI(1 To 1)
ReDim sUI(1 To 1)

Dim cUP As Variant
Dim sUP As Variant
ReDim cUP(1 To 1)
ReDim sUP(1 To 1)

Dim cUTT As Variant
Dim sUT As Variant
ReDim cUTT(1 To 1)
ReDim sUT(1 To 1)

k = 0
t = 0

' for all unique IDs present in both CAD + SAP
For i = 1 To UBound(dATotID)

    ' current unique ID
    tmpID = dATotID(i)

    ' CAD
    For m = 1 To cL

        ' current unique overlap ID on CAD-side
        cTmp = cI(m)

        If StrComp(tmpID, cTmp) = 0 Then
        ' identical found, store output data
        ' multiple occurrences of same ID possible
            k = k + 1
            ReDim Preserve cUI(1 To k)
            ReDim Preserve cUP(1 To k)
            ReDim Preserve cUA(1 To k)
            ReDim Preserve cUTT(1 To k)
            cUI(k) = cI(m)
            cUP(k) = cPosID(m)
            cUA(k) = cA(m)
            cUTT(k) = cT(m)
        End If
    Next m

    ' SAP
    For n = 1 To sL

        ' current unique overlap ID on SAP-side
        sTmp = sI(n)

        If StrComp(tmpID, sTmp) = 0 Then
        ' identical found, store output data
        ' multiple occurrences of same ID possible
            t = t + 1
            ReDim Preserve sUI(1 To t)
            ReDim Preserve sUP(1 To t)
            ReDim Preserve sUA(1 To t)
            ReDim Preserve sUT(1 To t)
            sUI(t) = sI(n)
            sUP(t) = sPosID(n)
            sUA(t) = sA(n)
            sUT(t) = sT(n)
        End If
    Next n
Next i


' -------------------------------------------------
' ----------- output only if cUA <> sUA -----------
' ----------------- generate csUDel ----------------
' -------------------------------------------------

' initialize arr to store items to be removed
Dim cUDel() As Integer
Dim sUDel() As Integer
ReDim Preserve cUDel(1 To 1)
ReDim Preserve sUDel(1 To 1)


' ------- loop twice to flag entries where:
' ------- posID1 = posID2
' ------- AND
' ------- A1 = A2


' ------- CAD, 1st loop to identify items to be removed from data
k = 0
For i = 1 To UBound(cUP)
    cTmp = cUP(i) ' current CAD-position
    For j = 1 To UBound(sUP)
        sTmp = sUP(j) ' compare to SAP-positions
        If StrComp(cTmp, sTmp) = 0 Then ' same BG position
            If cUA(i) = sUA(j) Then ' also same amount --> flag index to be removed
                k = k + 1
                ReDim Preserve cUDel(1 To k)
                cUDel(k) = i
            End If
        End If
    Next j
Next i


' ------- SAP, 1st loop to identify items to be removed from data
t = 0
For i = 1 To UBound(sUP)
    sTmp = sUP(i) ' current SAP-position
    For j = 1 To UBound(cUP)
        cTmp = cUP(j) ' compare to CAD-positions
        If StrComp(sTmp, cTmp) = 0 Then ' same BG position
            If sUA(i) = cUA(j) Then ' also same amount --> flag index to be removed
                t = t + 1
                ReDim Preserve sUDel(1 To t)
                sUDel(t) = i
            End If
        End If
    Next j
Next i


' --------------------------------------------------
' ------------------- Backup csU -------------------
' --------------------------------------------------

Dim cUA_bup As Variant
Dim sUA_bup As Variant
cUA_bup = cUA
sUA_bup = sUA

Dim cUI_bup As Variant
Dim sUI_bup As Variant
cUI_bup = cUI
sUI_bup = sUI

Dim cUP_bup As Variant
Dim sUP_bup As Variant
cUP_bup = cUP
sUP_bup = sUP

Dim cUTT_bup As Variant
Dim sUT_bup As Variant
cUTT_bup = cUTT
sUT_bup = sUT


' ---------------------------------------------------------------------------
' --------- check if for every dATotID at least one csPosID present ---------
' ----------- add all csPosID if either c or s is empty at dATotID ----------
' ---------------------------------------------------------------------------

' ----------- find missing entries in csU by comparing to dAtotID


' find missing cU-output

m = 0
Dim cUNotFound() As Integer
ReDim cUNotFound(1 To 1)
cUNotFound(1) = 0

For i = 1 To UBound(dATotID) ' check all IDs with dATot <> 0
    tmpID = dATotID(i)
    For j = 1 To UBound(cUI) ' compare with current output IDs
        tmpStr = cUI(j)
        If StrComp(tmpID, tmpStr) = 0 Then ' go to next tmpID if match found
            GoTo cUFound
        End If
    Next j
    ' not found if all cUI checked and arrived at this code location
    m = m + 1
    ReDim Preserve cUNotFound(1 To m)
    cUNotFound(m) = i
cUFound:
Next i

' find missing sU-output

n = 0
Dim sUNotFound() As Integer
ReDim sUNotFound(1 To 1)
sUNotFound(1) = 0

For i = 1 To UBound(dATotID) ' check all IDs with dATot <> 0
    tmpID = dATotID(i)
    For j = 1 To UBound(sUI) ' compare with current output IDs
        tmpStr = sUI(j)
        If StrComp(tmpID, tmpStr) = 0 Then ' go to next tmpID if match found
            GoTo sUFound
        End If
    Next j
    ' not found if all cUI checked and arrived at this code location
    n = n + 1
    ReDim Preserve sUNotFound(1 To n)
    sUNotFound(n) = i
sUFound:
Next i



' ----------- add missing output by applying csUNotFound to csU_bup
' ----------- get ID by applying csU to dAtotID

Dim addElemLocation As Integer
Dim addElemTmpID As String

' --------- add missing to CAD if csUNotFound is non-empty

If cUNotFound(1) <> 0 Then ' check if csUNotFound is non-empty

For i = 1 To UBound(cUNotFound)
    tmpInt = cUNotFound(i) ' index of ID in dATotID which was not found in CAD-output
    tmpStr = dATotID(tmpInt) ' ID which was not found
    For j = 1 To UBound(cUI_bup)
        ' check all IDs in backup-output (bup = before removal) --> get input for addElem
        tmpID = cUI_bup(j)
        If StrComp(tmpStr, tmpID) = 0 Then ' add missing output from backup arrays
            ' find location where to add ElemId
            ' IDs are sorted from low to high
            For k = 1 To UBound(dATotID)
                addElemTmpID = dATotID(k)
                ' ID to insert is larger i.e. comes after variable of innermost loop
                If StrComp(addElemTmpID, tmpID) > 0 Then
                    ' target insertion location = k - 1
                    addElemLocation = k - 1
                    GoTo cAddElemLocationFound
                End If
            Next k
cAddElemLocationFound:
            cUA = addElemToArray(cUA, cUA_bup(j), addElemLocation)
            cUI = addElemToArray(cUI, cUI_bup(j), addElemLocation)
            cUP = addElemToArray(cUP, cUP_bup(j), addElemLocation)
        End If
    Next j
Next i

End If


' --------- add missing to SAP if csUNotFound is non-empty

If sUNotFound(1) <> 0 Then ' check if csUNotFound is non-empty

For i = 1 To UBound(sUNotFound)
    tmpInt = sUNotFound(i) ' index of ID in dATotID which was not found in SAP-output
    tmpStr = dATotID(tmpInt) ' ID which was not found
    For j = 1 To UBound(sUI_bup)
        ' check all IDs in backup-output (bup = before removal) --> get input for addElem
        tmpID = sUI_bup(j)
        If StrComp(tmpStr, tmpID) = 0 Then ' add missing output from backup arrays
            ' find location where to add ElemId
            ' IDs are sorted from low to high
            For t = 1 To UBound(dATotID)
                addElemTmpID = dATotID(t)
                ' ID to insert is larger i.e. comes after variable of innermost loop
                If StrComp(addElemTmpID, tmpID) > 0 Then
                    ' target insertion location = t - 1
                    addElemLocation = t - 1
                    GoTo sAddElemLocationFound
                End If
            Next t
sAddElemLocationFound:
            sUA = addElemToArray(sUA, sUA_bup(j), addElemLocation)
            sUI = addElemToArray(sUI, sUI_bup(j), addElemLocation)
            sUP = addElemToArray(sUP, sUP_bup(j), addElemLocation)
        End If
    Next j
Next i

End If


' -------------------------------------------------------------------------------------------
' ---------- prep output for unknown units (UU), evtl. unbekannte Einheiten im SAP ----------
' -------------------------------------------------------------------------------------------

Dim sUUI() As String ' ID
Dim sUUT() As String ' Titel
Dim sUUP() As String ' Position
Dim sUUA() As Double ' Menge, amount
Dim sUU() As String  ' unit

' initialize output to start with no unknown units
ReDim sUUI(1 To 1)
ReDim sUUT(1 To 1)
ReDim sUUP(1 To 1)
ReDim sUUA(1 To 1)
ReDim sUU(1 To 1)

sUUI(1) = "none"
sUUT(1) = "none"
sUUP(1) = "none"
sUUA(1) = 0
sUU(1) = "none"

n = 0

For i = 1 To sL
    If StrComp(sU(i), " ") = 0 Or StrComp(sU(i), "m") = 0 Then
        GoTo checkNextUnit
    Else
        ' not found if all cUI checked and arrived at this code location
        n = n + 1
        ReDim Preserve sUUI(1 To n)
        ReDim Preserve sUUT(1 To n)
        ReDim Preserve sUUP(1 To n)
        ReDim Preserve sUUA(1 To n)
        ReDim Preserve sUU(1 To n)
        
        sUUI(n) = sI(i)
        sUUT(n) = sT(i)
        sUUP(n) = sPosID(i)
        sUUA(n) = sA(i)
        sUU(n) = sU(i)
        
    End If
checkNextUnit:
Next i


' ------------------------------------------------------------------
' ------------------- WRITE OUTPUT DATA TO EXCEL -------------------
' ------------------------------------------------------------------

' -------- first output = name of top assembly which was compared here --------

' initialize
Dim topBGVar As Variant
Dim topBG As String
Dim topBG_Range As String

' get data
topBG = Sheets(sheetNameSAP).Range("F3").Value
topBGVar = delSpaces(topBG)
topBG = CStr(topBGVar(1))
topBG_Range = "B3"

' output
Sheets(sheetNameHome).Range(topBG_Range).HorizontalAlignment = xlCenter
Sheets(sheetNameHome).Range(topBG_Range).VerticalAlignment = xlCenter
Sheets(sheetNameHome).Range(topBG_Range) = topBG
'Sheets(sheetNameHome).Range(topBG_Range).FontSize = 16 ' Error, cannot define FontSize this way

' ---------------- output data columns ----------------

' counter for output columns
Dim colL As String ' column Letter
Dim colN As Integer ' column Number

colN = 2 ' skip columns "A" and "B"


' ----------- only in CAD

' notInSAP_ID
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(notInSAP_ID, sheetNameHome, colL)

' notInSAP_T
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(notInSAP_T, sheetNameHome, colL, 1)

' notInSAP_pos
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(notInSAP_pos, sheetNameHome, colL, 1) ' leftAlign posID

' notInSAP_A
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(notInSAP_A, sheetNameHome, colL)

' ----------- only in SAP

' notInCAD_ID
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(notInCAD_ID, sheetNameHome, colL)

' notInCAD_T
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(notInCAD_T, sheetNameHome, colL, 1)

' notInCAD_pos
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(notInCAD_pos, sheetNameHome, colL, 1) ' leftAlign posID

' notInCAD_A
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(notInCAD_A, sheetNameHome, colL)


' ------------ compare current versions

' versionID = ID with different versions
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(versionID, sheetNameHome, colL)

' versionT
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(versionT, sheetNameHome, colL, 1)

' current version CAD
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(versionCAD, sheetNameHome, colL)

' current version SAP
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(versionSAP, sheetNameHome, colL)


' ------------ difference in total amount

' dATotID = total amount ID
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(dATotID, sheetNameHome, colL)

' --- output both CAD and SAP titles
' --- cdATotT = total amount title CAD
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(cdATotT, sheetNameHome, colL, 1)

' --- sdATotT = total amount title SAP
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(sdATotT, sheetNameHome, colL, 1)


' dATot = total amount
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(dATot, sheetNameHome, colL)


' -------- structure update --------

' ------ structure update CAD

' cUI = ID where different structure found in CAD
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(cUI, sheetNameHome, colL)

' cUTT = CAD title --> cUT is a VBA-specific variable name
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(cUTT, sheetNameHome, colL, 1)

' cUP = current CAD structure
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(cUP, sheetNameHome, colL, 1)

' cUA = current CAD amount
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(cUA, sheetNameHome, colL) '


' ------ structure update SAP

' sUI = ID where different structure found in SAP
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(sUI, sheetNameHome, colL)

' sUT = SAP title
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(sUT, sheetNameHome, colL, 1)

' sUP = current SAP structure
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(sUP, sheetNameHome, colL, 1)

' sUA = current SAP amount
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(sUA, sheetNameHome, colL) '


' ------ unknown units SAP

' if no unknown units, suppress output
If StrComp(sUUI(1), "none") <> 0 Then

' sUUI = ID where unknown unit found in SAP
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(sUUI, sheetNameHome, colL)

' sUUT = title of unknown unit
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(sUUT, sheetNameHome, colL)

' sUUP = position of unknown unit
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(sUUP, sheetNameHome, colL)

' sUUA = amount of unknown unit
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(sUUA, sheetNameHome, colL)

' sUU = unknown unit
colN = colN + 1
colL = Number2Letter(colN)
ok = outputData(sUU, sheetNameHome, colL)

End If




' -------------- visualize output --------------

' color every other row
Dim rowN As Integer
For i = 1 To 500
    rowN = 2 + 2 * i
    Sheets(sheetNameHome).Range("B" & rowN & ":" & colL & rowN).Interior.Color = RGB(230, 230, 230)
Next i

' separate rows by vertical lines
' content sections by thick vertical lines
rowN = 0

' array of strings to hold right borders of content sections
Dim contentSections() As String
contentSections = Split("B,F,J,N,R,V,Z,AE", ",")

' content section counter, array starting at index 0
k = 0

' loop through column letters
For i = 1 To Letter2Number(contentSections(UBound(contentSections)))

    ' column letter
    colL = Number2Letter(i)

    ' draw vertical line
    With Sheets(sheetNameHome).Range(colL & "1:" & colL & "1000").Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)

        ' thick border if content section detected by StrComp
        If StrComp(colL, contentSections(k)) = 0 Then
            k = k + 1
            .Weight = xlThick
        Else
            .Weight = xlThin
        End If
    End With
Next i



' after code end
Application.ScreenUpdating = True

' ################################################
' ################### MAIN END ###################
' ################################################
End Sub



' ToDo NEXT
' get either CAD or SAP titles for unique IDs
Private Function getTitles(uniqueIDs() As Variant, IDs() As String, Titles() As String)

' initialize array lengths
Dim Nunique As Integer
Dim Ndata As Integer
Nunique = UBound(uniqueIDs)
Ndata = UBound(IDs)

Dim uniqueTitles() As String
ReDim uniqueTitles(1 To Nunique) As String

For j = 1 To Nunique
    For i = 1 To Ndata
        If StrComp(IDs(i), uniqueIDs(j)) = 0 Then
            uniqueTitles(j) = Titles(i)
            GoTo nextUnique
        End If
    Next i
nextUnique:
Next j

getTitles = uniqueTitles

End Function




' -----------------------------------------------------
' sort2(arr1,arr2)
' sort data in arr2 by key value in arr1 in ascending order

' This sort2 example outputs the order "4", "1", "2", "3", "5"
'Sub MAIN()
'    Dim Name(), Street()
'    Name = Array("B", "C", "D", "A", "E")
'    Street = Array("1", "2", "3", "4", "5")
'
'    Call sort2(Name(), Street())
'
'    For Each S In Street
'        MsgBox S
'    Next S
'End Sub


' sort arr2 by ascending key values in arr1
Sub sort2(key() As String, other() As String)
Dim i As Long, j As Long, Low As Long
Dim Hi As Long, Temp As String
    Low = LBound(key)
    Hi = UBound(key)

    j = (Hi - Low + 1) \ 2
    Do While j > 0
        For i = Low To Hi - j
          If key(i) > key(i + j) Then
            Temp = key(i)
            key(i) = key(i + j)
            key(i + j) = Temp
            Temp = other(i)
            other(i) = other(i + j)
            other(i + j) = Temp
          End If
        Next i
        For i = Hi - j To Low Step -1
          If key(i) > key(i + j) Then
            Temp = key(i)
            key(i) = key(i + j)
            key(i + j) = Temp
            Temp = other(i)
            other(i) = other(i + j)
            other(i + j) = Temp
          End If
        Next i
        j = j \ 2
    Loop
End Sub


' insert value at target location into array
' preserve array content by resizing and shifting data downward
Private Function addElemToArray(arrIn As Variant, NewVal As Variant, ElemId As Integer)

Dim i As Integer
Dim a As Variant
a = arrIn

ReDim Preserve a(1 To UBound(a) + 1)

For i = UBound(a) To ElemId + 1 Step -1
    a(i) = a(i - 1)
Next i
a(ElemId) = NewVal

addElemToArray = a

End Function


' sumup A2Top for all ID=tgtID
' ID(1 To N)
' A2Top(1 To N)
Private Function atSameID_Atot(ID As Variant, A2Top As Variant, tgtID As String)

' initialize counters
Dim i As Integer
Dim m As Integer
Dim tmpID As String
m = 0

' initialize dynamic array holding values to be summed up
Dim atSameID_A2Top() As Double
ReDim atSameID_A2Top(1 To 1)

' initialize output sum
Dim sumOut As Double
sumOut = 0

' loop through IDs
For i = 1 To UBound(ID)
    
    ' current ID
    tmpID = ID(i)
    
    ' if ID found --> output data
    If StrComp(tmpID, tgtID) = 0 Then
        
        m = m + 1 ' dynamically extend output array
        ReDim Preserve atSameID_A2Top(1 To m)
        atSameID_A2Top(m) = A2Top(i)
        
    End If
Next i

' calculate output sum
' counter m stopped at outputLength
For i = 1 To m
    sumOut = sumOut + atSameID_A2Top(i)
Next i

atSameID_Atot = sumOut


End Function


' convert letter to corresponding number
Private Function Letter2Number(inputL As String)
    Letter2Number = Range(inputL & 1).Column
End Function


' convert number to corresponding letter
Private Function Number2Letter(inputN As Integer)
    Number2Letter = Split(Cells(1, inputN).Address, "$")(1)
End Function


' output data to target worksheet and column
Private Function outputData(dataArr As Variant, sheetNameHome As String, tgtColumn As String, Optional leftAlign As Integer)
    Dim shiftRows As Integer
    shiftRows = 2
    
    For i = 1 To UBound(dataArr)
        
        Sheets(sheetNameHome).Range(tgtColumn & i + shiftRows).VerticalAlignment = xlCenter
        
        ' left align structure output
        If leftAlign = 1 Then
            Sheets(sheetNameHome).Range(tgtColumn & i + shiftRows).HorizontalAlignment = xlLeft
        Else
            Sheets(sheetNameHome).Range(tgtColumn & i + shiftRows).HorizontalAlignment = xlCenter
        End If
        
        ' -------- plot data after alignment
        Sheets(sheetNameHome).Range(tgtColumn & i + shiftRows) = dataArr(i)
        
    Next i
    outputData = "done"
End Function


' use arrays "structure" and "ID" to create posID at target index
Private Function concatenatePosID(S, ID, idx As Integer)

    ' initialize posID with current name at deepest level
    ' start potential loop if depth>1
    Dim posID As String
    posID = ID(idx)
    
    ' array to contain indices of assembly levels branching from input idx to main BG
    Dim indices() As Integer
    ReDim indices(1 To WorksheetFunction.Max(S))
    indices(1) = idx
    
    ' structure depth = number of names to concatenate until top level
    Dim depth As Integer
    depth = S(idx)
    
    If depth = 1 Then
        posID = "BG_" & posID
    End If
    
    ' loop through structure
    If depth > 1 Then
        For i = 2 To depth

            ' index of higher assembly level
            indices(i) = idxOfUpperLvl(S, indices(i - 1))

            ' concatenate new name at beginning of posID
            posID = ID(indices(i)) & "_" & posID

            ' if top level reached, finish posID by adding BG_ at the beginning
            If i = depth Then
                posID = "BG_" & posID
                GoTo posID_done
            End If
        Next i
    End If
    
posID_done:
    concatenatePosID = posID ' output result

End Function


' sum-up amounts while propagating backwards through structure until S=1
' S = Stufe, structure, level
' A = Anzahl, amount
Private Function sumUntilTopLvl(S As Variant, a As Variant, idx As Integer)
    
    ' initialize
    Dim i As Integer
    Dim m As Integer
    m = S(idx) ' input = starting level for back-propagation
    
    Dim amounts() As Double ' array to hold amount at each level
    ReDim amounts(1 To m)
    
    Dim idxLvlUp As Integer ' idx for back-propagation
    idxLvlUp = idx
    
    For i = 1 To m
        amounts(i) = a(idxLvlUp)
        idxLvlUp = idxOfUpperLvl(S, idxLvlUp)
    Next i

    sumUntilTopLvl = amounts(1)
    i = 0
    If m > 1 Then
        For i = 2 To m
            sumUntilTopLvl = sumUntilTopLvl * amounts(i)
        Next i
    End If

End Function


' get index of upper level, i.e. assembly wherein target idx is included
' S = Stufe, structure, level
Private Function idxOfUpperLvl(S As Variant, idx As Integer)
    
    ' initialize
    Dim i As Integer
    i = idx
    
    ' S(i) = structure "S" with index "i"
    ' top level at S=1
    If S(idx) = 1 Then ' this also covers i=1, since S(1)=1
        idxOfUpperLvl = idx
    Else ' search backwards until next higher level reached
        While S(idx) <= S(i - 1)
        ' continue search until higher level reached (S=2 is above S=3).
        ' stop searching if e.g. S(idx)=2 and S(i-1)=1
            i = i - 1 ' Decrement Counter
        Wend
        idxOfUpperLvl = i - 1 ' search ended at targetIdx + 1
    End If
End Function


' get indices of all lower levels, i.e. parts inside assembly
' S = Stufe, structure, level
Private Function lastIdxOfLowerLvls(S As Variant, idx As Integer) As Integer
    
' -------- debugging --------
'Dim idxOut As Integer
'Dim idxIn As Integer
'idxIn = 9
'
'idxOut = lastIdxOfLowerLvls(cS, idxIn)
'Debug.Print "in="; idxIn
'Debug.Print "out="; idxOut
'Debug.Print " "
    
    ' initialize
    Dim i As Integer
    Dim k As Integer
    Dim L As Integer
    i = idx
    k = 0
    L = UBound(S)
    
    ' maximum structure depth
    Dim maxS As Integer
    maxS = Application.WorksheetFunction.Max(S)
    
    ' S(i) = structure "S" with index "i"
    ' lowest level at S=maxS
    If S(idx) = maxS Then ' already on lowest level
        lastIdxOfLowerLvls = idx
    
    ' last element has no lower levels
    ElseIf i = L Then
        lastIdxOfLowerLvls = idx
    
    ' if next level is the same or higher --> that is not a sublevel
    ElseIf S(idx) >= S(i + 1) Then
        lastIdxOfLowerLvls = idx
        
    Else ' search forwards until all lower levels reached
         ' continue searching until:
            ' either same level or higher reached
            ' or array end reached
        While S(idx) < S(i + 1) And i <= L
        ' continue search until same or higher level reached (S=2 is above S=3).
        ' stop searching if e.g. S(idx)=2 and {S(i+1)=2 or S(i+1)=1}
            i = i + 1 ' Increment Counter
            If i = L Then
                GoTo arrEndReached
            End If
        Wend
        
arrEndReached:
        lastIdxOfLowerLvls = i ' search ended at targetIdx
        
    End If
    
End Function


' arrIn = array with data to be excluded
' excludeFirsts = array of indices to start excluding from
' excludeLasts = array of indices to end excluding from
' arrOut = cropped array
Private Function applyExclusionRanges(arrIn, excludeFirsts, excludeLasts)

' initialize
Dim i As Integer
Dim L As Integer
L = UBound(arrIn)

' ----- get exclusionFlags

' 0 = don't exclude
' 1 = apply exclude

' initialize exclusionFlags with 0
Dim exclusionFlags() As Integer
ReDim exclusionFlags(1 To L)
For i = 1 To L
    exclusionFlags(i) = 0
Next i

' initialize loop through exclusion ranges
Dim k As Integer
Dim L_exclusions As Integer
L_exclusions = UBound(excludeFirsts)

Dim excludeFirst As Integer
Dim excludeLast As Integer

' flip 0 to 1 if index in any exclusion range
For i = 1 To L
    For k = 1 To L_exclusions
        excludeFirst = excludeFirsts(k)
        excludeLast = excludeLasts(k)
        
        If i >= excludeFirst And i <= excludeLast Then ' index to be excluded
            exclusionFlags(i) = 1
        End If
    Next k
Next i


' ----- apply exclusionFlags

' set values to be excluded to empty
For i = 1 To L
    If exclusionFlags(i) = 1 Then ' flag the data
        If VarType(arrIn(1)) = 2 Then ' integer
            arrIn(i) = 0
        ElseIf VarType(arrIn(1)) = 5 Then ' double
            arrIn(i) = 0
        ElseIf VarType(arrIn(1)) = 8 Then ' string
            arrIn(i) = ""
        End If
    End If
Next i

' delete flagged elements and output cropped array
applyExclusionRanges = deleteEmpty(arrIn)

End Function



' delete empty and zero entries from array
Private Function deleteEmpty(arrIn)

ReDim arrOut(LBound(arrIn) To UBound(arrIn))
For i = 1 To UBound(arrIn)
    If arrIn(i) <> "" And arrIn(i) <> 0 Then ' there is a value to be stored
        j = j + 1
        arrOut(j) = arrIn(i)
    End If
Next i

If IsEmpty(j) = True Then ' no item deleted
    arrOut = arrIn
Else
    ReDim Preserve arrOut(LBound(arrIn) To j)
End If

deleteEmpty = arrOut

End Function



' delete empty and zero entries from array
Private Function deleteEmptyInt(ByVal arrIn)

Dim j As Integer
j = 0

Dim arrOut() As Integer
ReDim arrOut(1 To UBound(arrIn))

For i = 1 To UBound(arrIn)
    If arrIn(i) <> 0 Then  ' there is a value to be stored
        j = j + 1
        arrOut(j) = arrIn(i)
    End If
Next i

If j = 0 Then ' no item deleted
    arrOut = arrIn
Else
    ReDim Preserve arrOut(LBound(arrIn) To j)
End If

deleteEmptyInt = arrOut

End Function



' delete empty and zero entries from array
Private Function deleteEmptyStr(ByVal arrIn)

Dim j As Integer
j = 0

Dim arrOut() As String
ReDim arrOut(1 To UBound(arrIn))

For i = 1 To UBound(arrIn)
    If arrIn(i) <> "" Then  ' there is a value to be stored
        j = j + 1
        arrOut(j) = arrIn(i)
    End If
Next i

If j = 0 Then ' no item deleted
    arrOut = arrIn
Else
    ReDim Preserve arrOut(LBound(arrIn) To j)
End If

deleteEmptyStr = arrOut

End Function


' remove non-unique entries from array
' need sorted input for getUnique
Private Function getUnique(arrIn As Variant)
    ' sort input data to ensure unique by comparing with next
    Call QuickSort(arrIn, 1, UBound(arrIn))
   
    ' initialize
    Dim arrOut() As Variant
    ReDim arrOut(1 To UBound(arrIn))
    arrOut(1) = arrIn(1)
    
    Dim k As Integer
    k = 1
    
    For i = 2 To UBound(arrIn)
        If arrIn(i) = arrIn(i - 1) Then
            ' do nothing
        Else
            k = k + 1
            arrOut(k) = arrIn(i)
        End If
    Next i
    
    getUnique = deleteEmpty(arrOut)
    
End Function


' remove non-unique entries from array
' need sorted input for getUnique
Private Function getUniqueInt(ByVal arrIn)
    ' sort input data to ensure unique by comparing with next
    Call QuickSort(arrIn, 1, UBound(arrIn))
    
    ' initialize
    Dim arrOut() As Integer
    ReDim arrOut(1 To UBound(arrIn))
    arrOut(1) = arrIn(1)
    
    Dim k As Integer
    k = 1
    
    For i = 2 To UBound(arrIn)
        If arrIn(i) = arrIn(i - 1) Then
            ' do nothing
        Else
            k = k + 1
            arrOut(k) = arrIn(i)
        End If
    Next i
    
    getUniqueInt = deleteEmptyInt(arrOut)
    
End Function


' remove non-unique entries from array
' need sorted input for getUnique
Private Function getUniqueStr(ByVal arrIn)
    ' sort input data to ensure unique by comparing with next
    Call QuickSort(arrIn, 1, UBound(arrIn))
    
    ' initialize
    Dim arrOut() As String
    ReDim arrOut(1 To UBound(arrIn))
    arrOut(1) = arrIn(1)
    
    Dim k As Integer
    k = 1
        
    For i = 2 To UBound(arrIn)
        If arrIn(i) = arrIn(i - 1) Then
            ' do nothing
        Else
            k = k + 1
            arrOut(k) = arrIn(i)
        End If
    Next i
    
    getUniqueStr = deleteEmptyStr(arrOut)
    
End Function

' remove non-unique entries from array --> get corresponding titles
' need pre-sorted input for getUniqueT
Private Function getUniqueT(arrIn As Variant, arrInT As Variant)
    
    Dim arrOut() As Variant
    ReDim arrOut(1 To UBound(arrInT))
    
    Dim k As Integer
    k = 1
    
    For i = 1 To UBound(arrIn)
        If i = 1 Then
            arrOut(k) = arrInT(i) ' store title
        ElseIf StrComp(arrIn(i), arrIn(i - 1)) = 0 Then
            ' do nothing since not unique
        Else
            k = k + 1
            arrOut(k) = arrInT(i) ' store title
        End If
    Next i
    
    getUniqueT = deleteEmpty(arrOut)
    
End Function


' QuickSort array
Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub


' delete spaces from input string array
Private Function delSpaces(arrIn As Variant)
    
    Dim L As Integer
    Dim arrOut() As Variant
    
    ' check if individual string supplied
    If VarType(arrIn) = 8 Then ' string if only single entry, otherwise variant array of strings
        L = 1
        ReDim arrOut(1 To 1)
        
        arrOut(1) = Replace(arrIn, " ", "")
    Else
        L = UBound(arrIn)
        ReDim arrOut(1 To L)
        
        Dim i As Integer
        For i = 1 To L
            arrOut(i) = Replace(arrIn(i), " ", "")
        Next i
    End If
    
    delSpaces = arrOut

End Function


' get number of input rows
Private Function lastrow(SheetName As String)
    lastrow = Sheets(SheetName).Range("B" & Rows.Count).End(xlUp).Row
End Function


' read user-defined number circles to be excluded from further analysis
Private Function readExcludeStarts(SheetName As String)

    ' initialize excludes
    Dim firstExclude As Integer
    Dim lastExclude As Integer
    firstExclude = 22 ' ignore entries listet starting from "A(firstExclude+1)"
    lastExclude = 1000

    Dim excludes() As String
    ReDim excludes(1 To lastExclude - firstExclude)
    
    Dim L As Integer
    L = UBound(excludes)
    
    ' read excludes set by user
    For i = 1 To L
        If i = 1 And IsEmpty(Sheets(SheetName).Range("A" & firstExclude + i).Value) Then ' no exclusion at all
            ReDim excludes(1 To 1)
            excludes(1) = ""
            readExcludeStarts = ""
            GoTo noExclusionsAtAll
        ElseIf IsEmpty(Sheets(SheetName).Range("A" & firstExclude + i).Value) Then
            GoTo noMoreExcludes ' no more excludes
        Else
            excludes(i) = Sheets(SheetName).Range("A" & firstExclude + i).Value
        End If
    Next i
    
noMoreExcludes:
    
    ' remove empty entries
    readExcludeStarts = deleteEmpty(excludes)
    
noExclusionsAtAll:
    
End Function


' read input data
Private Function readData(ByVal SheetName As String, ByVal first As Integer, ByVal last As Integer, ByVal tgtCol As String) As Variant
    
    ' initialize
    Dim i As Integer
    Dim L As Integer
    L = last - first + 1
    Dim outputData() As Variant
    ReDim outputData(1 To L)
    Dim tmpRead As Variant
    tmpRead = Sheets(SheetName).Range(tgtCol & first & ":" & tgtCol & last).Value
    
    ' ensure there are no empty amounts in CAD raw data
    ' sheetNameCAD, firstCAD, lastCAD, "C"
    If StrComp(Right(SheetName, 3), "CAD") = 0 And StrComp(tgtCol, "C") = 0 Then
        For i = 1 To L
            If IsEmpty(tmpRead(i, 1)) Then
                tmpRead(i, 1) = 1
            End If
        Next i
    End If
    
    ' input data can have 1 or 2 dimensions
    Dim dimNum As Integer
    dimNum = NumberOfDimensions(tmpRead)
    
    ' ensure 1D-column output
    If dimNum = 1 Then
        For i = 1 To L
            outputData(i) = tmpRead(i)
        Next i
    ElseIf dimNum = 2 Then
        For i = 1 To L
            outputData(i) = tmpRead(i, 1)
        Next i
    Else
        Debug.Print "dimData > 2"
    End If
    
    readData = outputData
    
End Function


' get number of dimensions in array
Private Function NumberOfDimensions(ByVal vArray As Variant) As Long

Dim dimNum As Long
On Error GoTo FinalDimension

For dimNum = 1 To 60000
    ErrorCheck = LBound(vArray, dimNum)
Next

FinalDimension:
    NumberOfDimensions = dimNum - 1

End Function


' Stufen, get levels
' SAP starting from 1, CAD starting from 0 --> compensate by "Typ"
Private Function Stufen(Struktur As Variant, Typ As Integer) As Variant

    Dim first As Integer
    Dim last As Integer
    Dim Str As String
      
    first = LBound(Struktur)
    last = UBound(Struktur)
    Dim tmp() As Variant
    ReDim tmp(first To last)
    
    For i = first To last
        Str = Struktur(i) 'Str
        ' CAD structure contains .
        ' SAP structure can contain , and .
        tmp(i) = CountChrInString(Str, ",") + CountChrInString(Str, ".") + Typ
    Next i
    
    Stufen = tmp
End Function


' count character in string
Private Function CountChrInString(Expression As String, Character As String) As Long
    Dim iResult As Long
    Dim sParts() As String

    sParts = Split(Expression, Character)

    iResult = UBound(sParts, 1)

    If (iResult = -1) Then
    iResult = 0
    End If

    CountChrInString = iResult
    
End Function



' remove item and resize array
Public Sub ArrayRemoveItem(ItemArray As Variant, ByVal ItemElement As Long)
'PURPOSE:       Remove an item from an array, then
'               resize the array
'PARAMETERS:    ItemArray: Array, passed by reference, with
'               item to be removed.  Array must not be fixed

'               ItemElement: Element to Remove
'                    'EXAMPLE:
'           Dim iCtr As Integer
'           Dim sTest() As String
'           ReDim sTest(2) As String
'           sTest(0) = "Hello"
'           sTest(1) = "World"
'           sTest(2) = "!"
'
''           MsgBox sTest(1)
'
'           ArrayRemoveItem sTest, 1
'           For iCtr = 0 To UBound(sTest)
'               Debug.Print sTest(iCtr)
'           Next
'
''           MsgBox sTest(1)
'
''           Prints
''           "Hello"
''           "!"
''           To the Debug Window --> Ctrl+G
Dim lCtr As Long
Dim lTop As Long
Dim lBottom As Long

If Not IsArray(ItemArray) Then
    Err.Raise 13, , "Type Mismatch"
    Exit Sub
End If

lTop = UBound(ItemArray)
lBottom = LBound(ItemArray)

If ItemElement < lBottom Or ItemElement > lTop Then
    Err.Raise 9, , "Subscript out of Range"
    Exit Sub
End If

For lCtr = ItemElement To lTop - 1
    ItemArray(lCtr) = ItemArray(lCtr + 1)
Next
On Error GoTo ErrorHandler:

ReDim Preserve ItemArray(lBottom To lTop - 1)

Exit Sub
ErrorHandler:
  'An error will occur if array is fixed
    Err.Raise Err.Number, , _
       "You must pass a resizable array to this function"
End Sub

