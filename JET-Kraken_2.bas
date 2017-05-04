Attribute VB_Name = "Modul1"
Function nameToHeader(nameOf As String) As String

Select Case nameOf
    Case "BENUTZER_NAME"
        nameToHeader = "02_Sum_Fehlender_Benutzername"
    Case "BELEG_TEXT"
        nameToHeader = "02_Sum_Fehlender_Belegtext"
    Case "BENUTZER"
        nameToHeader = "02_Sum_Fehlender_Benutzer"
    Case "BETRAG"
        nameToHeader = "02_Sum_Fehlender_Betrag"
    Case "BELEG_TYP"
        nameToHeader = "02_Sum_Fehlender_Belegtyp"
    Case "ZEIT_ERFASST"
        nameToHeader = "02_Sum_Fehlende_Erfasszeit"
    Case "DATUM_ERFASST"
        nameToHeader = "02_Sum_Fehlender_Erfassdatum"
End Select

End Function

Function nameToPath(nameOf As String) As String

Select Case nameOf
    Case "BENUTZER_NAME"
        nameToPath = "noname.xlsx"
    Case "BELEG_TEXT"
        nameToPath = "notxt.xlsx"
    Case "BENUTZER"
        nameToPath = "nouser.xlsx"
    Case "BETRAG"
        nameToPath = "noamount.xlsx"
    Case "BELEG_TYP"
        nameToPath = "notype.xlsx"
    Case "ZEIT_ERFASST"
        nameToPath = "notime.xlsx"
    Case "DATUM_ERFASTT"
        nameToPath = "nodate.xlsx"
End Select

End Function

Sub importTableOnName(nameOf As String, headerOf As String, folderPath As String, startCell As String, mainWorkbook As Workbook)
'copy WS and rename it according to test
Worksheets("02_Sum_fehlende_Felder").Copy After:=Worksheets(2)


Worksheets(3).name = headerOf
'get path of folder there test file is
filePath = Left(folderPath, Len(folderPath) - 12) + nameToPath(nameOf)
'open test file
Workbooks.Open (filePath)
Set sourceWorkbook = ActiveWorkbook
Set objTable = ActiveSheet.ListObjects.Add(xlSrcRange, ActiveSheet.UsedRange, , xlYes)

sourceWorkbook.Worksheets(1).UsedRange.Copy
mainWorkbook.Worksheets(headerOf).Range(startCell).PasteSpecial
Application.CutCopyMode = False
sourceWorkbook.Close SaveChanges:=False

End Sub

Sub renameHeaders(ws As Worksheet)
Dim s As Integer
Dim headcol As String
s = 1
Do While ws.Cells(25, s).Value <> ""
    headcol = ws.Cells(25, s).Value
    If Left(headcol, 2) = "Z_" Then
        ws.Cells(25, s).Value = Mid(ws.Cells(25, s).Value, 3, Len(ws.Cells(25, s).Value) - 2)
        s = s + 1
    Else
        Select Case headcol
        Case "BETRAG"
            Range(ws.Cells(26, 9).Address, Range(ws.Cells(26, 9).Address).End(xlDown)).NumberFormat = "#,###.##0;[Red]-#,###.##0"
        Case "BENUTZER"
            Range(ws.Cells(26, s).Address, Range(ws.Cells(26, s).Address).End(xlDown)).Interior.ColorIndex = 44
        Case "BELEG_TEXT"
            Range(ws.Cells(26, s).Address, Range(ws.Cells(26, s).Address).End(xlDown)).Interior.ColorIndex = 22
        End Select
        s = s + 1
    End If
Loop
End Sub

Function getStartCell(ws As Worksheet) As Integer
Dim i As Integer
Dim first As Integer
ws.Rows("20:21").Delete
Do While ws.Cells(20 + i, 2).Value <> ""
    i = i + 1
Loop
first = 20 + i
ws.Rows(first & ":28").Delete

getStartCell = 21 + i
End Function


Sub eatdis()
'get path of folder there test file is
Dim myPath As String
Dim folderPath As String
Dim filePath As String
folderPath = Application.ActiveWorkbook.Path
'delete "Auswertung" at the end of the filepath (like as one level backwards in the folder)
folderPath = Left(folderPath, Len(folderPath) - 12)
Dim rIndex() As Long
Dim maxArray() As Long
Dim minArray() As Long
Dim sumArray() As Long
Dim bufferArray() As Long
Dim nameArray() As String
Dim bestArray() As String
Dim bestNameArray() As String
Dim bufferIndex As Long
Dim maxBuffer As Long
Dim minBuffer As Long
Dim anzahlBu As Long
Dim median As Long
Dim stufe As Integer
Dim flag As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim s As Integer
Dim t As Integer
Dim rowHeight As Integer
Dim emptyuser As Integer
Dim tempTable As Worksheets
Dim summary As Worksheet
Dim nouser As Worksheet
Dim notxt As Worksheet
Dim noname As Worksheet
Dim konto As Worksheet
Dim user As Worksheet
Dim beleg As Worksheet
Dim periode As Worksheet
Dim gkto As Worksheet
Dim gkto1 As Worksheet
Dim gkto2 As Worksheet
Dim gkto3 As Worksheet
Dim personen As Worksheet
Dim erdat As Worksheet
Dim benford1 As Worksheet
Dim benford2 As Worksheet
Dim butxt As Worksheet
Dim wknd As Worksheet
Dim time As Worksheet
Dim kopfz As Worksheet
Dim wsname As String
Dim cellname As String
Dim bufferName As String
Dim bufferValue As String
Dim aSatz21 As String
Dim aSatz22 As String
Dim aSatz As String
Dim startCell As String
Dim headerOf As String



Dim sourceWorkbook As Workbook
Dim mainWorkbook As Workbook
Dim objTable As ListObject

Set mainWorkbook = ActiveWorkbook
folderPath = Application.ActiveWorkbook.Path
Set summary = Worksheets(1)
'anzahl der buchungen
anzahlBu = Worksheets(7).Cells(50, 2).Value



'Worksheets(7).Cells(51, 2).Value = anzahlBu

'--------------------------------------------------------------------------------------------
'test 2 Fehlende Felder
'create list of worksheets
i = 0
startCell = Replace("A" + Str(getStartCell(Worksheets(3))), " ", "")
aSatz22 = "Darüber hinaus waren bis auf die nachfolgenden Informationen alle Beleginformationen im analysierten Buchungsstoff enthalten:" + vbCrLf
Do While Worksheets(3).Cells(20 + i, 2).Value <> ""
    If Worksheets(3).Cells(20 + i, 2).Value = anzahlBu Then
        Worksheets(3).Cells(20 + i, 5).Value = "Informationen sind im Datenabzug nicht vorhanden"
        aSatz22 = aSatz22 + "- " + Worksheets(3).Cells(20 + i, 1).Value + vbCrLf
        Else
        cellname = Worksheets(3).Cells(20 + i, 1).Value
        aSatz21 = aSatz21 + "Insgesamt " + CStr(Worksheets(3).Cells(20 + i, 2).Value) + " Buchungen (  Buchungsbelege) identifiziert, für die kein " + cellname + " vorhanden ist." + vbCrLf + vbCrLf
        
        Select Case cellname
            Case "BENUTZER_NAME"
                
                'later
            Case "BELEG_TEXT"
                'startCell = "A25"
                headerOf = nameToHeader(cellname)
                importTableOnName cellname, headerOf, folderPath, startCell, mainWorkbook
                Set notxt = Worksheets(headerOf)
                renameHeaders notxt
                
                
            Case "BENUTZER"
                'startCell = "A25"
                headerOf = nameToHeader(cellname)
                importTableOnName cellname, headerOf, folderPath, startCell, mainWorkbook
                Set nouser = Worksheets(headerOf)
                renameHeaders nouser
        End Select
    End If
   i = i + 1
Loop
'define all worksheets
For i = 3 To ThisWorkbook.Worksheets.Count
    wsname = Worksheets(i).name
    Select Case wsname
        Case "02_Sum_fehlender_Benutzer"
            Set nouser = Worksheets(i)
        Case "02_Sum_fehlender_Belegtext"
            Set notxt = Worksheets(i)
        Case "02_Sum_fehlender_Name"
            Set noname = Worksheets(i)
        Case "03_Sum_Konto"
            Set konto = Worksheets(i)
        Case "04_Sum_Benutzer"
            Set user = Worksheets(i)
        Case "05_Sum_Benutzer_Belegtyp"
            Set beleg = Worksheets(i)
        Case "06_Sum_Buchungsmonat"
            Set periode = Worksheets(i)
        Case "07_Gegenkontoanalyse_Auswahl"
            Set gkto = Worksheets(i)
        Case "07_Gegenkontoanalyse_Ergebnis1"
            Set gkto1 = Worksheets(i)
        Case "07_Gegenkontoanalyse_Ergebnis2"
            Set gkto2 = Worksheets(i)
        Case "07_Gegenkontoanalyse_Ergebnis3"
            Set gkto3 = Worksheets(i)
        Case "08_Nahestehende_Personen"
            Set personen = Worksheets(i)
        Case "09_Sum_Bumonat_Erstdat"
            Set erdat = Worksheets(i)
        Case "10a_Benford_eine_Ziffer"
            Set benford1 = Worksheets(i)
        Case "10b_Benford_zwei_Ziffern"
            Set benford2 = Worksheets(i)
        Case "11_UngewÃ¶hnlich_Buchtext"
            Set butxt = Worksheets(i)
        Case "12_Pivot_Buchungen_WE_FT"
            Set wknd = Worksheets(i)
        Case "13_Pivot_Ungew_Zeiten"
            Set time = Worksheets(i)
        Case "Kopfzeile"
            Set kopfz = Worksheets(i)
    End Select
Next i
'Fill text for test 2
aSatz = Replace(Replace(Replace(Replace(Replace(aSatz21 + aSatz22, "BELEG_TYP", "Belegtyp"), "BENUTZER_NAME", "Benutzername"), "ZEIT_ERFASST", "Erfassungszeit"), "BENUTZER", "Benutzer"), "BELEG_TEXT", "Belegtext")
summary.Cells(30, 5).Value = aSatz

'----------------------------------------------------------------------------------------------
'test 3 Konto-Zusammenfassung
'get account count
i = 0
Do While Len(konto.Cells(i + 23, 1).Value) <> 0
    i = i + 1
Loop
aSatz = "Anzahl der bebuchten Konten: " + CStr(i) + vbCrLf + vbCrLf + "Gemessen an der Anzahl der Buchungszeilen wurden im Wesentlichen die folgenden Konten bebucht:"
'fill array of index with amounts (ANZAHL)
ReDim rIndex(i)
For j = 0 To i - 1
   rIndex(j) = konto.Cells(23 + j, 3).Value
Next j
'find 5 max amounts (ANZAHL)
For k = 1 To 5
    maxBuffer = 0
    For j = 0 To i - 1
        If rIndex(j) > maxBuffer Then
            maxBuffer = rIndex(j)
            bufferIndex = j
        End If
    Next j
    rIndex(bufferIndex) = 0
    aSatz = aSatz + vbCrLf + " -" + CStr(konto.Cells(bufferIndex + 23, 1).Value) + " " + CStr(konto.Cells(bufferIndex + 23, 2).Value)
Next k
aSatz = aSatz + vbCrLf + vbCrLf + "Bezogen auf den absoluten Buchungssaldo wurden zusätzlich die folgenden Konten bebucht:"
'fill array of index with amounts (Abs.SALDO)
For j = 0 To i - 1
   rIndex(j) = konto.Cells(23 + j, 7).Value
Next j
'find ca. 5 amx amounts (Abs.SALDO)
For k = 1 To 5
    maxBuffer = 0
    For j = 0 To i - 1
        If rIndex(j) > maxBuffer Then
            maxBuffer = rIndex(j)
            bufferIndex = j
        End If
    Next j
    rIndex(bufferIndex) = 0
    aSatz = aSatz + vbCrLf + " -" + CStr(konto.Cells(bufferIndex + 23, 1).Value) + " " + CStr(konto.Cells(bufferIndex + 23, 2).Value) + " (" + CStr(konto.Cells(bufferIndex + 23, 7).Value) + "€ )"
Next k
aSatz = aSatz + vbCrLf + vbCrLf + "Bezogen auf den absoluten Buchungssaldo wurden zusätzlich die folgenden Konten vereinzelt bebucht:"
ReDim maxArray(5)
ReDim bufferArray(5)
Do While stufe < 7 Or flag = 1
    For j = 0 To i - 1
       rIndex(j) = konto.Cells(23 + j, 7).Value
    Next j
    flag = 1
    
    For k = 1 To 5
        maxArray(k - 1) = 0
        For j = 0 To i - 1
            If rIndex(j) > maxArray(k - 1) And konto.Cells(j + 23, 3).Value < stufe Then
                maxArray(k - 1) = rIndex(j)
                bufferArray(k - 1) = j
            End If
        Next j
        rIndex(bufferArray(k - 1)) = 0
        If maxArray(k - 1) < 1000000 Then
            flag = 0
        End If
    Next k
    stufe = stufe + 1
Loop
For j = 0 To 4
    aSatz = aSatz + vbCrLf + " -" + CStr(konto.Cells(bufferArray(j) + 23, 1).Value) + " " + CStr(konto.Cells(bufferArray(j) + 23, 2).Value) + " (" + CStr(konto.Cells(bufferArray(j) + 23, 7).Value) + "€ )"
Next j
'fill text of test 2 in summary
summary.Cells(31, 5).Value = aSatz

'------------------------------------------------------------------------
'test 4 User-Summary
'get list of user counts
i = 19
j = 0

Do While Len(user.Cells(i, 3).Value) <> 0
    j = j + 1
    i = i + 1
Loop
j = j - 1
k = 0
ReDim maxArray(j)
ReDim nameArray(j)
emptyuser = 0
For i = 19 To 18 + j
    
    If user.Cells(i, 1).Value <> "" Then
        maxArray(k) = user.Cells(i, 3).Value
        nameArray(k) = user.Cells(i, 1).Value
        k = k + 1
        Else
        emptyuser = i
    End If
Next i
aSatz = "Anzahl der Buchungsersteller: " + CStr(k)
'get users with largest and smallest posting counts, set the list of usernames and their counts
maxBuffer = maxArray(0)
minBuffer = maxArray(0)
For i = 1 To k - 1
    If maxArray(i) > maxBuffer Then
        maxBuffer = maxArray(i)
    End If
Next i
For i = 1 To k - 1
    If maxArray(i) < minBuffer Then
        minBuffer = maxArray(i)
    End If
Next i
ReDim bestArray(k)
ReDim worstArray(k)
If emptyuser <> 0 Then
    If user.Cells(emptyuser, 3).Value > maxBuffer Then
        aSatz = aSatz + vbCrLf + vbCrLf + "Die Buchungen wurden im Wesentlichen ohne Benutzerangabe erfasst." + vbCrLf + vbCrLf + "Soweit die Information über den Buchungsersteller vorhanden ist, wurden die Buchungen zum größten Teil von de"
        Else
        aSatz = aSatz + vbCrLf + vbCrLf + "Soweit die Information über den Buchungsersteller vorhanden ist, wurden die Buchungen zum größten Teil von de"
    End If
Else
    aSatz = aSatz + vbCrLf + vbCrLf + "Die Buchungen wurden zum größten Teil von de"
End If
s = 0
t = 0
'create a list of users with largest and smallest count of postings
For i = 0 To k - 1
    If maxArray(i) > maxBuffer * 0.6 Then
        bestArray(s) = nameArray(i)
        s = s + 1
    Else
        If maxArray(i) < minBuffer * 2 Or maxArray(i) < minBuffer + 100 Then
            worstArray(t) = nameArray(i)
            t = t + 1
        End If
    End If
Next i
'creat text for largest list
If s > 1 Then
    aSatz = aSatz + "n Benutzern " + """" + bestArray(0) + """"
    For i = 1 To s - 1
        If i = s - 1 Then
            aSatz = aSatz + " und " + """" + bestArray(i) + """"
        Else
            aSatz = aSatz + ", " + """" + bestArray(i) + """"
        End If
    Next i
    aSatz = aSatz + " erstellt."
Else
    aSatz = aSatz + "m Benutzer " + """" + bestArray(0) + """" + " erstellt."
End If
'create list for smallest if exist
If t <> 0 Then
    aSatz = aSatz + vbCrLf + vbCrLf + "Dagegen wurden durch d"
    If t = 1 Then
        aSatz = aSatz + "en Benutzer " + """" + worstArray(0) + """"
    Else
        aSatz = aSatz + "ie Benutzer " + """" + worstArray(0) + """"
        For i = 1 To t - 1
            If i = t - 1 Then
                aSatz = aSatz + " und " + """" + worstArray(i) + """"
            Else
                aSatz = aSatz + ", " + """" + worstArray(i) + """"
            End If
        Next i
    End If
    aSatz = aSatz + " nur Buchungen in geringem Umfang vorgenommen."
End If
'entry text to test 4 in summary
summary.Cells(32, 5).Value = aSatz
'---------------------------------------------------------------------------------------
'test 5 DocumentType-Summary
i = 19
j = 0
'get count rows in result
Do While Len(beleg.Cells(i, 4).Value) <> 0
    j = j + 1
    i = i + 1
Loop
j = j - 1
'set the list of document types, not unique
ReDim nameArray(j)
ReDim sumArray(j)
For i = 19 To 18 + j
    nameArray(i - 19) = beleg.Cells(i, 3).Value
Next i
'sum counts for every unique document type
t = 0
ReDim bestNameArray(j)

For k = 0 To j - 1
    If nameArray(k) <> "wasted" Then
        bufferName = nameArray(k)
        
        nameArray(k) = "wasted"
        
        bestNameArray(t) = beleg.Cells(19 + k, 3).Value
        
        sumArray(t) = beleg.Cells(19 + k, 4).Value
        
        For i = 19 + k To 18 + j
            
            If nameArray(i - 19) = bufferName Then
                sumArray(t) = sumArray(t) + beleg.Cells(i, 4).Value
                nameArray(i - 19) = "wasted"
            End If
            'beleg.Cells(14 + i - t, t + 1).Value = t
        Next i
        
        t = t + 1
    End If
Next k

'get max count of type
maxBuffer = sumArray(0)
For i = 1 To t - 1
    If sumArray(i) > maxBuffer Then
        maxBuffer = sumArray(i)
    End If
Next i
'get first best counts of types
ReDim bestArray(t)
s = 0
aSatz = "Anzahl der verwendeten Belegtypen: " + CStr(t) + vbCrLf + vbCrLf
For i = 0 To t - 1
    If sumArray(i) > maxBuffer * 0.6 And sumArray(i) + 15000 > maxBuffer Then
        bestArray(s) = sumArray(i)
        bestNameArray(s) = bestNameArray(i)
        s = s + 1
    End If
Next i
aSatz21 = ""
If t = 1 Then
    aSatz = aSatz + "Es wurde nur ein Belegtyp " + """" + bestNameArray(0) + """" + " verwendet."
Else
    aSatz21 = """" + bestNameArray(0) + """"
    For i = 1 To s - 1
        If i = s - 1 Then
            aSatz21 = aSatz21 + " und " + """" + bestNameArray(i) + """"
        Else
            aSatz21 = aSatz21 + ", " + """" + bestNameArray(i) + """"
        End If
    Next i
    If s = t Then
        aSatz = aSatz + "Die Belegtypen " + aSatz21 + " wurden relativ gleichmäßig verwendet."
    Else
        If s = 1 Then
            aSatz = aSatz + "Im Wesentlichen wurde der Belegtyp " + aSatz21 + " verwendet."
        Else
            aSatz = aSatz + "Im Wesentlichen wurden die Belegtypen " + aSatz21 + " verwendet."
        End If
    End If
End If
summary.Cells(33, 5).Value = aSatz
'----------------------------------------------------------------------------------------------------------
'test 6 period
'get year
Dim year As String
year = Mid(periode.Cells(38, 1).Value, 1, 4)
aSatz = "Erhöhtes Buchungsaufkommen in de"
ReDim bufferArray(12)
For i = 0 To 11
    bufferArray(i) = periode.Cells(38 + i, 2).Value
Next i

maxBuffer = WorksheetFunction.Max(bufferArray)
'bubblesort
For j = 11 To 0 Step -1
    For i = 0 To j - 1
        If bufferArray(i) > bufferArray(i + 1) Then
            bufferValue = bufferArray(i)
            bufferArray(i) = bufferArray(i + 1)
            bufferArray(i + 1) = bufferValue
        End If
    Next i
Next j
'get median
median = (bufferArray(5) + bufferArray(6)) * 0.5
For i = 0 To 11
    bufferArray(i) = periode.Cells(38 + i, 2).Value
Next i
'define period names
Dim months() As String
ReDim months(12)
months(0) = "Januar"
months(1) = "Februar"
months(2) = "März"
months(3) = "April"
months(4) = "Mai"
months(5) = "Juni"
months(6) = "Juli"
months(7) = "August"
months(8) = "September"
months(9) = "Oktober"
months(10) = "November"
months(11) = "Dezember"
t = 0
s = 0
'get indexes of high and low posting counts
ReDim maxArray(6)
ReDim minArray(6)
For i = 0 To 11
    If bufferArray(i) > median * 1.25 Then
        maxArray(t) = i
        t = t + 1
    Else
        If bufferArray(i) < median * 0.8 Then
            minArray(s) = i
            s = s + 1
        End If
    End If
Next i
'get list of high periods
If t > 0 Then
    If t = 1 Then
        aSatz = aSatz + "r Buchungsperiode " + months(maxArray(0)) + " " + year + "."
    Else
        aSatz = aSatz + "n Buchungsperioden "
        For i = 0 To t - 1
            If i = t - 1 Then
                aSatz = Left(aSatz, Len(aSatz) - 2) + " und " + months(maxArray(i)) + " " + year + "."
            Else
                aSatz = aSatz + months(maxArray(i)) + ", "
            End If
        Next i
    End If
End If
'get list of low periods
If s > 0 Then
    aSatz = aSatz + vbCrLf + vbCrLf + "Dagegen besteht ein verringertes Buchungsaufkommen i"
    If s = 1 Then
        aSatz = aSatz + "m Monat" + months(minArray(0)) + " " + year + "."
    Else
        aSatz = aSatz + "n den Monaten "
        For i = 0 To s - 1
            If i = s - 1 Then
                aSatz = Left(aSatz, Len(aSatz) - 2) + " und " + months(minArray(i)) + " " + year + "."
            Else
                aSatz = aSatz + months(minArray(i)) + ", "
            End If
        Next i
    End If
End If
'check if postings are relativ equal
If s = 0 And t = 0 Then
    aSatz = "Buchungsaufkommen im Jahr " + year + " ist über die Buchungsperioden gleichmäßig verteilt."
End If
'set text to summary
summary.Cells(34, 5).Value = aSatz
'------------------------------------------------------------------------------------------------------
'import fehlende Felder
'notxt, nouser etc.


End Sub



