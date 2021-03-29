Attribute VB_Name = "Module1"
Option Explicit

Global wbPath As String
Global tPath As String
Global oPath As String
Global currentYear As String
Global currentMonth As String
Global currentGender As String
Global currentAge As String
Global fileList As Variant

Global Word As Object
Global primaryFile As String
Global primaryFilePath As String

Sub Initialize()
    wbPath = ThisWorkbook.Path
    tPath = wbPath & "\Templates"
    oPath = wbPath & "\out"
    currentYear = Year(Date)
    currentMonth = MonthName(Month(Date))
    fileList = Array()
    primaryFile = "Birthday Letters For " & currentMonth & ".docx"
    primaryFilePath = oPath & "\" & primaryFile
End Sub

Sub CloseWordObjects() 'close all instances of word.
    Dim objWord As Object
    Dim blnHaveWorkObj As Boolean

    ' assume a Word object is there to be quit
    blnHaveWorkObj = True
    
    ' loop until no Word object available
    Do
        On Error Resume Next
        Set objWord = GetObject(, "Word.Application")
        objWord.DisplayAlerts = 0
        If objWord Is Nothing Then
            ' quit loop
            blnHaveWorkObj = False
        Else
            ' quit Word
            objWord.Quit
            ' clean up
            Set objWord = Nothing
        End If
    Loop Until Not blnHaveWorkObj
End Sub

Sub CreateEmptyWordDoc()
    Dim WordDoc3 As Object
    Dim Word3 As Object
    Set Word3 = CreateObject("Word.Application")
    Set WordDoc3 = Word3.Documents.Add()
    Word3.Visible = False
    WordDoc3.SaveAs2 primaryFilePath, 16
    Word3.Quit
End Sub

Sub SortDataByDay() 'sorting by birth day (not year).
    Dim Column As Range: Set Column = Application.Range("F:F")
    Column.Insert Shift:=xlShiftToRight, CopyOrigin:=xlFormatFromRightOrBelow

    Dim i, originalValue, toRemove, parsedValue
    Dim x: x = Range("G2")
    Dim k: k = Range("G2").CurrentRegion.Rows.Count - 1
    
    For i = 2 To k + 1
        originalValue = Range("G" & i)
        toRemove = Len(originalValue) - 5
        parsedValue = Left(originalValue, toRemove)
            Range("F" & i).Value = parsedValue
    Next i
    
    Range("A:O").Sort Key1:=Range("F1"), Order1:=xlAscending, Header:=xlYes
    Range("F:F").Delete
End Sub

Sub FormatGender(gender)
    If (gender = "M") Then
        currentGender = "Male"
    ElseIf (gender = "F") Then
        currentGender = "Female"
    Else
        currentGender = "Undefined"
    End If
End Sub

Sub FormatAge(age)
    Dim toRemove As String: toRemove = Right(age, 6)
    currentAge = Trim(Replace(age, toRemove, ""))
End Sub

Sub ListCurrentFiles()
    fileList = Array()
    Dim primaryFile As String: primaryFile = "Birthday Letters For " & currentMonth & ".docx"
    Dim toScan As String: toScan = Dir(oPath & "\*.docx")

    Do While Len(toScan) > 0
    Dim counter As Integer
        If Right(toScan, 4) = "docx" And toScan <> primaryFile Then
            counter = counter + 1
            ReDim Preserve fileList(UBound(fileList) + 1)
            fileList(UBound(fileList)) = toScan
            'Debug.Print fileList(counter - 1)
        End If
        toScan = Dir
    Loop
End Sub

Sub CheckPrimaryExists()
    fileList = Array()
    Dim toScan As String: toScan = Dir(oPath & "\" & primaryFile)
    Dim toDelete As String: toDelete = oPath & "\" & primaryFile

    If Len(toScan) = 0 Then
        CreateEmptyWordDoc
    Else
        Kill toDelete
        CreateEmptyWordDoc
    End If

End Sub

Sub ReplaceTextFromTemplates()
    Dim tName, tName2, tFilePath, fileName, fileName2
    Dim fName, lName, gender, bDay, age, aLine1, aLine2, cCity, pCode, birthYear, parsedDate, fmtDate, fmtDate2, currentName As String
    Dim i As Integer
    Dim k: k = 201
    Debug.Print k
        Initialize
        CloseWordObjects
        ListCurrentFiles
        
        Dim Word As Object
        Set Word = CreateObject("Word.Application")

        For i = 0 To k
            If i < 101 Then
                currentGender = "Male"
                currentAge = i
            Else
                currentGender = "Female"
                currentAge = i - 101
            End If
            
            tName = currentAge & " " & currentGender & ".dotx"
            tName2 = "Document" & i & ".docx"
            tFilePath = tPath & "\" & tName

            birthYear = Right(bDay, 4)
            parsedDate = Replace(bDay, birthYear, currentYear)
            fmtDate = Format(parsedDate, "mmmm dd, yyyy")
            fmtDate2 = Format(parsedDate, "mmmm dd") & " " & currentYear & " " & Format(Now(), "hh-mm-ss")

            fileName = oPath & "\Birthday Letters For " & currentMonth & ".docx"
            currentName = fName & " " & lName

                Dim WordDoc As Object
                Set WordDoc = Word.Documents.Open(tFilePath, AddToRecentFiles:=False, Visible:=True)
                With WordDoc.Content.Find
                    .Execute FindText:="<<Text To Replace>>", ReplaceWith:="<<Text To Replace With", Replace:=wdReplaceAll
                End With
                Word.DisplayAlerts = False
                Word.Documents(tFilePath).Close SaveChanges:=wdSaveChanges

                Debug.Print Now(), "Iteration: " & i, "Name: " & tName

         Next i
         MsgBox "Documents generated successfully."
         Word.Quit SaveChanges:=wdDoNotSaveChanges
         CloseWordObjects
End Sub

Sub main()
    Dim tName, tName2, tFilePath, fileName, fileName2
    Dim fName, lName, gender, bDay, age, aLine1, aLine2, cCity, pCode, birthYear, parsedDate, fmtDate, fmtDate2, currentName As String
    Dim i As Integer
    Dim k: k = Range("F1").CurrentRegion.Rows.Count
    Debug.Print k
    If k > 1 Then
        Initialize
        CloseWordObjects
        CheckPrimaryExists
        SortDataByDay
        ListCurrentFiles

        Set Word = CreateObject("Word.Application")

        progressBox.Show vbModeless
        progressBox.progressBar.Width = 1

        For i = 2 To k + 1
            fName = Range("C" & i)
            lName = Range("D" & i)
            gender = Range("E" & i)
            bDay = Range("F" & i)
            age = Range("G" & i)
            aLine1 = Range("J" & i)
            aLine2 = Range("K" & i)
            cCity = Range("L" & i)
            pCode = Range("M" & i)

            FormatAge (age)
            FormatGender (gender)

            tName = currentAge & " " & currentGender & ".dotx"
            tName2 = "Document" & i & ".docx"
            tFilePath = tPath & "\" & tName

            birthYear = Right(bDay, 4)
            parsedDate = Replace(bDay, birthYear, currentYear)
            fmtDate = Format(parsedDate, "mmmm dd, yyyy")
            fmtDate2 = Format(parsedDate, "mmmm dd") & " " & currentYear & " " & Format(Now(), "hh-mm-ss")

            fileName = oPath & "\Birthday Letters For " & currentMonth & ".docx"
            currentName = fName & " " & lName

            If gender = "M" Or gender = "F" Then

                Dim WordDoc As Object
                Set WordDoc = Word.Documents.Add(Template:=tFilePath, NewTemplate:=False, DocumentType:=0, Visible:=True)
                With WordDoc.Content.Find
                    .Execute FindText:="<<ClientFirstName>>", ReplaceWith:=fName, Replace:=wdReplaceAll
                    .Execute FindText:="<<ClientLastName>>", ReplaceWith:=lName, Replace:=wdReplaceAll
                    .Execute FindText:="<<Birthday>>", ReplaceWith:=fmtDate, Replace:=wdReplaceAll
                    .Execute FindText:="<<AddressLine1>>", ReplaceWith:=aLine1, Replace:=wdReplaceAll
                    .Execute FindText:="<<AddressLine2>>", ReplaceWith:=aLine2, Replace:=wdReplaceAll
                    .Execute FindText:="<<City>>", ReplaceWith:=cCity, Replace:=wdReplaceAll
                    .Execute FindText:="<<PostalCode>>", ReplaceWith:=pCode, Replace:=wdReplaceAll
                End With

                WordDoc.Sections(1).Range.Copy

                Dim WordDoc2 As Object
                Set WordDoc2 = Word.Documents.Open(primaryFilePath, AddToRecentFiles:=False, Visible:=True)


                Dim LastPage: LastPage = WordDoc2.Sections.Count
                WordDoc2.Sections(LastPage).Range.Paste
                If i < k Then
                    WordDoc2.Sections.Add
                End If
                WordDoc2.SaveAs2 primaryFilePath, 16

                Dim x: x = (i / k) * 100
                Dim y: y = (x * progressBox.staticBar.Width) / 100

                progressBox.progressBar.Width = y
                progressBox.progressText.Caption = currentName & " " & "(" & i - 1 & "/" & k - 1 & ")"

                Debug.Print Now(), "Iteration: " & i, "Name: " & currentName
            End If

         Next i
         MsgBox "Documents generated successfully."
         Unload progressBox
         Word.Quit SaveChanges:=wdDoNotSaveChanges
         CloseWordObjects
    Else
        MsgBox "There is no data to process."
    End If
End Sub
