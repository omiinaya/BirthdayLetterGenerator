Attribute VB_Name = "Module1"
Option Explicit

 Public Declare Function FindWindow Lib "user32" _
                Alias "FindWindowA" _
               (ByVal lpClassName As String, _
                ByVal lpWindowName As String) As Long


    Public Declare Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long


    Public Declare Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long


    Public Declare Function DrawMenuBar Lib "user32" _
               (ByVal hWnd As Long) As Long

Global wbPath As String
Global tPath As String
Global oPath As String
Global currentYear As String
Global currentMonth As String
Global currentGender As String
Global currentAge As String
Global fileList As Variant
Global primaryFile As String
Global primaryFilePath As String
Global selectedFileName
Global selectedFilePath

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
    Dim Column As Range: Set Column = Workbooks(selectedFileName).Worksheets(1).Range("F:F")
    Column.Insert Shift:=xlShiftToRight, CopyOrigin:=xlFormatFromRightOrBelow

    Dim i, originalValue, toRemove, parsedValue
    Dim x: x = Workbooks(selectedFileName).Worksheets(1).Range("G2")
    Dim k: k = Workbooks(selectedFileName).Worksheets(1).Range("G2").CurrentRegion.Rows.Count - 1
    
    For i = 2 To k + 1
        originalValue = Workbooks(selectedFileName).Worksheets(1).Range("G" & i)
        toRemove = Len(originalValue) - 5
        parsedValue = Left(originalValue, toRemove)
            Workbooks(selectedFileName).Worksheets(1).Range("F" & i).Value = parsedValue
    Next i
    
    Workbooks(selectedFileName).Worksheets(1).Range("A:O").Sort Key1:=Range("F1"), Order1:=xlAscending, Header:=xlYes
    Workbooks(selectedFileName).Worksheets(1).Range("F:F").Delete
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

Sub FormatAge(Age)
    Dim toRemove As String: toRemove = Right(Age, 6)
    currentAge = CStr(CInt(Trim(Replace(Age, toRemove, ""))) + 1)
    Debug.Print currentAge
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
    Dim i As Integer
    Dim k: k = 201
    'Debug.Print k
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

                Dim WordDoc As Object
                Set WordDoc = Word.Documents.Open(tFilePath, AddToRecentFiles:=False, Visible:=False)
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

Sub HideBar(frm As Object)

Dim Style As Long, Menu As Long, hWndForm As Long
hWndForm = FindWindow("ThunderDFrame", frm.Caption)
Style = GetWindowLong(hWndForm, &HFFF0)
Style = Style And Not &HC00000
SetWindowLong hWndForm, &HFFF0, Style
DrawMenuBar hWndForm

End Sub


Sub main()
    
    Initialize
    CloseWordObjects
    CheckPrimaryExists
    ListCurrentFiles
    
    
    Dim fName, lName, gender, bDay, Age, aLine1, aLine2, cCity, pCode, tName, tName2, tFilePath, tFileName, birthYear, parsedDate, fileName, fmtDate, fmtDate2, currentName
    Dim i, j, k, n, o, z
    
    Dim Word As Object
    Dim Excel As Object
    
    Set Word = CreateObject("Word.Application")
    Set Excel = CreateObject("Excel.Application")
    
    
    Excel.Visible = False
    Word.Visible = False
    
    'Word.ScreenUpdating = False
    
    With Excel.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb", 1
        .Show
        selectedFilePath = .SelectedItems(1)
        selectedFileName = Dir(selectedFilePath)
    End With
    Workbooks.Open selectedFilePath
    
    SortDataByDay
    
    progressBox.Show vbModeless
    progressBox.progressBar.Width = 1
    
    k = Workbooks(selectedFileName).Worksheets(1).Range("F1").CurrentRegion.Rows.Count
    n = Word.Documents.Count
    j = 1
    For i = 2 To k
        j = j + 1
        fName = Workbooks(selectedFileName).Worksheets(1).Range("C" & i)
        lName = Workbooks(selectedFileName).Worksheets(1).Range("D" & i)
        gender = Workbooks(selectedFileName).Worksheets(1).Range("E" & i)
        bDay = Workbooks(selectedFileName).Worksheets(1).Range("F" & i)
        Age = Workbooks(selectedFileName).Worksheets(1).Range("G" & i)
        aLine1 = Workbooks(selectedFileName).Worksheets(1).Range("J" & i)
        aLine2 = Workbooks(selectedFileName).Worksheets(1).Range("K" & i)
        cCity = Workbooks(selectedFileName).Worksheets(1).Range("L" & i)
        pCode = Workbooks(selectedFileName).Worksheets(1).Range("M" & i)
        
        FormatAge (Age)
        FormatGender (gender)

        tName = currentAge & " " & currentGender & ".dotx"
        tFilePath = tPath & "\" & tName
        tFileName = "Document" & j - 1

        birthYear = Right(bDay, 4)
        parsedDate = Replace(bDay, birthYear, currentYear)
        fmtDate = Format(parsedDate, "mmmm dd, yyyy")
        fmtDate2 = Format(parsedDate, "mmmm dd") & " " & currentYear & " " & Format(Now(), "hh-mm-ss")

        fileName = oPath & "\Birthday Letters For " & currentMonth & ".docx"
        currentName = fName & " " & lName
        
        If gender = "M" Or gender = "F" Then
                Dim WordDoc As Object
                'Set WordDoc = Nothing
                Set WordDoc = Word.Documents.Add(Template:=tFilePath, NewTemplate:=False, DocumentType:=0, Visible:=False)
                'Word.Visible = False
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
                Set WordDoc2 = Word.Documents.Open(primaryFilePath, AddToRecentFiles:=False, Visible:=False)

                Dim LastPage: LastPage = WordDoc2.Sections.Count
                WordDoc2.Sections(LastPage).Range.Paste
                
                Word.Documents(tFileName).Close SaveChanges:=wdDoNotSaveChanges
                
                If i < k - 1 Then
                    WordDoc2.Sections.Add
                End If
                    
                WordDoc2.SaveAs2 primaryFilePath, 16
                
                Set WordDoc = Nothing
                Set WordDoc2 = Nothing
                
                Dim x: x = (i / k) * 100
                Dim y: y = (x * progressBox.staticBar.Width) / 100
                
                progressBox.progressBar.Width = y
                progressBox.progressText.Caption = currentName & " " & "(" & i - 1 & "/" & k - 1 & ")"
        Else
            j = j - 1
        End If
        Debug.Print Now(), fName, lName, gender, bDay, i, n
    Next i
    
    'Word.ScreenUpdating = True
    z = Shell("powershell.exe kill -processname winword", vbHide)
    Workbooks(selectedFileName).Close SaveChanges:=False
    Word.Quit SaveChanges:=wdDoNotSaveChanges
    MsgBox "Documents generated successfully."
        Unload progressBox
        Set WordDoc = Nothing
        Set WordDoc2 = Nothing
        CloseWordObjects
    
End Sub

Sub lol()
    Dim x: x = Shell("powershell.exe kill -processname winword", vbHide)
End Sub
