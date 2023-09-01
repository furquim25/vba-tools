Attribute VB_Name = "Tools"
'######################### Tools Module ##########################
'Created by: Rafael Furquim
'Last Updated (dd/mm/yyyy): 02/06/2022
'Working on: Version 2102 (Build 13801.20808)
'#################################################################

Public Sub test()
    'Open word
    Dim docWord As Word.Document
    Dim wordApp As Word.Application
    Set wordApp = New Word.Application
    wordApp.Visible = False
    docPath = "C:\Users\rfurquim\OneDrive - Digicorner\3 - SOO\Documento Operacional Tool\2 - Desenvolvimento\Documento Operacional Tool rev4\Auxiliar Doc Tool\teste.docx"
    Set docWord = wordApp.Documents.Open(docPath)
    docWord.AutoSaveOn = False
    
    'Call RemoveTextBetweenWords(docWord, "#DELETEANEXO5#", "#DELETEANEXO5#")
    'Call DeleteWordPage(docWord, 24)
    Page = findPageFromText(docWord, "#FINDPAGE#")
    'Close and quit word
    docWord.Close Savechanges:=True
    wordApp.Quit
End Sub

'From text location, return page
Function findPageFromText(docWord As Word.Document, searchText As String) As Long
    Dim p As Long 'page number
    Dim rngFound As Find
    Set rngFound = docWord.Range.Find
    rngFound.Text = searchText
    rngFound.Execute
    If rngFound.Found Then
        Page = rngFound.Parent.Information(wdActiveEndPageNumber)
    Else
        'not found
    End If
    findPageFromText = Page
End Function

'Deletes a Word page given the pageNumber
Public Sub DeleteWordPage(docWord As Word.Document, pageNumber As Variant):
    Set rng = docWord.Range(0, 0)
    Set rng = rng.GoTo(What:=wdGoToPage, Name:=pageNumber)
    Set rng = rng.GoTo(What:=wdGoToBookmark, Name:="\page")
    rng.Delete
    Set rng = Nothing
End Sub

'Given a word document object it will delete the specifed part between two words
Public Sub RemoveTextBetweenWords(docWord As Word.Document, strFirstWord As Variant, strLastWord As Variant, Optional removeWords As Boolean = True)
    Debug.Print (docWord.Name)
    With docWord.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = strFirstWord & "*" & strLastWord
        If removeWords = True Then
            .Replacement.Text = ""
        Else
            .Replacement.Text = strFirstWord & strLastWord
        End If
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    docWord.Info
    docWord.Selection.Information (wdActiveEndPageNumber)
End Sub

'Replace word in a Word documento Story object. (Reference: https://wordmvp.com/FAQs/Customization/ReplaceAnywhere.htm)
Public Sub SearchAndReplaceInStory(ByVal rngStory As Word.Range, ByVal strSearch As String, ByVal strReplace As String)
  With rngStory.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = strSearch
    .Replacement.Text = strReplace
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
  End With
End Sub

'Count Unique values in Range (Deprecated. Use UBound(Application.WorksheetFunction.Unique) instead)
Public Function CountUnique(rng As Range) As Integer
    Dim dict As Scripting.Dictionary
    Dim cell As Range
    Set dict = New Scripting.Dictionary
    For Each cell In rng.Cells
         If Not dict.Exists(cell.Value) Then
            dict.Add cell.Value, 0
        End If
    Next
    CountUnique = dict.Count
End Function

'Return unique values from Range (Deprecated. Use Application.WorksheetFunction.Unique instead)
Public Function UniqueValues(rng As Range) As Variant
    Dim dict As Scripting.Dictionary
    Dim cell As Range
    Set dict = New Scripting.Dictionary
    For Each cell In rng.Cells
        If Not dict.Exists(cell.Value) Then
            dict.Add cell.Value, 0
        End If
    Next
    countNoInterno = dict.Count
    
    'Save cell value appearence in order to array
    Dim uniqueArray As Variant
    ReDim uniqueArray(1 To countNoInterno)
    i = 1
    For Each cell In rng.Cells
        If dict.Exists(cell.Value) Then
            uniqueArray(i) = cell.Value
            i = i + 1
            dict.remove (cell.Value)
        End If
    Next
    UniqueValues = uniqueArray
    
End Function

'Compares if a string is in array of strings. Return True or False.
Public Function IsInArray(stringToBeFound As String, Arr As Variant) As Boolean
    Dim i
    For i = LBound(Arr) To UBound(Arr)
        If Arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

'Returns an array of unique values in vertical range/array A in vertical range/array B. If isInRange = False, it does the oposite (Range A not in Range B)
Public Function UniqueRangeAInRangeB(rangeA As Variant, rangeB As Variant, Optional invert As Boolean = False) As Variant
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    Dim rngValueA As Variant
    
    'Get unique arrays
    uniqueRangeA = Application.WorksheetFunction.Unique(rangeA)
    uniqueRangeB = Application.WorksheetFunction.Unique(rangeB)
    
    'Converts to 1D arrays
    If Tools.GetDimension(uniqueRangeA) > 1 Then uniqueRangeA = Application.Transpose(uniqueRangeA)
    If Tools.GetDimension(uniqueRangeB) > 1 Then uniqueRangeB = Application.Transpose(uniqueRangeB)
    
    'Save uniqueRangeA A values that are or not in uniqueRangeB in dictionary
    For i = 1 To UBound(uniqueRangeA)
        rngValueA = uniqueRangeA(i)
        If Tools.IsInArray(CStr(rngValueA), uniqueRangeB) = Not invert Then
            dict.Add rngValueA, 0
        End If
    Next
    
    UniqueRangeAInRangeB = dict.Keys
End Function

'Get the dimension of array
Function GetDimension(var As Variant) As Long
    On Error GoTo Err
    Dim i As Long
    Dim tmp As Long
    i = 0
    Do While True
        i = i + 1
        tmp = UBound(var, i)
    Loop
Err:
    GetDimension = i - 1
End Function

'Finds local path even if its inside one drive
Public Function AdresseLocal$(ByVal fullPath$)
    'fullPath$ = ThisWorkbook.Path
    'Finds local path for a OneDrive file URL, using environment variables of OneDrive
    'Reference https://stackoverflow.com/questions/33734706/excels-fullname-property-with-onedrive
    'Authors: Philip Swannell 2019-01-14, MatChrupczalski 2019-05-19, Horoman 2020-03-29, P.G.Schild 2020-04-02
    Dim ii&
    Dim iPos&
    Dim oneDrivePath$
    Dim endFilePath$
    Dim NbSlash
    
    If Left(fullPath, 8) = "https://" Then
        If InStr(1, fullPath, "sharepoint.com/") <> 0 Then 'Commercial OneDrive
            NbSlash = 4
        Else 'Personal OneDrive
            NbSlash = 2
        End If
        iPos = 8 'Last slash in https://
        For ii = 1 To NbSlash
            iPos = InStr(iPos + 1, fullPath, "/")
        Next ii
        endFilePath = Mid(fullPath, iPos)
        endFilePath = Replace(endFilePath, "/", Application.PathSeparator)
        For ii = 1 To 3
            oneDrivePath = Environ(Choose(ii, "OneDriveCommercial", "OneDriveConsumer", "OneDrive"))
            If 0 < Len(oneDrivePath) Then Exit For
        Next ii
        AdresseLocal = oneDrivePath & endFilePath
        While Len(Dir(AdresseLocal, vbDirectory)) = 0 And InStr(2, endFilePath, Application.PathSeparator) > 0
            endFilePath = Mid(endFilePath, InStr(2, endFilePath, Application.PathSeparator))
            AdresseLocal = oneDrivePath & endFilePath
        Wend
    Else
        AdresseLocal = fullPath
    End If
End Function

