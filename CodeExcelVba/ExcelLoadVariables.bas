Attribute VB_Name = "ExcelLoadVariables"
Option Explicit
Public returnedExcelFilePath As String
Public returnedFileName As String
Public variableMap As Object
Public originalWorkbook As Workbook


Sub ExcelLoadVariable_initiate()
    Set variableMap = CreateObject("scripting.dictionary")
    Set originalWorkbook = ActiveWorkbook
    
    GetExcelFilePath
    
    Application.ScreenUpdating = False
    OpenAWorkbook variableMap
    Application.ScreenUpdating = True
    originalWorkbook.Activate
    
    ProcessVariables variableMap
    
    MsgBox "Refreshing of variables have been completed."
    
End Sub

Sub AddToVariableMap(dict As Object, variableName As String, variableValue As String)
    Dim key, val
    
    key = variableName: val = variableValue
    If Not dict.Exists(key) Then
        dict.Add key, val
    Else
        dict(key) = value
    End If
End Sub

Sub GetExcelFilePath()

    Dim filePath As String
    Dim fileName As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = ""
    
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show <> 0 Then
            filePath = .SelectedItems(1)
        End If
    End With
    
    If filePath <> "" Then
        returnedExcelFilePath = filePath
        returnedFileName = fso.GetFilename(filePath)
    Else
        MsgBox "Failed to assign file path to this document"
        End
    End If
End Sub

Public Sub OpenAWorkbook(dict As Object)
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlWorksheet As Object
    Dim i As Long
        
    
    Dim xRet As Boolean
    xRet = IsWorkBookOpen(returnedFileName)
    Debug.Print xRet
    If xRet Then
        Set xlBook = Workbooks(returnedFileName)
    Else
        On Error Resume Next
            Set xlBook = Workbooks.Open(fileName:=returnedExcelFilePath)
            If Err Then
                MsgBox "File not found on current path"
                End
            End If
        On Error GoTo 0
    End If
    On Error Resume Next
        Set xlWorksheet = xlBook.Sheets("VARIABLES")
        If Err Then
            MsgBox "File does not contain sheet called VARIABLES, please alter file or use CTRL+ALT+F to mark new file path"
            End
        End If
    On Error GoTo 0
    
    i = 2
        
    Do While (xlWorksheet.Range("A" & i).value <> vbNullString Or xlWorksheet.Range("A" & i).value <> "")
        If (xlWorksheet.Range("A" & i).value <> vbNullString Or xlWorksheet.Range("A" & i).value <> "") Then
            dict(xlWorksheet.Range("A" & i).Text) = xlWorksheet.Range("B" & i).Text
        End If
        i = i + 1
    Loop
    
lbl_Exit:
    Set xlApp = Nothing
    Set xlBook = Nothing
    Exit Sub
End Sub

Function IsWorkBookOpen(name As String) As Boolean
    Dim xWb As Workbook
    On Error Resume Next
        Set xWb = Workbooks(name)
        Debug.Print (Not xWb Is Nothing)
        IsWorkBookOpen = (Not xWb Is Nothing)
End Function


Sub ProcessVariables(dict As Object)

    Dim mainWorkBook As Workbook
    Dim activeWorkSheet As Worksheet
    Set mainWorkBook = ActiveWorkbook
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim key As String
    
    lastRow = 0
    lastColumn = 0
    
    For i = 1 To 3
        Set activeWorkSheet = mainWorkBook.Sheets(mainWorkBook.Sheets(i).name)
        On Error Resume Next
            lastRow = activeWorkSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        On Error Resume Next
            lastColumn = activeWorkSheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
                        
        If lastRow > 0 And lastColumn > 0 Then
            For j = 1 To lastColumn
                For k = 1 To lastRow
                    If (activeWorkSheet.Cells(k, j).NoteText <> vbNullString) Then
                        key = Application.WorksheetFunction.Trim(activeWorkSheet.Cells(k, j).NoteText)
                        activeWorkSheet.Cells(k, j) = splitString(dict, key)
                    End If
                Next k
            Next j
        End If
    Next i
End Sub

Function splitString(dict As Object, originalString As String) As String

    Dim subStringArr() As String
    Dim targetString As String
    Dim partialRepString As String
    Dim value As String
    Dim key As String
    
    
    Dim i As Integer

    targetString = originalString
    subStringArr = Split(originalString, "{{")
    
    If UBound(subStringArr) > -1 Then
        For i = 0 To UBound(subStringArr)
            If InStr(subStringArr(i), "}}") > 0 Then
                key = Application.WorksheetFunction.Trim(Left(subStringArr(i), InStr(subStringArr(i), "}}") - 1))
                partialRepString = "{{" + Left(subStringArr(i), InStr(subStringArr(i), "}}") - 1) + "}}"
                If dict.Exists(key) Then
                    value = Application.WorksheetFunction.Trim(dict(key))
                    targetString = Replace(targetString, partialRepString, value)
                End If
            End If
        Next i
    End If
    splitString = targetString

End Function
