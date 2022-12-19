Attribute VB_Name = "ExcelFileOperations"
Public Sub GetExcelFilePath()

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
    Dim xlRet As Boolean
    
    i = 2
    xlRet = IsWorkBookOpen(returnedFileName)

    If xlRet Then
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
