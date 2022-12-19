Attribute VB_Name = "ExcelLoadVariables"
Public returnedFromSub As Boolean
Public returnedExcelFilePath As String
Public returnedExcelFilePathExact As String


Sub LoadDocPropertiesFromExcel_Initialization()
    returnedFromSub = False
    OpenAWorkbook
    UpdateAllDocVariable
    MsgBox "All variables have been refreshed!"
End Sub


Sub GetExcelFilePath_Initialization()
    Dim filePath As String
    'Dim fso As New Scripting.FileSystemObject
    
    filePath = ""
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show <> 0 Then
            filePath = .SelectedItems(1)
        End If
    End With
    
    If filePath <> "" Then
        Call WriteProp(sPropName:="ExcelFilePath", sValue:=filePath)
        returnedExcelFilePath = filePath
        
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        returnedExcelFilePathExact = objFSO.GetParentFolderName(filePath)
        
        MsgBox "Path: " + filePath + " has been added to this document"
        returnedFromSub = True
    Else
        MsgBox "Failed to assign file path to this document"
        End
    End If
End Sub

Public Sub OpenAWorkbook()
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlWorksheet As Object
    Dim i As Long
        
    Dim strWorkbookName As String
    
    If (returnedExcelFilePath <> vbNullString And returnedExcelFilePath <> "") Then
        strWorkbookName = returnedExcelFilePath
    Else
        For Each p In ActiveDocument.CustomDocumentProperties
            If p.Name = "ExcelFilePath" Then
                strWorkbookName = p.Value
                Exit For
            End If
        Next
    End If
    
    On Error Resume Next
        Set xlApp = GetObject(, "Excel.Application")
        If Err Then
            Set xlApp = CreateObject("Excel.Application")
            xlApp.Visible = True
        End If
    On Error GoTo 0
    
    
    On Error Resume Next
        Set xlBook = xlApp.Workbooks.Open(fileName:=strWorkbookName)
        If Err Then
            MsgBox "File not found on current path, please use CTRL+ALT+F to mark the file path"
            End
        End If
    On Error GoTo 0
    
    On Error Resume Next
        Set xlWorksheet = xlBook.Sheets("VARIABLES")
        If Err Then
            MsgBox "File does not contain sheet called VARIABLES, please alter file or use CTRL+ALT+F to mark new file path"
            End
        End If
    On Error GoTo 0
    
    i = 2
    
    Do While (xlWorksheet.Range("A" & i).Value <> vbNullString Or xlWorksheet.Range("A" & i).Value <> "")
        If (xlWorksheet.Range("A" & i).Value <> vbNullString Or xlWorksheet.Range("A" & i).Value <> "") Then
            Call WriteProp(sPropName:=xlWorksheet.Range("A" & i).Value, sValue:=xlWorksheet.Range("B" & i).Value)
        End If
        i = i + 1
    Loop
    
lbl_Exit:
    Set xlApp = Nothing
    Set xlBook = Nothing
    returnedFromSub = True
    Exit Sub
End Sub
