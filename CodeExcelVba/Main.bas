Attribute VB_Name = "Main"
Option Explicit
Public returnedExcelFilePath As String
Public returnedFileName As String
Public variableMap As Object
Public originalWorkbook As Workbook

Sub ProcessExcelVariables_initiate()
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
