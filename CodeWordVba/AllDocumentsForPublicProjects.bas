Attribute VB_Name = "AllDocumentsForPublicProjects"
Public scenario As String

Sub AllDocumentsCreate_Initialization()
    
    Dim sysFolder As String
    Dim sysName As String
    Dim i As Long
    Dim j As Long
    Dim file As Object
    Dim listOfFileNames() As String

           
    GetExcelFilePath_Initialization
    OpenAWorkbook
    
    sysFolder = ""
    
    For Each p In ActiveDocument.CustomDocumentProperties
        If (p.Name = "SystemovyPriecinok") Then
            sysFolder = p.Value
        End If
        
        If (sysFolder <> "") Then
            Exit For
        End If
    Next
    
    If (sysFolder <> "") Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFolders = objFSO.GetFolder(sysFolder).SubFolders
        folderCount = objFolders.Count
        
        If folderCount > 0 Then
            
            For Each objFolder In objFolders
                DecideScenario.ComboBox1.AddItem objFolder.Name
            Next objFolder
            DecideScenario.ComboBox1.ListIndex = 0
            DecideScenario.Show
                
            If (scenario <> "") Then
                sysFolder = sysFolder + "\" + scenario
                Debug.Print (sysFolder)
                
                Set objFiles = objFSO.GetFolder(sysFolder).Files
                fileCount = objFiles.Count
                
                If fileCount > 0 Then
                    ReDim listOfFileNames(1 To fileCount)
                    j = 0
                    For Each objFile In objFiles
                        j = j + 1
                        listOfFileNames(j) = objFile.Name
                    Next objFile
                Else
                    MsgBox "No templates found on path: " + sysFolder
                    End
                End If
            
            Else
                MsgBox "No scenario was selected"
                End
            End If
        End If
    
        For i = LBound(listOfFileNames) To UBound(listOfFileNames)
            Call ProcessDocument(listOfFileNames(i), sysFolder)
        Next i
        Documents.Close
    Else
        MsgBox "System unable to create directory due to incorrect folder path"
        End
    End If
End Sub


Public Sub ProcessDocument(fileName As String, sysFolder As String)
    Debug.Print fileName
    Debug.Print folderPath
    
    Dim strFile As String
    Dim oDoc As Documents

    strFile = sysFolder + "\" + fileName     'change to path of your file
    Debug.Print (strFile)
    
    If Dir(strFile) <> "" Then    'First we check if document exists at all at given location
    
        On Error Resume Next
            Documents.Open fileName:=strFile, ReadOnly:=True, Visible:=True
            If Err Then
                MsgBox "System was not able to open template file: " + fileName
                End
            End If
        On Error GoTo 0
    
    
        OpenAWorkbook
        Call WriteProp(sPropName:="ExcelFilePath", sValue:=returnedExcelFilePath)
        UpdateAllDocVariable
        
        On Error Resume Next
            ActiveDocument.SaveAs2 fileName:=returnedExcelFilePathExact + "\" + fileName, fileFormat:=wdFormatXMLDocumentMacroEnabled
            If Err Then
                MsgBox "System was not able to create file from template file: " + fileName
                End
            End If
        On Error GoTo 0
    End If
End Sub
