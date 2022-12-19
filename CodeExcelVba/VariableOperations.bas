Attribute VB_Name = "VariableOperations"
Public Sub AddToVariableMap(dict As Object, variableName As String, variableValue As String)
    Dim key, val
    
    key = variableName: val = variableValue
    If Not dict.Exists(key) Then
        dict.Add key, val
    Else
        dict(key) = value
    End If
End Sub

Public Sub ProcessVariables(dict As Object)

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
