Attribute VB_Name = "DocPropertyOperations"
Sub AddDocProperties_Initialization()

    AddDocProperty.Show
    Unload AddDocProperty
End Sub
Sub EditDocProperties_Initialization()

    
Dim theLabel As Object
Dim edtBox_n As Control
Dim labelCounter As Long

labelCounter = 0

For Each p In ActiveDocument.CustomDocumentProperties
    labelCounter = labelCounter + 1
    Set theLabel = editDocProperties.PropertyFrame.Controls.Add("Forms.Label.1", p.Name, True)
    With theLabel
        .Caption = p.Name
        .Left = 10
        .Width = 150
        .Top = 20 * labelCounter + 5
    End With
    
    Set edtBox_n = editDocProperties.PropertyFrame.Controls.Add("Forms.TextBox.1", BASE64SHA1(p.Name), True)
    With edtBox_n
        .Top = 20 * labelCounter
        .Left = 160
        .MultiLine = True
        .EnterKeyBehavior = True
        .Height = 20
        .Width = 500
        .Value = p.Value
    End With
Next
    editDocProperties.Show
    Unload editDocProperties
End Sub


Public Sub WriteProp(sPropName As String, sValue As String, _
      Optional lType As Long = msoPropertyTypeString)

'In the above declaration, "Optional lType As Long = msoPropertyTypeString" means
'that if the Document Property's Type is Text, we don't need to include the lType argument
'when we call the procedure; but if it's any other Prpperty Type (e.g. date) then we do

Dim bCustom As Boolean

  On Error GoTo ErrHandlerWriteProp

  'Try to write the value sValue to the custom documentproperties
  'If the customdocumentproperty does not exists, an error will occur
  'and the code in the errorhandler will run
  ActiveDocument.BuiltInDocumentProperties(sPropName).Value = sValue
  'Quit this routine
  Exit Sub

Proceed:
  'We know now that the property is not a builtin documentproperty,
  'but a custom documentproperty, so bCustom = True
  bCustom = True

Custom:
  'Try to set the value for the customproperty sPropName to sValue
  'An error will occur if the documentproperty doesn't exist yet
  'and the code in the errorhandler will take over
  ActiveDocument.CustomDocumentProperties(sPropName).Value = sValue
  Exit Sub

AddProp:
  'We came here from the errorhandler, so know we know that
  'property sPropName is not a built-in property and that there's
  'no custom property with this name
  'Add it
  On Error Resume Next
  ActiveDocument.CustomDocumentProperties.Add Name:=sPropName, _
    LinkToContent:=False, Type:=lType, Value:=sValue

  If Err Then
    'If we still get an error, the value isn't valid for the Property Type
    'e,g an invalid date was used
    Debug.Print "The Property " & Chr(34) & _
     sPropName & Chr(34) & " couldn't be written, because " & _
     Chr(34) & sValue & Chr(34) & _
     " is not a valid value for the property type"
  End If

  Exit Sub

ErrHandlerWriteProp:
  Select Case Err
    Case Else
   'Clear the error
   Err.Clear
   'bCustom is a boolean variable, if the code jumps to this
   'errorhandler for the first time, the value for bCustom is False
   If Not bCustom Then
     'Continue with the code after the label Proceed
     Resume Proceed
   Else
     'The errorhandler was executed before because the value for
     'the variable bCustom is True, therefor we know that the
     'customdocumentproperty did not exist yet, jump to AddProp,
     'where the property will be made
     Resume AddProp
   End If
  End Select

End Sub

Sub UpdateAllDocVariable()

'   Update all DocVariable fields in a document, even if in headers/footers or textboxes

'   Based on code at http://www.gmayor.com/installing_macro.htm
'   Charles Kenyon
'   18 October 2018
'
    Dim oStory As Range
    Dim oField As field
    '
    For Each oStory In ActiveDocument.StoryRanges
        For Each oField In oStory.Fields
            If oField.Type = wdFieldDocProperty Then oField.Update
        Next oField
        '
        If oStory.StoryType <> wdMainTextStory Then
            While Not (oStory.NextStoryRange Is Nothing)
            Set oStory = oStory.NextStoryRange
                For Each oField In oStory.Fields
                    If oField.Type = wdFieldDocProperty Then oField.Update
                Next oField
            Wend
        End If
        '
    Next oStory
    '
    Set oStory = Nothing
    Set oField = Nothing
End Sub
