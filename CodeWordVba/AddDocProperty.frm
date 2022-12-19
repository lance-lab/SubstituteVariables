VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddDocProperty 
   Caption         =   "Add Property"
   ClientHeight    =   3405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11940
   OleObjectBlob   =   "AddDocProperty.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddDocProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddPropertyCancel_Click()
    FormCancelButton = True
    Unload AddDocProperty
End Sub
Private Sub AddPropertyOk_Click()
    FormCancelButton = False
    AddDocProperty.hide
  
    If (AddDocProperty.AddDocPropertyNameTxt.Value <> vbNullString Or AddDocProperty.AddDocPropertyValueTxt.Value <> vbNullString) Then
        Call WriteProp(sPropName:=AddDocProperty.AddDocPropertyNameTxt.Value, sValue:=AddDocProperty.AddDocPropertyValueTxt.Value)
        Call UpdateAllDocVariable
    Else
        MsgBox "Either name or value is empty!"
    End If
    
    Call UpdateAllDocVariable
End Sub

Private Sub UserForm_Click()

End Sub
