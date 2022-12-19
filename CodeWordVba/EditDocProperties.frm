VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditDocProperties 
   Caption         =   "Edit DOC Properties"
   ClientHeight    =   11025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15375
   OleObjectBlob   =   "EditDocProperties.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EditDocProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub EditPropertyCancel_Click()
    FormCancelButton = True
    Unload editDocProperties
End Sub
Private Sub EditPropertyOk_Click()
    FormCancelButton = False
    editDocProperties.hide
    
    Dim ctrl As Control
    Dim absorb_text As String
    
    For Each p In ActiveDocument.CustomDocumentProperties
        For Each ctrl In editDocProperties.Controls
            If (TypeName(ctrl) = "TextBox" And ctrl.Name = BASE64SHA1(p.Name)) Then
                Call WriteProp(sPropName:=p.Name, sValue:=ctrl.Text)
                Debug.Print (ctrl.Name)
            End If
        Next ctrl
    Next

    Call UpdateAllDocVariable
End Sub

Private Sub CommandButton1_Click()

End Sub
