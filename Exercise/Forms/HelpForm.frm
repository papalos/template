VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HelpForm 
   Caption         =   "Справка"
   ClientHeight    =   7224
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   8232.001
   OleObjectBlob   =   "HelpForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public refer As String

Private Sub UserForm_Activate()
    TextBoxReference.Text = refer
End Sub

