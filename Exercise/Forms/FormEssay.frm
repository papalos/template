VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormEssay 
   Caption         =   "����"
   ClientHeight    =   3192
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3768
   OleObjectBlob   =   "FormEssay.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormEssay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonEssayExit_Click()
    Unload FormEssay
End Sub

Private Sub CommandButtonOkOneAns_Click()
    ' �������� �� ������ �������� �����
    If TextBoxNumEss.Value = "" Or TextBoxLongAns.Value = "" Or TextBoxScoreEss.Value = "" Then
        MsgBox "� ������ ����� ���� �� ����� ���� �������!"
        Exit Sub
    End If
    
    ' �������� �� ������ �������� �����
    If TextBoxLongAns.Value < 1 Or TextBoxLongAns.Value > 1024 Then
        MsgBox "�������� ���� ""����� ������"" � ������������ ���������!"
        Exit Sub
    End If
    
    ' �������� �� ���������� ���� ������
    If Num(TextBoxScoreEss.Value) = "error" Then
        MsgBox "���������� ������ �� ������������� ������!"
        Exit Sub
    End If
    
    Selection.TypeText _
    "== ������� " + TextBoxNumEss.Value + " ==" + vbLf _
    + "���� ����������� ����� �������." + vbLf _
    + "=== �������� ===" + vbLf _
    + "### ����� ������: " + TextBoxLongAns.Value + vbLf _
    + "=== ������ ===" + vbLf _
    + "### (" + TextBoxScoreEss.Value + " " + Num(TextBoxScoreEss.Value) + ") "
    
    Selection.TypeText vbLf + vbLf

    
    Unload FormEssay
End Sub

Private Sub UserForm_Click()

End Sub
