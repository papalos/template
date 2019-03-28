VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormOpenAns 
   Caption         =   "�������� ������"
   ClientHeight    =   4632
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   5520
   OleObjectBlob   =   "FormOpenAns.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormOpenAns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public varAnsStr As String
Public varAnsNum As Integer



Private Sub CommandButtonOpenClose_Click()
    ' �������� �� �� ��� ������ ������� ���� ������ ������ ��������,
    ' ��� ���������� ������ ������ ��������
    If varAnsStr = "" Then
        MsgBox "������ ������� ������ ""�������� �������"", ��� ���������� �������� ������!"
        Exit Sub
    End If
    
    Selection.TypeText _
    "== ������� " + TextBoxOpenNum.Value + " ==" + vbLf _
    + "������� ���� ����� �������." + vbLf _
    + "=== ������ ===" + vbLf _
    + varAnsStr
    
    ' ��������� ������ �������
    Selection.TypeText vbLf + vbLf
    
    varAnsStr = ""
    varAnsNum = 0
    Unload FormOpenAns
End Sub


Private Sub CommandButtonOpenExit_Click()
    varAnsStr = ""
    varAnsNum = 0
    Unload FormOpenAns
End Sub

Private Sub CommandButtonOpenOK_Click()
    ' �������� �� ������ �������� �����
    If TextBoxOpenNum.Value = "" Or TextBoxOpenVarAns.Value = "" Or TextBoxOpenVarScore.Value = "" Then
        MsgBox "� ������ ����� ���� �� ����� ���� �������!"
        Exit Sub
    End If
    
    ' �������� �� ���������� ���� ������
    If Num(TextBoxOpenVarScore.Value) = "error" Then
        MsgBox "���������� ������ �� ������������� ������!"
        Exit Sub
    End If
    
    
   If Trim(TextBoxOpenVarAns.Value) = "" Then
        MsgBox "��������� ����� ����������� ���������!"
        Exit Sub
   End If
    
    ' ���� ������ �� ����
    varAnsStr = varAnsStr + "### (" + TextBoxOpenVarScore.Value + " " + Num(TextBoxOpenVarScore.Value) + ") " + TextBoxOpenVarAns.Value + vbLf
    varAnsNum = varAnsNum + 1
    TextBoxVarOpen.Text = "������� ������ �" + CStr(varAnsNum) + " ��������!" + vbLf + varAnsStr
    TextBoxOpenNum.Enabled = False
    TextBoxOpenVarScore.Value = ""
    TextBoxOpenVarAns.Value = ""
End Sub

Private Sub help_Click()
    HelpForm.refer = "������� ���������� ����� �������." + vbLf + vbLf _
                   + "����� ������� ������� ����������� ������ � ���������� ������ �� ����." + vbLf + vbLf _
                   + "���� ������� ���������, �� ����� ������� �� ������ ""�������� �����"" ������� ��������� ������� � ���������� ������ �� ����." + vbLf + vbLf _
                   + "����� ����������� ��������� ����� ���� ����� ����, �� ��� ������� ����� ���� �������." + vbLf + vbLf _
                   + "��������, �� ������: ������� ���������� ���� � �����, �� ������ ������?" + vbLf + vbLf _
                   + "����������� ���� ��������� ���������: 3 - 10 ������, ��� - 10 ������, ���� - 10 ������" + vbLf + vbLf _
                   + "������� ���� ��� ����� ������ �� �����������, ""������"" � ""������"" � ""������"" - ��� ������� ��� ���������� �������� (��. ���������� �� ���������� �������)." + vbLf + vbLf _
                   + "������� �� ������ ""�������� �����"" ����� � ������ �������� ������ � ������, ������ ��� � ���� ��� ����������� ��������� ������, � ����������� ����� ������ ��������." _
                   + "(��������� �������� ������ ��������� ������� ����������� ����� ���!)" + vbLf + vbLf _
                   + "������� �� ������ ""�������� ������"" �������� � ������ ������� � ����� ��������� � �������� �����" + vbLf + vbLf _
                   + "������� �� ������ ""�������"" �������� � �������� ����� ��� ������ �������."
    HelpForm.Show
End Sub

Private Sub TextBoxOpenVarAns_Change()

End Sub

Private Sub UserForm_Click()

End Sub
