VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormTitle 
   Caption         =   "���������"
   ClientHeight    =   2880
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   6456
   OleObjectBlob   =   "FormTitle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButtonCancel_Click()
TextBoxProf.Value = ""
TextBoxGrade.Value = ""
TextBoxVar.Value = ""
FormTitle.Hide
End Sub

Private Sub CommandButtonOK_Click()
' �������� �� ������ �������� �����
    If TextBoxProf.Value = "" Or TextBoxGrade.Value = "" Or TextBoxVar.Value = "" Then
        MsgBox "� ������ ����� ���� �� ����� ���� �������!"
        Exit Sub
    End If
    
    ' �������� �� ��, ��� ������� �����
    If Num(TextBoxVar.Value) = "error" Then
        MsgBox "� ����� �������������� �����, ���������� ����!"
        Exit Sub
    End If
    
    ' �������� �� ��, ��� ������� ����� (�� ������������ � ���� ������)
'    If CInt(TextBoxGrade.Value) < 7 Or CInt(TextBoxGrade.Value) > 12 Then
'        MsgBox "����� ����� ���� �� 7 �� 11-���!"
'        Exit Sub
'    End If
    
    ' �������� �� ��, ��� ������� �����
    If CInt(TextBoxVar.Value) < 1 Or CInt(TextBoxVar.Value) > 10 Then
        MsgBox "�� ����� 10 ���������!"
        Exit Sub
    End If
    
    
Selection.TypeText "= " + TextBoxProf.Value + ", " + TextBoxGrade.Value + " �����" + ", " + "������� " + TextBoxVar.Value + " =" + vbLf
TextBoxProf.Value = ""
TextBoxGrade.Value = ""
TextBoxVar.Value = ""
FormTitle.Hide
End Sub

Private Sub help_Click()
    '����� �������
    HelpForm.refer = "������� �������� ������� �� �������� ����������� �������, " _
               + "������ ������� ����� ��� �������� ������������ ������� � ���������� ������ ����� ��������." + vbLf + vbLf _
               + "���� �������������� ���� ������� �������, ��� ����� ������� � ���� �������� ������� (�������)." + vbLf + vbLf _
               + "����� ���� ������� ������ ""��"", � ��������� ����������� ��������� ���������� � ��������� �������." + vbLf + vbLf _
               + "������� �� ������ ""������"" ������� ����� ��� ���������� ����������."

    HelpForm.Show
End Sub

