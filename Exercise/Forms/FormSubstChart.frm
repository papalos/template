VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSubstChart 
   Caption         =   "������������������ �����������"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   9036.001
   OleObjectBlob   =   "FormSubstChart.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSubstChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If ListBoxAns.ListIndex > -1 Then
        ListBoxRightSec.AddItem ListBoxAns.Value
    End If
End Sub

' �������� �������� �� ������ �������
Private Sub CommandButton2_Click()
    If ListBoxRightSec.ListIndex > -1 Then
        Dim remove As Integer
        remove = ListBoxRightSec.ListIndex
        ListBoxRightSec.ListIndex = -1
        ListBoxRightSec.RemoveItem remove
    End If
End Sub

Private Sub CommandButtonExit_Click()
    Unload FormSubstChart
End Sub

Private Sub CommandButtonVar_Click()
    For i = 0 To ListBoxRightSec.ListCount - 1 Step 1
        FormSelect.ListBoxLeft.AddItem ListBoxRightSec.List(i)
    Next i
    Unload FormSubstChart
    FormSelect.Show
End Sub

Private Sub help_Click()
    HelpForm.refer = "�� ����� ���������� ���������� ���������� ������� ������ ������� � ������������ � ���������� � �������." + vbLf + vbLf _
                   + "������ ���� �� �������� ������ � ����� ���� ���������� ��� � ������." + vbLf + vbLf _
                   + "����������� ���������� ������������������ ������� � ������ ����." + vbLf + vbLf _
                   + "������ ����� �� �������� � ������ ���� ������� ������ � ���������� �� ������� � ����� ����." + vbLf + vbLf _
                   + "������� ������ ""������ ��������"" ��������� ����� ""������������ ��������� �������"""
    HelpForm.Show
End Sub

'Private Sub ListBoxAns_Click()
'    Dim remove As Integer
'    remove = ListBoxAns.ListIndex
'    ListBoxRightSec.AddItem ListBoxAns.Value
'    ListBoxAns.ListIndex = -1
'    ListBoxAns.RemoveItem remove     '������������������ �������� � ������ ����� �� ���������
'End Sub
'
'Private Sub ListBoxRightSec_Click()
'    Dim remove As Integer
'    remove = ListBoxRightSec.ListIndex
'    ListBoxAns.AddItem ListBoxRightSec.Value  ' ����������������� ������� � ������ ���� �� ���������
'    ListBoxRightSec.ListIndex = -1
'    ListBoxRightSec.RemoveItem remove
'End Sub

Private Sub UserForm_Click()

End Sub
