VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormDifVarAns 
   Caption         =   "����� ���������� ��������� �������"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   8748.001
   OleObjectBlob   =   "FormDifVarAns.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormDifVarAns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public numAns As String
Public ans As String
Dim varAnsStr As String

Private Sub CommandButtonCancel_Click()
    Unload FormDifAns
    Unload FormDifVarAns
End Sub

Private Sub CommandButtonEnter_Click()
    ' �������� �� ���������� ���� ������ (Num - ��������� ����� �� ���)
    If Num(TextBoxVarScore.Value) = "error" Then
        MsgBox "���������� ������ �� ������������� ������!"
        Exit Sub
    End If
    
    ' �������� �� ����� ���� �� ������ ����������� ������
    For Z = 0 To ListBoxAllRightAns.ListCount - 1 Step 1
        If ListBoxAllRightAns.Selected(Z) Then
            Z = -1
            Exit For
        End If
    Next Z
    If Z <> -1 Then
        MsgBox "��������� �� ������ ������ ��� ��������!"
        Exit Sub
    End If
    
    ' ���� ������ ���
    Dim numbersOfAnsver As String
    numbersOfAnsver = ""
    ' ���� ������ ���������� ������� �� ����
    If ListBoxAllRightAns.ListCount <> 0 Then
        For p = 0 To ListBoxAllRightAns.ListCount - 1 Step 1
            If ListBoxAllRightAns.Selected(p) Then
                ' �������� ���������� ������ ���� ������� �� ����� FormDifAns
                For i = 0 To FormDifAns.ListBoxRight.ListCount - 1 Step 1
                    ' � ���������� �� � �������� ��������� �� ����� FormDifVarAns
                    If FormDifAns.ListBoxRight.List(i) = ListBoxAllRightAns.List(p) Then
                        ' ���� ������� ����� ��� ���������� ����� �� ������ ������ ���� ��������
                        numbersOfAnsver = numbersOfAnsver + CStr(i + 1) + ", "
                    End If
                Next i
            End If
        Next p

        If numbersOfAnsver <> "" Then numbersOfAnsver = Left(numbersOfAnsver, Len(numbersOfAnsver) - 2)
        varAnsStr = varAnsStr + "### (" + TextBoxVarScore.Value + " " + Num(TextBoxVarScore.Value) + ") " + numbersOfAnsver + vbLf
        TextBoxRightAns.Text = varAnsStr + vbLf
    End If
    
    '������� �������� � ���� ������ � ������� ��������� � ������ �������
    TextBoxVarScore.Value = ""
    For i = 0 To ListBoxAllRightAns.ListCount - 1
        ListBoxAllRightAns.Selected(i) = False
    Next
    
    
End Sub

Private Sub CommandButtonExit_Click()
    If varAnsStr = "" Then
        MsgBox "�������� ���������� ������� �� ������!"
        Exit Sub
    End If

    ' ������� ������ � ��������
    Selection.TypeText _
    "== ������� " + numAns + " ==" + vbLf _
    + "������� ���� ����� �������." + vbLf _
    + "=== ������ (������������� �����) ===" + vbLf _
    + ans _
    + "=== ������ ===" + vbLf _
    + varAnsStr
    
    ' ��������� ������ �������
    Selection.TypeText vbLf + vbLf
    
    Unload FormDifAns
    Unload FormDifVarAns
End Sub


Private Sub help_Click()
    HelpForm.refer = "�������� ������������������ ���������� �������, ������� ���������� ������ ����������� �� ��� ������������������ � ������� ������ ""�������� �����""." + vbLf + vbLf _
                   + "� ������ ���� ����������� ������� ������ � ��������� ������ �� ����." + vbLf + vbLf _
                   + "�� ������ ���������� ������� ���������� ������ ������� � ��������� �� �������� ������ ""�������� �����""." + vbLf + vbLf _
                   + "������� �� ������ ""������� ������"" �������� � �������� ����� � ������ ������� � ����� ��������� Word." + vbLf + vbLf _
                   + "� ����������� ���������� ������ ��������������� ��������� � ������ �������, ������ ����� ������� ��������������� � ��������� Word." + vbLf + vbLf _
                   + "������� �� ������ ""������� �����"" �������� � �������� ����� ��� ���������� ���������� � ������ ������� �� �����."
    HelpForm.Show
End Sub

Private Sub UserForm_Click()

End Sub
