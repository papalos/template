VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSubstVar_NotUse_ 
   Caption         =   "���������� ����������� ������"
   ClientHeight    =   4656
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8280.001
   OleObjectBlob   =   "FormSubstVar_NotUse_.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSubstVar_NotUse_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonIns_Click()
    Selection.TypeText TextBoxAnswers.Text
    TextBoxAnswers.Text = ""
End Sub

Private Sub CommandButtonSubstCancel_Click()
    Unload FormSubstVar
End Sub

Private Sub CommandButtonSubstOK_Click()

    Dim ArrText() As String
    Dim ins As String
    Dim counting As Integer
    
    counting = 0
    ArrText = Split(TextBoxSubst.Text, vbLf)
    
    For Each s In ArrText
        counting = counting + 1
        If counting = 1 Then
            If counting = UBound(ArrText) + 1 Then
                ins = ins + CStr(counting) + "- " + s
            Else
                ins = CStr(counting) + "- " + Left(s, Len(s) - 1)
            End If
        ElseIf counting = UBound(ArrText) + 1 Then
            ins = ins + ";" + CStr(counting) + "- " + s
        Else
            ins = ins + ";" + CStr(counting) + "- " + Left(s, Len(s) - 1)
        End If
    Next s

    
    ' �������� �� ������ �������� �����
    If TextBoxSubst.Value = "" Or TextBoxSubstScore.Value = "" Then
        MsgBox "� ������ ����� ���� �� ����� ���� �������!"
        Exit Sub
    End If
    
    ' �������� �� ���������� ���� ������
    If Num(TextBoxSubstScore.Value) = "error" Then
        MsgBox "���������� ������ �� ������������� ������!"
        Exit Sub
    End If
    
    TextBoxAnswers.Text = TextBoxAnswers.Text + "### (" + TextBoxSubstScore.Value + " " + Num(TextBoxSubstScore.Value) + ") " + ins + vbLf
    TextBoxSubstScore.Value = ""
    TextBoxSubst.Value = ""
End Sub

Private Sub help_Click()
    HelpForm.refer = "��� ��� � ���������� ���� �� ���� ������ ��������� �����������, � ���� ���� ���������� �� ������������." + vbLf + vbLf _
                   + "� ������� ���� ������� ������������������ �����������, ������� ���� �� ��� � ������� ""�������� �����""." + vbLf + vbLf _
                   + "������������������ ����������� � ������ ���� ""�������� �������""" + vbLf + vbLf _
                   + "�������� ����������� ���������� �������������������, �������� ���� �� ������, ����������� ���� �������� ������ ""�������� �����""" + vbLf + vbLf _
                   + "������� �� ������ ""������� ������"" �������� � ������ ������������������� ��������� � ���� ""�������� �������"" � ����� ��������� Word  � �������� ����." + vbLf + vbLf _
                   + "������� ""���� ���������� �������"" �� ������� ""���������� �������"" ����� ��������������� �������� ��� �������� ����� ""���� ���������� �������"", ����� �������� ����������� �������� �������." _
                   + "������ ""�������� �����"" ���������� ������� ������ � �������� � ����� ��������� ������� (������� �� ���, ����� �� ��������� � ������ ����� ����� ��������� ===������===."
    HelpForm.Show
End Sub

Private Sub UserForm_Click()

End Sub
