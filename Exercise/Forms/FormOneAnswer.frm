VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormOneAnswer 
   Caption         =   "���� �����"
   ClientHeight    =   7212
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   5880
   OleObjectBlob   =   "FormOneAnswer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormOneAnswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonCloseOneAns_Click()
    Unload FormOneAnswer
End Sub

'������� ���������� �������
Private Sub CommandButtonDel_Click()
    ListBoxVarAnsOne.RemoveItem ListBoxVarAnsOne.ListIndex
End Sub

' ���������� ������� � ������ ����
Private Sub CommandButtonInsert_Click()
    '���� ��������� ���� �� ������
    If TextBoxVarAnsOne.Text <> "" Then
        Dim ArrAns() As String
        Dim j As Integer
        k = 0
        ArrAns = Split(TextBoxVarAnsOne.Text, vbLf)
        
        ' /////�������� �� ���������� �������� �������////
        If Not Duty.NoIdentical(ArrAns) Then
            MsgBox "�� ����� ���� ���� ���������� �������!"
            TextBoxVarAnsOne.SetFocus
            Exit Sub
        End If
        For h = 0 To UBound(ArrAns)
        For t = 0 To ListBoxVarAnsOne.ListCount - 1
            If h = UBound(ArrAns) Then
                If ArrAns(h) = ListBoxVarAnsOne.List(t) Then
                    MsgBox "�� ����� ���� ���� ���������� �������!"
                    TextBoxVarAnsOne.SetFocus
                    Exit Sub
                End If
            Else
                If Duty.NotEndSimbol(CStr(ArrAns(h))) = ListBoxVarAnsOne.List(t) Then
                    MsgBox "�� ����� ���� ���� ���������� �������!"
                    TextBoxVarAnsOne.SetFocus
                    Exit Sub
                End If
            End If
        Next t
        Next h
        '////////////////////////////////////////////////
            
        
        ' ��������� ������ �� textBox � listBox
        For Each Item In ArrAns
            If j = UBound(ArrAns) Then
                If Item <> "" Then
                If Item <> " " Then
                    ListBoxVarAnsOne.AddItem Item
                End If
                End If
            Else
                If Left(Item, Len(Item) - 1) <> "" Then  ' ���������� ������ ������ � �������
                If Left(Item, Len(Item) - 1) <> " " Then
                    ListBoxVarAnsOne.AddItem Left(Item, Len(Item) - 1)
                End If
                End If
            End If
            j = j + 1
        Next Item
        TextBoxVarAnsOne.Text = ""
    End If
End Sub

Private Sub CommandButtonOkOneAns_Click()
    ' �������� �� ������ �������� �����
    If TextBoxNumAnsOne.Value = "" Or TextBoxScoreOneAns.Value = "" Or ListBoxVarAnsOne.ListCount < 1 Then
        MsgBox "� ������ ����� ���� �� ����� ���� �������!"
        Exit Sub
    End If
    
    ' �������� �� ���������� ��������� �������, �� ������ ���� �� 2 �� 7
    If ListBoxVarAnsOne.ListCount < 2 Then
            MsgBox "��������� ������� ������ ���� ������ ������!"
            TextBoxVarAnsOne.SetFocus
            Exit Sub
    ElseIf ListBoxVarAnsOne.ListCount > 7 Then
            MsgBox "���������� ��������� ������� �� ������ ��������� ����!"
            TextBoxVarAnsOne.SetFocus
            Exit Sub
    End If
    
    ' �������� �� ���������� ���� ������
    If Num(TextBoxScoreOneAns.Value) = "error" Then
        MsgBox "���������� ������ �� ������������� ������!"
        Exit Sub
    End If
    
    ' �������� �� ��, ��� ������ ������� ����������� ������
    If ListBoxVarAnsOne.ListIndex < 0 Then
        MsgBox "�� ������ ���������� �����!"
        Exit Sub
    End If
    

    
    ' ���� ������ �� ����
    Dim str As String
    str = ""
    
    If ListBoxVarAnsOne.ListCount <> 0 Then
        Dim k As Integer
        k = 0

        Do While k < ListBoxVarAnsOne.ListCount
            If ListBoxVarAnsOne.List(k, 0) = ListBoxVarAnsOne.Value Then
                str = str + "## " + ListBoxVarAnsOne.List(k, 0) + vbLf
            Else
                str = str + "# " + ListBoxVarAnsOne.List(k, 0) + vbLf
            End If
            k = k + 1
        Loop
    End If
    
    
    '' ����� ������� ������� � ��������
    Selection.TypeText _
    "== ������� " + TextBoxNumAnsOne.Value + " ==" + vbLf _
    + "������� ���� ����� �������." + vbLf _
    + "=== ������ (������������ �����) ===" + vbLf _
    + str _
    + "=== ������ ===" + vbLf _
    + "### (" + TextBoxScoreOneAns.Value + " " + Num(TextBoxScoreOneAns.Value) + ") " + CStr(ListBoxVarAnsOne.ListIndex + 1) + vbLf 'TextBoxRightOneAns.Value
    
    ' ��������� ������ �������
    Selection.TypeText vbLf + vbLf
    
    Unload FormOneAnswer

End Sub


Private Sub helpOne_Click()
    '����� �������
    HelpForm.refer = "������� ���������� ����� �������." + vbLf + vbLf _
                          + "����� ����������� �������� ������� � ������� ��������� ����, ����� ������ ��������� ����� �� ����� �������." _
                          + "��� �������� �� ����� ������ ����������� ��������� ������ Shift+Enter." + vbLf + vbLf _
                          + "������� �� ������ ""�������� ��������"" ��������� ����� �� �������� ���� � ������." _
                          + "������� �� ������ ""������� ����������"" ������ ���������� ������� �� ������� ����" _
                          + "�� ���������� ������������ ������� ������� � ������ ���� ���������� ������� ���������� �����, ������� �� ��� �����." + vbLf + vbLf _
                          + "������ ������� ���������� ������, ����������� �� ������ �����." + vbLf + vbLf _
                          + "����� ������� �� ������ ""�������� ������"" � ����� ������ ��������� ����� ������� ������." + vbLf + vbLf _
                          + "����� ���� ���������� ��������������� ������ ��������������� � ��������� Word, ������ ����� �������." + vbLf + vbLf _
                          + "�������� ����� �������� �� ������ ""������� �����"" �������� � �� �������� ��� ���������� ����������."
    HelpForm.Show
End Sub



Private Sub UserForm_Click()

End Sub
