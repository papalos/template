VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormDifAns 
   Caption         =   "������ � ����������� ���������� ������"
   ClientHeight    =   6552
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   6588
   OleObjectBlob   =   "FormDifAns.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormDifAns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonClose_Click()
    Unload FormDifAns
End Sub
' ��������� ���������� ������ � ������ ���� ���������� �������
Private Sub CommandButtonDifAdd_Click()
    Dim AllAns() As String
    
    Dim j As Integer
    j = 0
    AllAns = Split(TextBoxAllAns.Text, vbLf)
    
    
    
    '  ++++ �������� �� ���������� �������� ������� ++++
    If Not Duty.NoIdentical(AllAns) Then
        MsgBox "�� ����� ���� ���� ���������� �������!"
        TextBoxAllAns.SetFocus
        Exit Sub
    End If
    For Each h In AllAns
        ' �������� �� ������ ������� ������
'        If h = "" Then
'            MsgBox "� ����������� ��������� ������������ ������ ������� ������, ������� ���, ��� �������� �����!"
'            TextBoxAllAns.SetFocus
'            Exit Sub
'        End If
        For t = 0 To ListBoxRight.ListCount - 1
            If h = ListBoxRight.List(t) Then
                MsgBox "�� ����� ���� ���� ���������� �������!"
                TextBoxAllAns.SetFocus
                Exit Sub
            End If
        Next t
    Next h
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    ' ��������� ListBox ��� ������ ���������� �������
    For Each Item In AllAns
        If j < UBound(AllAns) Then
            If Left(Item, Len(Item) - 1) <> "" Then  ' ���������� ������ ������ � �������
            If Left(Item, Len(Item) - 1) <> " " Then
                ListBoxRight.AddItem Left(Item, Len(Item) - 1)
            End If
            End If
        Else
            If Item <> "" Then
            If Item <> " " Then
                ListBoxRight.AddItem Item
            End If
            End If
        End If
        j = j + 1
    Next Item
    
    ' ������� ���� �����
    TextBoxAllAns.Value = ""
End Sub
' �������� ��������� �� ������
Private Sub CommandButtonDivRemove_Click()
    Dim f As Integer
    f = 0
    Do While (f < ListBoxRight.ListCount)
        If ListBoxRight.Selected(f) Then
            ListBoxRight.RemoveItem (f)
            f = -1
        End If
        f = f + 1
    Loop
End Sub

Private Sub CommandButtonOK_Click()
    Dim ans As String   '������ ��� �������
    ans = ""
    
    ' �������� �� ������ �������� �����
    If TextBoxNumDifAns.Value = "" Or ListBoxRight.ListCount < 1 Then
        MsgBox "� ������ ����� ���� �� ����� ���� �������!"
        Exit Sub
    End If
    
    ' �������� �� ���������� ��������� �������, �� ������ ���� �� 2 �� 7
    If ListBoxRight.ListCount < 2 Then
        MsgBox "��������� ������� ������ ���� ������ ������!"
        TextBoxAllAns.SetFocus
        Exit Sub
    ElseIf ListBoxRight.ListCount > 7 Then
        MsgBox "���������� ��������� ������� �� ������ ��������� ����!"
        TextBoxAllAns.SetFocus
        Exit Sub
    End If
    
    ' �������� �� ����� ���� �� ������ ����������� ������
    For y = 0 To ListBoxRight.ListCount - 1 Step 1
        If ListBoxRight.Selected(y) Then
            y = -1
            Exit For
        End If
    Next y
    If y <> -1 Then
        MsgBox "��������� �� ������ ����������� ������!"
        Exit Sub
    End If
    

    ' ���� ������ �� ����
    Dim j As Integer
    j = 0
    For k = 0 To ListBoxRight.ListCount - 1 Step 1
        If ListBoxRight.Selected(k) Then
            '������� �� ����� ������ ��������� ������ ���������� �������
            FormDifVarAns.ListBoxAllRightAns.AddItem ListBoxRight.List(k, 0)
            '�������� ����� ��� ����������
            ans = ans + "## " + ListBoxRight.List(k, 0) + vbLf
        Else
            '�������� ����� ��� �� ���������
            ans = ans + "# " + ListBoxRight.List(k, 0) + vbLf
        End If
    Next k
    
    '�������� �� ����� ������ ��������� ������ ���������� ������ � ����� ������
    FormDifVarAns.ans = ans
    FormDifVarAns.numAns = TextBoxNumDifAns.Value
    
    '�������� ������� ����� � ���������� ����� ������ ���������
    FormDifAns.Hide
    FormDifVarAns.Show
    
End Sub


Private Sub help_Click()
    HelpForm.refer = "������� ���������� ����� �������." + vbLf + vbLf _
                   + "����� ��������� ��������� ���� ���������� ��������, ����� ������ ����� � ����� ������. " _
                   + "(������� �� ����� ������ �������������� �������� ��������� ������ Shift+Enter.)" + vbLf + vbLf _
                   + "��� ������� �� ������ ""�������� ��������"" ������ �� �������� ���� ��������� � ������ � ������." _
                   + "������ ""������� ����������"" ������� ������ �� ������� ���� ���������� ���������." + vbLf + vbLf _
                   + "����� � ������ ���� ����� ������� ���������� �������� �������, ������� �� ��� �����." + vbLf + vbLf _
                   + "������� ������ �� ������� ����� ���� �� ������� ������ ��� ��������� ���������� �������."
    HelpForm.Show
End Sub

Private Sub Label2_Click()

End Sub

Private Sub UserForm_Click()

End Sub
