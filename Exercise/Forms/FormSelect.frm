VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSelect 
   Caption         =   "������������ ��������� �������"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   8880.001
   OleObjectBlob   =   "FormSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DinamicArrAns() As String
Dim numElemDinamicArr As Integer ' ��� �������� �� ���������� ������ � �������� ���������� �������������������
Dim numberVarAnsSubst() As Integer
Dim ex As Integer ' ���������� ��� ��������������� ������� ������� VarAnsSubst

Private Sub CommandButtonClose_Click()
    Unload FormSelect
End Sub

'�������� ���������� �� ������� ����� � �� �� ��������� ��������� ������� ������, � ��������� ��� � �����
Private Sub CommandButtonEnter_Click()
    Dim n As Integer
    
    '�������� ���� ��������� ������� ������ �� �������
    If ListBoxRight.ListCount = 0 Then
        MsgBox "���� ������� �������� �����!"
        Exit Sub
    End If
    
    ' �������� �� ���������� ���� ������
    If Num(TextBoxScore.Value) = "error" Then
        MsgBox "���������� ������ �� ������������� ������!"
        Exit Sub
    End If
    
    Dim str As String
    str = ""
    
    '�������� (�������) =============================
'
'    For i = 0 To ListBoxRight.ListCount - 1 Step 1
'        For k = 0 To ListBoxMy.ListCount - 1 Step 1
'            If ListBoxMy.List(k) = ListBoxRight.List(i) Then
'                n = k + 1
'                Exit For
'            End If
'        Next k
'        If i = 0 Then
'            str = CStr(n) + "- " + ListBoxRight.List(i)
'        Else
'            str = str + ";" + CStr(n) + "- " + ListBoxRight.List(i)
'        End If
'    Next i
    
    
    '======================================
    
    For i = 0 To ListBoxRight.ListCount - 1 Step 1
'        For k = 0 To ListBoxLeft.ListCount - 1 Step 1
'            If ListBoxLeft.List(k) = ListBoxRight.List(i) Then n = k + 1
'        Next k
        If i = 0 Then
            str = ListBoxRight.List(i)
        Else
            str = str + ";" + ListBoxRight.List(i)
        End If
    Next i
    
    '---- �������� �� ���������� ������ � �������� ���������� �������������������----
    ReDim Preserve DinamicArrAns(numElemDinamicArr)
    If UBound(DinamicArrAns) > 0 Then
        For Each mmm In DinamicArrAns
            If str = mmm Then
                MsgBox "����� ����� ��� ������������!"
                Exit Sub
            End If
        Next mmm
    End If
    DinamicArrAns(numElemDinamicArr) = str
    numElemDinamicArr = numElemDinamicArr + 1
    '-------------------------------------------------------------------------------
    
    TextBoxAnswers.Text = TextBoxAnswers.Text + "### (" + TextBoxScore.Value + " " + Num(TextBoxScore.Value) + ") " + str + vbLf
    ListBoxRight.Clear
    TextBoxScore.Value = ""
    'ex = 0
    
End Sub

Private Sub CommandButtonIns_Click()
    Selection.TypeText Task.textSubst + TextBoxAnswers.Text + vbLf + vbLf
    Unload FormSelect
End Sub
' ������������ ������ � ������ ����
Private Sub CommandButtonMove_Click()
'    ReDim Preserve numberVarAnsSubst(ex)
'
'    If UBound(numberVarAnsSubst) > 0 Then
'        For Each Item In numberVarAnsSubst
'            If Item = ListBoxLeft.ListIndex Then
'                MsgBox "��������� ������� ��� �������� � ������!"
'                Exit Sub
'            End If
'        Next Item
'    End If

    If ListBoxRight.ListCount > 0 Then
        For Each Item In ListBoxRight.List
            If Item = (CStr(ListBoxLeft.ListIndex + 1) + "- " + ListBoxLeft.Value) Then
                MsgBox "��������� ������� ��� �������� � ������!"
                Exit Sub
            End If
        Next Item
    End If
    
    'numberVarAnsSubst(ex) = ListBoxLeft.ListIndex
    ListBoxRight.AddItem CStr(ListBoxLeft.ListIndex + 1) + "- " + ListBoxLeft.Value
    'ex = ex + 1
End Sub

Private Sub CommandButtonRemove_Click()
    ListBoxRight.Clear
End Sub

Private Sub help_Click()
    HelpForm.refer = "� ����� ���� ����������� ��� ����������� ������� �� ����� �� ���������� �����." + vbLf + vbLf _
                   + "������� ������ � ����� ���� ���������� �� � ������ ������� ���� � ������� ������ "">>>"" " _
                   + "� ������� ������������������ ��������������� ������������������ ��������� ""#___#"" � ������ �������." + vbLf + vbLf _
                   + "����� � ���� ""����"" ������� ���� ����������� �� �������������� ������������������, � ������� ""�������� �������""" + vbLf + vbLf _
                   + "� ������ ������ ���� ����� ������������ �������� ������������������� ������ �������." + vbLf + vbLf _
                   + "������ ""������� ������"" ������� � �������� Word ��� ������������ � ���� ""�������� �������"" ������������������."
    HelpForm.Show
End Sub



Private Sub UserForm_Initialize()
    numElemDinamicArr = 0
    ex = 0
End Sub
