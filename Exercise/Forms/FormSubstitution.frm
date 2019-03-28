VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSubstitution 
   Caption         =   "�����������"
   ClientHeight    =   4224
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6600
   OleObjectBlob   =   "FormSubstitution.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSubstitution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim substit As String        ' ��� ���������� ������ ��� �������� ���������� ��������


Private Sub CommandButtonSubstitutionExit_Click()
    ' �������� ����� �� ������� ������ ������
    Unload FormSubstitution
End Sub

Private Sub CommandButtonSubstOK_Click()
    Dim counting As Integer
    Dim ArrText() As String
    Dim ArrNormalText() As String     ' ������ ��� ������ �����
    Dim gg As Integer                 ' ������� ������� ��� ������ �����
    counting = 0

    ' �������� �� ������ ����
    If TextBoxNumAnsSubst.Value = "" Or TextBoxSubstAns.Value = "" Then
        MsgBox "���������� ��������� ��� ����!"
        Exit Sub
    End If
    
    MyPos = InStr(TextBoxSubstAns.Value, ",")
    
    '��������� ����������� �� ���������� ������� � ����� � ��������
    If InStr(TextBoxSubstAns.Value, ",") > 0 Then
        MsgBox "������ ������� �� ����������� � �����������"
        Exit Sub
    End If
    If InStr(TextBoxSubstAns.Value, ";") > 0 Then
        MsgBox "������ ����� � ������� �� ����������� � �����������"
        Exit Sub
    End If
    
    
    Dim str As String, strIns As String, sum As String
    Dim strArr() As String
    
    str = ""
    strIns = ""
    
    ' ��������� ���������� ���������� ���� �� ������ �� ������� ����� ������
    ArrText = Split(TextBoxSubstAns.Text, vbLf)
    
    If Not Duty.NoIdentical(ArrText) Then
        MsgBox "�� ����� ���� ���� ���������� �������!"
        TextBoxSubstAns.SetFocus
        Exit Sub
    End If
    
    ' ���� ������� ������������ ������� ��������� ����� "�����������" � �����
    If CheckBoxSum.Value Then
        sum = " (�����������)"
    Else
        sum = ""
    End If

    
    ' ����� ������ ��� ������ ���������
    For kk = 0 To UBound(ArrText) Step 1                                 ' ���������� ��� �������� ������� ArrPasses
        
        If kk = UBound(ArrText) Then                                     ' ���� ��� ��������� �������
            If Trim(ArrText(kk)) <> "" Then                              ' ��������� ��� �� ������ ������
                ReDim Preserve ArrNormalText(gg)
                ArrNormalText(gg) = ArrText(kk)                          ' � ���������� � ������  ArrNormalText
                gg = gg + 1                                                ' ����������� ������� ������� ArrNormalText
            End If
        Else                                                               ' ���� ��� �� ��������� �������
            If Trim(Duty.NotEndSimbol(CStr(ArrText(kk)))) <> "" Then     ' �������� �� ���� ��������� ������ ����� ������ ������� ������� � ��������� �� �������
                ReDim Preserve ArrNormalText(gg)
                ArrNormalText(gg) = Duty.NotEndSimbol(CStr(ArrText(kk)))
                gg = gg + 1
            End If
        End If
    Next kk
    
    ' ������� ������� ��� ������ ����� � ������
    For Each elem In ArrNormalText                     ' ���������� ��� �������� ����������� �������
        counting = counting + 1
        ' ���� ��� ��������� ������� �������� ��� ����� ����� �� ����, ���� ��� �������� �� ���� ������ ����� ������
        'If counting = UBound(ArrNormalText) + 1 Then
            str = str + "# " + elem + vbLf             ' ��� ������ � ������
            FormSubstChart.ListBoxAns.AddItem (elem)   ' ��� �������� � ����� ������ ���������� �������
            FormSelect.ListBoxMy.AddItem (elem)
'        Else
'            str = str + "# " + Left(elem, Len(elem) - 1) + vbLf
'            FormSubstChart.ListBoxAns.AddItem (Left(elem, Len(elem) - 1))
'            FormSelect.ListBoxMy.AddItem (Left(elem, Len(elem) - 1))
'        End If
    Next elem
    
    strIns = str

    'Selection.TypeText
    Task.textSubst = _
    "== ������� " + TextBoxNumAnsSubst.Value + " ==" + vbLf _
    + "���� ����������� ����� �������. ��������, ���� #___# ���������� ������� ����� ������, � ���� #___# - ��������." + vbLf _
    + "=== ����������� ===" + vbLf _
    + strIns _
    + "=== ������" + sum + " ===" + vbLf
    
    
    Unload FormSubstitution
    
    FormSubstChart.Show
    
    
End Sub

Private Sub help_Click()
    HelpForm.refer = "� ����� ������ ���������� ����� �������." + vbLf + vbLf _
                   + "� ������������� ���� ���� ���������� ����������� ��� �������� ��� �����������, " _
                   + "����� ������ ����� ����� �� ����� ������, ��������� ��� �������� �� ����� ������ ��������� ������ Shift+Enter." + vbLf + vbLf _
                   + "���� ��������� ������������ ������ ��� ������ ���������� ������ " _
                   + "(���������, ����� ����� ���������� ����������� ��� ������ ������ �����������, " _
                   + "�������� �� ���������� ������� ����������� ��� ����� ������������� ���������), " _
                   + "���������� ��������� ������� ""����������� ������""." + vbLf + vbLf _
                   + "���� �� �������������� ������� ������ ������� ������� ������� (��. ���������� �� ���������� �������)." + vbLf + vbLf _
                   + "������� ������ ""��"" �������� � ������ ����� ������� � ����� ��������� " _
                   + "� �������� ���� ""������������ ��������� �������""." + vbLf + vbLf
    HelpForm.Show
End Sub

' ����� ���� ����������� �� ������ ��������� 1000 ��������
Private Sub TextBoxSubstAns_Change()
    If Len(TextBoxSubstAns.Text) > 1000 Then
        MsgBox "�� ��������� �����, ���������� �� �����������!"
        TextBoxSubstAns.Text = substit
    End If
    substit = TextBoxSubstAns.Text
End Sub


Private Sub UserForm_Click()

End Sub
