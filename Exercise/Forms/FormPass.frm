VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPass 
   Caption         =   "��������"
   ClientHeight    =   5916
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   7992
   OleObjectBlob   =   "FormPass.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim foll As Boolean
Dim follNum As Integer
Dim substit As String        ' ��� ���������� ������ ��� �������� ���������� ��������


Private Sub CheckBoxPassSum_Change()
    If CheckBoxPassSum.Value = True Then
        Label3.Caption = "��� ������������ �������, ������� �� ������ ����������� ������ �� ��� � ��������� ������������������, ����������� ���� �������� ������ ""�������� �����"""
        TextBoxPasses.MultiLine = False
        foll = False
    Else
        Label3.Caption = "������� ���������� ������� ������, ���������� ���������� �����������. ������� ������ ������������������ �����������, ����� ������ ����������� �� ��������� ������." _
                          + vbLf + "(��� �������� �� ����� ������ ����������� ��������� ������ shift+enter)"
        TextBoxPasses.MultiLine = True
        foll = True
    End If
End Sub


Private Sub CommandButtonAddPass_Click()
    Dim ArrText() As String                     '(���������) �� ������� ��� ������������
    Dim ArrNormalText() As String
    Dim insPass As String
    Dim cnt As Integer
    Dim jj As Integer
    
    ' �������� �� ������ �������� �����
    If TextBoxPasses.Value = "" Or TextBoxPassScore.Value = "" Or TextBoxNumAnsPass.Value = "" Then
        MsgBox "� ������ ����� ���� �� ����� ���� �������!"
        Exit Sub
    End If
    
    ' �������� �� ���������� ���� ������
    If Num(TextBoxPassScore.Value) = "error" Then
        MsgBox "���������� ������ �� ������������� ������!"
        Exit Sub
    End If
    
    '��������� ����������� �� ���������� ������� � ����� � ��������
    If InStr(TextBoxPasses.Value, ",") > 0 Then
        MsgBox "������ ������� �� ����������� � �����������"
        Exit Sub
    End If
    If InStr(TextBoxPasses.Value, ";") > 0 Then
        MsgBox "������ ����� � ������� �� ����������� � �����������"
        Exit Sub
    End If
    
    cnt = 0
    ' ��������� ����� � ������ ����� �� ������� ����� ������
    ArrPasses = Split(TextBoxPasses.Text, vbLf)
    
    ' ����� ������ ��� ������ ���������
    For ii = 0 To UBound(ArrPasses) Step 1                                 ' ���������� ��� �������� ������� ArrPasses
        
        If ii = UBound(ArrPasses) Then                                     ' ���� ��� ��������� �������
            If Trim(ArrPasses(ii)) <> "" Then                              ' ��������� ��� �� ������ ������
                ReDim Preserve ArrNormalText(jj)
                ArrNormalText(jj) = ArrPasses(ii)                          ' � ���������� � ������  ArrNormalText
                jj = jj + 1                                                ' ����������� ������� ������� ArrNormalText
            End If
        Else                                                               ' ���� ��� �� ��������� �������
            If Trim(Duty.NotEndSimbol(CStr(ArrPasses(ii)))) <> "" Then     ' �������� �� ���� ��������� ������ ����� ������ ������� ������� � ��������� �� �������
                ReDim Preserve ArrNormalText(jj)
                ArrNormalText(jj) = Duty.NotEndSimbol(CStr(ArrPasses(ii)))
                jj = jj + 1
            End If
        End If
    Next ii
        
    
    If foll Then
        ' ��������� �������� ������� �������� �� � ������ ����� ������� � ��������� ����������� ������
        cnt = 0
        For cnt = 0 To UBound(ArrNormalText)
            If cnt = 0 Then
                insPass = insPass + CStr(cnt + 1) + "- " + ArrNormalText(cnt)
            Else
                insPass = insPass + ";" + CStr(cnt + 1) + "- " + ArrNormalText(cnt)
            End If
        Next cnt
    Else
        follNum = follNum + 1
        If TextBoxPasses.Value <> "" Then
        If TextBoxPasses.Value <> " " Then
            insPass = insPass + CStr(follNum) + "- " + TextBoxPasses.Value
        End If
        End If
    End If

    
    
    
    TextBoxPassAnswers.Text = TextBoxPassAnswers.Text + "### (" + TextBoxPassScore.Value + " " + Num(TextBoxPassScore.Value) + ") " + insPass + vbLf
    TextBoxPassScore.Value = ""
    TextBoxPasses.Value = ""
    TextBoxNumAnsPass.TabStop = False
    CheckBoxPassSum.Enabled = False
End Sub

Private Sub CommandButtonIns_Click()
    Dim sum As String
    
    ' ��������� ����� �� ������� ����������� ���� ��, ��������� ����� ����������� � �����
    If CheckBoxPassSum.Value Then
        sum = " (�����������)"
    Else
        sum = ""
    End If

    Selection.TypeText _
    "== ������� " + TextBoxNumAnsPass.Value + " ==" + vbLf _
    + "���� ����������� ����� �������. ��������, ���� #___# ���������� ������� ����� ������, � ���� #___# - ��������." + vbLf _
    + "=== �������� ===" + vbLf _
    + "=== ������" + sum + " ===" + vbLf

    Selection.TypeText TextBoxPassAnswers.Text + vbLf + vbLf
    TextBoxPassAnswers.Text = ""
    Unload FormPass
End Sub

Private Sub CommandButtonPassCancel_Click()
    Unload FormPass
End Sub

Private Sub help_Click()
    HelpForm.refer = "� ���� ���� ���������� ������������ �������� ����������� � ������� ��������������� ��������� � �������." + vbLf + vbLf _
                   + "� ������� ���� ������� ������������������ �����������, ������� ���� �� ��� � ������� ""�������� �����"". " _
                   + "������������������ ����������� � ������ ���� ""�������� �������""" + vbLf + vbLf _
                   + "�������� ����������� ���������� �������������������, �������� ���� �� ������, ����������� ���� �������� ������ ""�������� �����""" + vbLf + vbLf _
                   + "������� � �������� ""�����������"" �������� � ���, ��� ����� �� ��������� ������ ������ �������������" + vbLf + vbLf _
                   + "������� �� ������ ""������� ������"" �������� � ������ ������������������� ��������� � ���� ""�������� �������"" � ����� ��������� Word  � �������� ����." + vbLf + vbLf
    HelpForm.Show
End Sub


Private Sub TextBoxPasses_Change()
    ' ����� ���� ����������� �� ������ ��������� 1000 ��������
    If Len(TextBoxPasses.Text) > 1000 Then
        MsgBox "�� ��������� �����, ���������� �� �����������!"
        TextBoxPasses.Text = substit
    End If
    substit = TextBoxPasses.Text
End Sub

Private Sub UserForm_Initialize()
    foll = True
    follNum = 0
End Sub
