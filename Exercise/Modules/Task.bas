Attribute VB_Name = "Task"
Public Const MESAGE As String = "������ ������ ��������� ����� �� ��������������, ���������� � ������������ ��� �������� ����� ������"
Public Const TITLE As String = "������ ��������� 1.0 alfa"
Public TotalScore As Integer
' ���������� ����������
Public textSubst As String


Sub FormTitleOpen()
    If check Then
        FormTitle.Show
    Else
        MsgBox Prompt:=MESAGE, TITLE:=TITLE
    End If
End Sub

Sub OneAnswer()
    If check Then
        ' ���������� ������ � ����� ���������
        Selection.EndKey Unit:=wdStory, Extend:=wdMove
        
        ' ��������� �����
        FormOneAnswer.Show
    Else
        MsgBox Prompt:=MESAGE, TITLE:=TITLE
    End If
End Sub

Sub MultiAnswer()
    If check Then
        ' ���������� ������ � ����� ���������
        Selection.EndKey Unit:=wdStory, Extend:=wdMove

        ' ��������� �����
        FormDifAns.Show
        'FormMultiAnsVarAns.Show
    Else
        MsgBox Prompt:=MESAGE, TITLE:=TITLE
    End If
    
End Sub

Sub OpenAnswer()
    If check Then
        ' ���������� ������ � ����� ���������
        Selection.EndKey Unit:=wdStory, Extend:=wdMove

        ' ��������� �����
        FormOpenAns.Show
    Else
        MsgBox Prompt:=MESAGE, TITLE:=TITLE
    End If
    
End Sub

Sub miniEssay()
    If check Then
        ' ���������� ������ � ����� ���������
        Selection.EndKey Unit:=wdStory, Extend:=wdMove

        ' ��������� �����
        FormEssay.Show
    Else
        MsgBox Prompt:=MESAGE, TITLE:=TITLE
    End If
    
End Sub

Sub Substitution()
    If check Then
        ' ���������� ������ � ����� ���������
        Selection.EndKey Unit:=wdStory, Extend:=wdMove

        ' ��������� �����
        FormSubstitution.Show
    Else
        MsgBox Prompt:=MESAGE, TITLE:=TITLE
    End If

End Sub

Sub Skip()
    If check Then
        Selection.TypeText " #___# "
    Else
        MsgBox Prompt:=MESAGE, TITLE:=TITLE
    End If
End Sub

Sub Passes()
    If check Then
        FormPass.Show
    Else
        MsgBox Prompt:=MESAGE, TITLE:=TITLE
    End If

End Sub

' ���������� ����� "����" � ���������� ������,
' ���� ���������� ������ �� ����� ��������������� � ����� ������������ "error"
Function Num(x As String) As String
    Dim y As Integer
    On Error GoTo Bag
    If CInt(x) > 4 And CInt(x) < 20 Then
        Num = "������"
    Else
        y = CInt(Right(x, 1))
        If y = 1 Then
            Num = "����"
        ElseIf y > 1 And y < 5 Then
            Num = "�����"
        ElseIf y > 4 And y < 10 Or y = 0 Then
            Num = "������"
        Else
            Num = "�����"
        End If
    End If
Bag:
    If Err.Number = 13 Then
        Num = "error"
    End If
End Function

Function check() As Boolean
    If Date < CDate("01.11.2020") Then
        check = True
    Else
        check = False
    End If
End Function



' �������� �� ���� ���������� ������������������� � �����
Function SequenceError(sequence As TextBox, standard As TextBox, margin As String) As Boolean

    For Each elem In Split(sequence.Text, ", ")
        On Error GoTo tEr
        If CInt(standard.Value) < CInt(elem) Then
            MsgBox "����� � ���� " + margin + " ��������� ���������� ��������� �������!"
            SequenceError = True
            Exit Function
        End If
        
        ' �������� �� ��������� ���� �������� �����
        If CDbl(elem) - CInt(elem) > 0 Then
            MsgBox "�������� ������ ����� ������� � ���� " + margin + " !"
            SequenceError = True
            Exit Function
        End If
tEr:
        If Err.Number = 13 Then
            MsgBox "������� ������ ������������������ � ���� " + margin + " !"
            SequenceError = True
            Exit Function
        End If
    Next elem
    SequenceError = False
    
End Function

' ��������� ������ �� ���� ������������������ � ������
Function NotMatch(parent As String, child As String)
    Dim flag As Integer
    flag = 0 ' �� ���������
    For Each x In Split(child, ", ")
        For Each y In Split(parent, ", ")
            If x = y Then
                flag = flag + 1
                Exit For
            End If
        Next y
    Next x
    If UBound(Split(child, ", ")) + 1 = flag Then
        NotMatch = False
    Else
        NotMatch = True
    End If
    
End Function
