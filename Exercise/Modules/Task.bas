Attribute VB_Name = "Task"
Public Const MESAGE As String = "Данная версия программы более не поддерживается, обратитесь к разработчику или скачайте новую версию"
Public Const TITLE As String = "Версия программы 1.0 alfa"
Public TotalScore As Integer
' Глобальные переменные
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
        ' Выставляем курсор в конец документа
        Selection.EndKey Unit:=wdStory, Extend:=wdMove
        
        ' запускаем форму
        FormOneAnswer.Show
    Else
        MsgBox Prompt:=MESAGE, TITLE:=TITLE
    End If
End Sub

Sub MultiAnswer()
    If check Then
        ' Выставляем курсор в конец документа
        Selection.EndKey Unit:=wdStory, Extend:=wdMove

        ' запускаем форму
        FormDifAns.Show
        'FormMultiAnsVarAns.Show
    Else
        MsgBox Prompt:=MESAGE, TITLE:=TITLE
    End If
    
End Sub

Sub OpenAnswer()
    If check Then
        ' Выставляем курсор в конец документа
        Selection.EndKey Unit:=wdStory, Extend:=wdMove

        ' запускаем форму
        FormOpenAns.Show
    Else
        MsgBox Prompt:=MESAGE, TITLE:=TITLE
    End If
    
End Sub

Sub miniEssay()
    If check Then
        ' Выставляем курсор в конец документа
        Selection.EndKey Unit:=wdStory, Extend:=wdMove

        ' запускаем форму
        FormEssay.Show
    Else
        MsgBox Prompt:=MESAGE, TITLE:=TITLE
    End If
    
End Sub

Sub Substitution()
    If check Then
        ' Выставляем курсор в конец документа
        Selection.EndKey Unit:=wdStory, Extend:=wdMove

        ' запускаем форму
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

' Возвращает слово "Балл" в правильном падеже,
' если переданная строка не может преобразоваться в чилсо возвращается "error"
Function Num(x As String) As String
    Dim y As Integer
    On Error GoTo Bag
    If CInt(x) > 4 And CInt(x) < 20 Then
        Num = "баллов"
    Else
        y = CInt(Right(x, 1))
        If y = 1 Then
            Num = "балл"
        ElseIf y > 1 And y < 5 Then
            Num = "балла"
        ElseIf y > 4 And y < 10 Or y = 0 Then
            Num = "баллов"
        Else
            Num = "балла"
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



' Проверка на ввод корректных последовательностей в полях
Function SequenceError(sequence As TextBox, standard As TextBox, margin As String) As Boolean

    For Each elem In Split(sequence.Text, ", ")
        On Error GoTo tEr
        If CInt(standard.Value) < CInt(elem) Then
            MsgBox "Номер в поле " + margin + " превышает количество вариантов ответов!"
            SequenceError = True
            Exit Function
        End If
        
        ' Проверка на случайный ввод дробного числа
        If CDbl(elem) - CInt(elem) > 0 Then
            MsgBox "Пропущен пробел после запятой в поле " + margin + " !"
            SequenceError = True
            Exit Function
        End If
tEr:
        If Err.Number = 13 Then
            MsgBox "Неверно задана последовательность в поле " + margin + " !"
            SequenceError = True
            Exit Function
        End If
    Next elem
    SequenceError = False
    
End Function

' Проверяем входит ли одна последовательность в другую
Function NotMatch(parent As String, child As String)
    Dim flag As Integer
    flag = 0 ' Не совпадают
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
