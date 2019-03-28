VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormOneAnswer 
   Caption         =   "Один ответ"
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

'удаляет выделенный элемент
Private Sub CommandButtonDel_Click()
    ListBoxVarAnsOne.RemoveItem ListBoxVarAnsOne.ListIndex
End Sub

' Добавление ответов в нижнее окно
Private Sub CommandButtonInsert_Click()
    'если текстовое поле не пустое
    If TextBoxVarAnsOne.Text <> "" Then
        Dim ArrAns() As String
        Dim j As Integer
        k = 0
        ArrAns = Split(TextBoxVarAnsOne.Text, vbLf)
        
        ' /////Проверка на одинаковые варианты ответов////
        If Not Duty.NoIdentical(ArrAns) Then
            MsgBox "Не может быть двух одинаковых ответов!"
            TextBoxVarAnsOne.SetFocus
            Exit Sub
        End If
        For h = 0 To UBound(ArrAns)
        For t = 0 To ListBoxVarAnsOne.ListCount - 1
            If h = UBound(ArrAns) Then
                If ArrAns(h) = ListBoxVarAnsOne.List(t) Then
                    MsgBox "Не может быть двух одинаковых ответов!"
                    TextBoxVarAnsOne.SetFocus
                    Exit Sub
                End If
            Else
                If Duty.NotEndSimbol(CStr(ArrAns(h))) = ListBoxVarAnsOne.List(t) Then
                    MsgBox "Не может быть двух одинаковых ответов!"
                    TextBoxVarAnsOne.SetFocus
                    Exit Sub
                End If
            End If
        Next t
        Next h
        '////////////////////////////////////////////////
            
        
        ' переносим список из textBox в listBox
        For Each Item In ArrAns
            If j = UBound(ArrAns) Then
                If Item <> "" Then
                If Item <> " " Then
                    ListBoxVarAnsOne.AddItem Item
                End If
                End If
            Else
                If Left(Item, Len(Item) - 1) <> "" Then  ' пропускаем пустые строки в массиве
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
    ' проверка на пустое значение полей
    If TextBoxNumAnsOne.Value = "" Or TextBoxScoreOneAns.Value = "" Or ListBoxVarAnsOne.ListCount < 1 Then
        MsgBox "В данной форме поля не могут быть пустыми!"
        Exit Sub
    End If
    
    ' Проверка на количество вариантов ответов, их должно быть от 2 до 7
    If ListBoxVarAnsOne.ListCount < 2 Then
            MsgBox "Вариантов ответов должно быть больше одного!"
            TextBoxVarAnsOne.SetFocus
            Exit Sub
    ElseIf ListBoxVarAnsOne.ListCount > 7 Then
            MsgBox "Количество вариантов ответов не должно превышать семи!"
            TextBoxVarAnsOne.SetFocus
            Exit Sub
    End If
    
    ' Проверка на корректный ввод баллов
    If Num(TextBoxScoreOneAns.Value) = "error" Then
        MsgBox "Количество баллов не представленно числом!"
        Exit Sub
    End If
    
    ' Проверка на то, что выбран вариант правильного ответа
    If ListBoxVarAnsOne.ListIndex < 0 Then
        MsgBox "Не выбран правильный ответ!"
        Exit Sub
    End If
    

    
    ' Если ошибок не было
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
    
    
    '' Вывод шаблона задания в документ
    Selection.TypeText _
    "== Задание " + TextBoxNumAnsOne.Value + " ==" + vbLf _
    + "Впишите сюда текст задания." + vbLf _
    + "=== Ответы (единственный выбор) ===" + vbLf _
    + str _
    + "=== Оценки ===" + vbLf _
    + "### (" + TextBoxScoreOneAns.Value + " " + Num(TextBoxScoreOneAns.Value) + ") " + CStr(ListBoxVarAnsOne.ListIndex + 1) + vbLf 'TextBoxRightOneAns.Value
    
    ' Добавляем пустых строчек
    Selection.TypeText vbLf + vbLf
    
    Unload FormOneAnswer

End Sub


Private Sub helpOne_Click()
    'Вызов справки
    HelpForm.refer = "Введите порядковый номер вопроса." + vbLf + vbLf _
                          + "Далее перечислите варианты ответов в верхнем текстовом поле, вводя каждый отдельный ответ на новой строчке." _
                          + "Для перехода на новую строку используйте сочетание клавиш Shift+Enter." + vbLf + vbLf _
                          + "Нажатие не кнопке ""Добавить варианты"" добавляет текст из верхнего окна в нижнее." _
                          + "Нажатие не кнопке ""Удалить выделенное"" удалит выделенный элемент из нижнего окна" _
                          + "По завершении формирования списков ответов в нижнем окне необходимо выбрать правильный ответ, щёлкнув на нем мышью." + vbLf + vbLf _
                          + "Числом укажите количество баллов, начисляемых за верный ответ." + vbLf + vbLf _
                          + "После нажатия на кнопку ""Вставить шаблон"" в конец текста документа будет выведен шаблон." + vbLf + vbLf _
                          + "После чего необходимо отредактировать шаблон непосредственно в редакторе Word, вписав текст задания." + vbLf + vbLf _
                          + "Закрытие формы нажатием на кнопку ""Закрыть форму"" приведет к ее закрытию без сохранения информации."
    HelpForm.Show
End Sub



Private Sub UserForm_Click()

End Sub
