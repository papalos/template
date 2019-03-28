VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormDifVarAns 
   Caption         =   "Выбор правильных вариантов ответов"
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
    ' Проверка на корректный ввод баллов (Num - проверяет число ли это)
    If Num(TextBoxVarScore.Value) = "error" Then
        MsgBox "Количество баллов не представленно числом!"
        Exit Sub
    End If
    
    ' Проверка на выбор хотя бы одного правильного ответа
    For Z = 0 To ListBoxAllRightAns.ListCount - 1 Step 1
        If ListBoxAllRightAns.Selected(Z) Then
            Z = -1
            Exit For
        End If
    Next Z
    If Z <> -1 Then
        MsgBox "Невыбрано ни одного ответа для варианта!"
        Exit Sub
    End If
    
    ' если ошибки нет
    Dim numbersOfAnsver As String
    numbersOfAnsver = ""
    ' Если список правильных ответов не пуст
    If ListBoxAllRightAns.ListCount <> 0 Then
        For p = 0 To ListBoxAllRightAns.ListCount - 1 Step 1
            If ListBoxAllRightAns.Selected(p) Then
                ' Начинаем перебирать список всех ответов на форме FormDifAns
                For i = 0 To FormDifAns.ListBoxRight.ListCount - 1 Step 1
                    ' и сравнивать их с выбраным значением на форме FormDifVarAns
                    If FormDifAns.ListBoxRight.List(i) = ListBoxAllRightAns.List(p) Then
                        ' если находим берем его порядковый номер из общего списка всех вопросов
                        numbersOfAnsver = numbersOfAnsver + CStr(i + 1) + ", "
                    End If
                Next i
            End If
        Next p

        If numbersOfAnsver <> "" Then numbersOfAnsver = Left(numbersOfAnsver, Len(numbersOfAnsver) - 2)
        varAnsStr = varAnsStr + "### (" + TextBoxVarScore.Value + " " + Num(TextBoxVarScore.Value) + ") " + numbersOfAnsver + vbLf
        TextBoxRightAns.Text = varAnsStr + vbLf
    End If
    
    'стираем значение в окне баллов и снимаем выделение в списке ответов
    TextBoxVarScore.Value = ""
    For i = 0 To ListBoxAllRightAns.ListCount - 1
        ListBoxAllRightAns.Selected(i) = False
    Next
    
    
End Sub

Private Sub CommandButtonExit_Click()
    If varAnsStr = "" Then
        MsgBox "Варианты правильных ответов не заданы!"
        Exit Sub
    End If

    ' Выводим шаблон в документ
    Selection.TypeText _
    "== Задание " + numAns + " ==" + vbLf _
    + "Впишите сюда текст задания." + vbLf _
    + "=== Ответы (множественный выбор) ===" + vbLf _
    + ans _
    + "=== Оценки ===" + vbLf _
    + varAnsStr
    
    ' Добавляем пустых строчек
    Selection.TypeText vbLf + vbLf
    
    Unload FormDifAns
    Unload FormDifVarAns
End Sub


Private Sub help_Click()
    HelpForm.refer = "Выберите последовательность правильных ответов, введите количество баллов начисляемых за эту последовательность и нажмите кнопку ""Записать ответ""." + vbLf + vbLf _
                   + "В нижнем поле отобразится вариант ответа с указанием баллов за него." + vbLf + vbLf _
                   + "Вы можете продолжить вводить комбинации верных ответов и сохранять их нажатием кнопки ""Записать ответ""." + vbLf + vbLf _
                   + "Нажатие по кнопке ""Вывести шаблон"" приведет к закрытию формы и выводу шаблона в конец документа Word." + vbLf + vbLf _
                   + "В последующем необходимо внести соответствующие изменения в шаблон задания, вписав текст задания непосредственно в документе Word." + vbLf + vbLf _
                   + "Нажатие на кнопку ""Закрыть форму"" приведет к закрытию формы без сохранения информации и шаблон выведен не будет."
    HelpForm.Show
End Sub

Private Sub UserForm_Click()

End Sub
