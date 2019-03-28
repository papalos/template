VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormOpenAns 
   Caption         =   "Открытый вопрос"
   ClientHeight    =   4632
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   5520
   OleObjectBlob   =   "FormOpenAns.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormOpenAns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public varAnsStr As String
Public varAnsNum As Integer



Private Sub CommandButtonOpenClose_Click()
    ' Проверка на то что хотябы однажды была нажата кнопка Записать,
    ' для сохранения хотябы одного варианта
    If varAnsStr = "" Then
        MsgBox "Прежде нажмите кнопку ""Записать вариант"", для сохранения варианта ответа!"
        Exit Sub
    End If
    
    Selection.TypeText _
    "== Задание " + TextBoxOpenNum.Value + " ==" + vbLf _
    + "Впишите сюда текст задания." + vbLf _
    + "=== Оценки ===" + vbLf _
    + varAnsStr
    
    ' Добавляем пустых строчек
    Selection.TypeText vbLf + vbLf
    
    varAnsStr = ""
    varAnsNum = 0
    Unload FormOpenAns
End Sub


Private Sub CommandButtonOpenExit_Click()
    varAnsStr = ""
    varAnsNum = 0
    Unload FormOpenAns
End Sub

Private Sub CommandButtonOpenOK_Click()
    ' проверка на пустое значение полей
    If TextBoxOpenNum.Value = "" Or TextBoxOpenVarAns.Value = "" Or TextBoxOpenVarScore.Value = "" Then
        MsgBox "В данной форме поля не могут быть пустыми!"
        Exit Sub
    End If
    
    ' Проверка на корректный ввод баллов
    If Num(TextBoxOpenVarScore.Value) = "error" Then
        MsgBox "Количество баллов не представленно числом!"
        Exit Sub
    End If
    
    
   If Trim(TextBoxOpenVarAns.Value) = "" Then
        MsgBox "Эталонный ответ представлен пробелами!"
        Exit Sub
   End If
    
    ' Если ошибок не было
    varAnsStr = varAnsStr + "### (" + TextBoxOpenVarScore.Value + " " + Num(TextBoxOpenVarScore.Value) + ") " + TextBoxOpenVarAns.Value + vbLf
    varAnsNum = varAnsNum + 1
    TextBoxVarOpen.Text = "Вариант ответа №" + CStr(varAnsNum) + " сохранен!" + vbLf + varAnsStr
    TextBoxOpenNum.Enabled = False
    TextBoxOpenVarScore.Value = ""
    TextBoxOpenVarAns.Value = ""
End Sub

Private Sub help_Click()
    HelpForm.refer = "Введите порядковый номер вопроса." + vbLf + vbLf _
                   + "Далее введите вариант правильного ответа и количество баллов за него." + vbLf + vbLf _
                   + "Если ответов несколько, то после нажатия на кнопке ""Записать ответ"" введите следующий вариант и количество баллов за него." + vbLf + vbLf _
                   + "Также потребуются отдельные вводы если ответ один, но его форматы могут быть разными." + vbLf + vbLf _
                   + "Например, на вопрос: Сколько пассажиров было в лодке, не считая собаки?" + vbLf + vbLf _
                   + "Потребуется ввод следующих вариантов: 3 - 10 баллов, три - 10 баллов, трое - 10 баллов" + vbLf + vbLf _
                   + "Регистр букв при вводе ответа не учитывается, ""Яблоко"" и ""яблоко"" и ""ЯБЛОКО"" - для системы это одинаковые значения (см. инструкцию по оформлению заданий)." + vbLf + vbLf _
                   + "Нажатие на кнопку ""Записать ответ"" ведет к записи варианта ответа в память, вывода его в поле для отображения вариантов ответа, и предложению ввода нового варианта." _
                   + "(повторите операцию записи вариантов ответов необходимое число раз!)" + vbLf + vbLf _
                   + "Нажатие на кнопку ""Вставить шаблон"" приводит к выводу шаблона в текст документа и закрытию формы" + vbLf + vbLf _
                   + "Нажатие на кнопку ""Закрыть"" приводит к закрытию формы без вывода шаблона."
    HelpForm.Show
End Sub

Private Sub TextBoxOpenVarAns_Change()

End Sub

Private Sub UserForm_Click()

End Sub
