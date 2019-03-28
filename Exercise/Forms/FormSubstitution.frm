VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSubstitution 
   Caption         =   "Подстановка"
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

Dim substit As String        ' для сохранения текста при подсчете количества символов


Private Sub CommandButtonSubstitutionExit_Click()
    ' Выгрузка формы по нажатию кнопки выхода
    Unload FormSubstitution
End Sub

Private Sub CommandButtonSubstOK_Click()
    Dim counting As Integer
    Dim ArrText() As String
    Dim ArrNormalText() As String     ' массив бес пустых строк
    Dim gg As Integer                 ' счетчик массива без пустых строк
    counting = 0

    ' Проверка на пустые поля
    If TextBoxNumAnsSubst.Value = "" Or TextBoxSubstAns.Value = "" Then
        MsgBox "Необходимо заполнить все поля!"
        Exit Sub
    End If
    
    MyPos = InStr(TextBoxSubstAns.Value, ",")
    
    'Проверяем подстановки на остутствие запятых и точек с запятыми
    If InStr(TextBoxSubstAns.Value, ",") > 0 Then
        MsgBox "Символ запятой не допускается в подстановке"
        Exit Sub
    End If
    If InStr(TextBoxSubstAns.Value, ";") > 0 Then
        MsgBox "Символ точки с запятой не допускается в подстановке"
        Exit Sub
    End If
    
    
    Dim str As String, strIns As String, sum As String
    Dim strArr() As String
    
    str = ""
    strIns = ""
    
    ' Разбиваем содержимое текстового поля на массив по символу конца строки
    ArrText = Split(TextBoxSubstAns.Text, vbLf)
    
    If Not Duty.NoIdentical(ArrText) Then
        MsgBox "Не может быть двух одинаковых ответов!"
        TextBoxSubstAns.SetFocus
        Exit Sub
    End If
    
    ' Если чекбокс суммирования включон добавляем слово "суммировать" в текст
    If CheckBoxSum.Value Then
        sum = " (суммировать)"
    Else
        sum = ""
    End If

    
    ' Новый массив без пустых элементов
    For kk = 0 To UBound(ArrText) Step 1                                 ' перебираем все элементы массива ArrPasses
        
        If kk = UBound(ArrText) Then                                     ' если это последний элемент
            If Trim(ArrText(kk)) <> "" Then                              ' проверяем его на пустую строку
                ReDim Preserve ArrNormalText(gg)
                ArrNormalText(gg) = ArrText(kk)                          ' и записываем в массив  ArrNormalText
                gg = gg + 1                                                ' увеличиваем счетчик массива ArrNormalText
            End If
        Else                                                               ' если это не последний элемент
            If Trim(Duty.NotEndSimbol(CStr(ArrText(kk)))) <> "" Then     ' отрезаем от него послединй символ конца строки убираем пробелы и проверяем на пустоту
                ReDim Preserve ArrNormalText(gg)
                ArrNormalText(gg) = Duty.NotEndSimbol(CStr(ArrText(kk)))
                gg = gg + 1
            End If
        End If
    Next kk
    
    ' перенос массива без пустых строк в список
    For Each elem In ArrNormalText                     ' Перебираем все элементы полученного массива
        counting = counting + 1
        ' Если это последний элемент передаем его таким какой он есть, если нет отрезаем от него символ конца строки
        'If counting = UBound(ArrNormalText) + 1 Then
            str = str + "# " + elem + vbLf             ' для вывода в шаблон
            FormSubstChart.ListBoxAns.AddItem (elem)   ' для передачи в форму выбора правильных ответов
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
    "== Задание " + TextBoxNumAnsSubst.Value + " ==" + vbLf _
    + "Сюда вписывается текст задания. Например, сюда #___# необходимо вписать слово яблоко, а сюда #___# - апельсин." + vbLf _
    + "=== Подстановки ===" + vbLf _
    + strIns _
    + "=== Оценки" + sum + " ===" + vbLf
    
    
    Unload FormSubstitution
    
    FormSubstChart.Show
    
    
End Sub

Private Sub help_Click()
    HelpForm.refer = "В форме вводим порядковый номер вопроса." + vbLf + vbLf _
                   + "В расположенном ниже окне необходимо перечислить все варианты для подстановок, " _
                   + "вводя каждый новый ответ на новой строке, используя для перехода на новую строку сочетание клавиш Shift+Enter." + vbLf + vbLf _
                   + "Если требуется суммирование баллов при каждом правильном ответе " _
                   + "(требуется, когда баллы необходимо суммировать при каждой верной подстановке, " _
                   + "невзирая на количество неверно заполненных или вовсе незаполненных пропусков), " _
                   + "необходимо поставить галочку ""Суммировать ответы""." + vbLf + vbLf _
                   + "Если же предполагаются строгие ответы галочку ставить ненужно (см. инструкцию по оформлению заданий)." + vbLf + vbLf _
                   + "Нажатие кнопки ""ОК"" приведет к выводу части шаблона в конец документа " _
                   + "и открытию окна ""Формирование вариантов ответов""." + vbLf + vbLf
    HelpForm.Show
End Sub

' объем всех подстановок не должен превышать 1000 символов
Private Sub TextBoxSubstAns_Change()
    If Len(TextBoxSubstAns.Text) > 1000 Then
        MsgBox "Вы превысили лимит, отведенный на подстановки!"
        TextBoxSubstAns.Text = substit
    End If
    substit = TextBoxSubstAns.Text
End Sub


Private Sub UserForm_Click()

End Sub
