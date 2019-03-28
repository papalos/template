VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPass 
   Caption         =   "Пропуски"
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
Dim substit As String        ' для сохранения текста при подсчете количества символов


Private Sub CheckBoxPassSum_Change()
    If CheckBoxPassSum.Value = True Then
        Label3.Caption = "При суммировании ответов, вводите по одному правильному ответу за раз в требуемой последовательности, подтверждая ввод нажатием кнопки ""Добавить ответ"""
        TextBoxPasses.MultiLine = False
        foll = False
    Else
        Label3.Caption = "Введите правильный вариант ответа, перечислив правильные подстановки. Укажите верную последовательность подстановок, вводя каждую подстановку на отдельной строке." _
                          + vbLf + "(для перехода но новую строку используйте сочетание клавиш shift+enter)"
        TextBoxPasses.MultiLine = True
        foll = True
    End If
End Sub


Private Sub CommandButtonAddPass_Click()
    Dim ArrText() As String                     '(проверить) не найдено где используется
    Dim ArrNormalText() As String
    Dim insPass As String
    Dim cnt As Integer
    Dim jj As Integer
    
    ' проверка на пустое значение полей
    If TextBoxPasses.Value = "" Or TextBoxPassScore.Value = "" Or TextBoxNumAnsPass.Value = "" Then
        MsgBox "В данной форме поля не могут быть пустыми!"
        Exit Sub
    End If
    
    ' Проверка на корректный ввод баллов
    If Num(TextBoxPassScore.Value) = "error" Then
        MsgBox "Количество баллов не представленно числом!"
        Exit Sub
    End If
    
    'Проверяем подстановки на остутствие запятых и точек с запятыми
    If InStr(TextBoxPasses.Value, ",") > 0 Then
        MsgBox "Символ запятой не допускается в подстановке"
        Exit Sub
    End If
    If InStr(TextBoxPasses.Value, ";") > 0 Then
        MsgBox "Символ точки с запятой не допускается в подстановке"
        Exit Sub
    End If
    
    cnt = 0
    ' Разбиваем текст в массив строк по символу конца строки
    ArrPasses = Split(TextBoxPasses.Text, vbLf)
    
    ' Новый массив без пустых элементов
    For ii = 0 To UBound(ArrPasses) Step 1                                 ' перебираем все элементы массива ArrPasses
        
        If ii = UBound(ArrPasses) Then                                     ' если это последний элемент
            If Trim(ArrPasses(ii)) <> "" Then                              ' проверяем его на пустую строку
                ReDim Preserve ArrNormalText(jj)
                ArrNormalText(jj) = ArrPasses(ii)                          ' и записываем в массив  ArrNormalText
                jj = jj + 1                                                ' увеличиваем счетчик массива ArrNormalText
            End If
        Else                                                               ' если это не последний элемент
            If Trim(Duty.NotEndSimbol(CStr(ArrPasses(ii)))) <> "" Then     ' отрезаем от него послединй символ конца строки убираем пробелы и проверяем на пустоту
                ReDim Preserve ArrNormalText(jj)
                ArrNormalText(jj) = Duty.NotEndSimbol(CStr(ArrPasses(ii)))
                jj = jj + 1
            End If
        End If
    Next ii
        
    
    If foll Then
        ' Перебирая элементы массива собираем их в строку через запятую с указанием порядкового номера
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
    
    ' Проверяем стоит ли галочка суммировать если да, добавляем слово суммировать в текст
    If CheckBoxPassSum.Value Then
        sum = " (суммировать)"
    Else
        sum = ""
    End If

    Selection.TypeText _
    "== Задание " + TextBoxNumAnsPass.Value + " ==" + vbLf _
    + "Сюда вписывается текст задания. Например, сюда #___# необходимо вписать слово яблоко, а сюда #___# - апельсин." + vbLf _
    + "=== Пропуски ===" + vbLf _
    + "=== Оценки" + sum + " ===" + vbLf

    Selection.TypeText TextBoxPassAnswers.Text + vbLf + vbLf
    TextBoxPassAnswers.Text = ""
    Unload FormPass
End Sub

Private Sub CommandButtonPassCancel_Click()
    Unload FormPass
End Sub

Private Sub help_Click()
    HelpForm.refer = "В этом окне необходимо сформировать варианты подстановок в порядке соответствующем пропускам в задании." + vbLf + vbLf _
                   + "В верхнем окне впишите последовательность подстановок, укажите балл за нее и нажмите ""Добавить ответ"". " _
                   + "Последовательность отобразится в нижнем окне ""Варианты ответов""" + vbLf + vbLf _
                   + "Добавьте необходимое количество последовательностей, указывая балл за каждую, подтверждая ввод нажатием кнопки ""Добавить ответ""" + vbLf + vbLf _
                   + "Галочка в чекбоксе ""Суммировать"" сообщает о том, что баллы за выбранные ответы должны суммироваться" + vbLf + vbLf _
                   + "Нажатие на кнопку ""Вывести шаблон"" приведет к выводу последовательностей указанных в окне ""Варианты ответов"" в текст документа Word  и закрытию окна." + vbLf + vbLf
    HelpForm.Show
End Sub


Private Sub TextBoxPasses_Change()
    ' объем всех подстановок не должен превышать 1000 символов
    If Len(TextBoxPasses.Text) > 1000 Then
        MsgBox "Вы превысили лимит, отведенный на подстановки!"
        TextBoxPasses.Text = substit
    End If
    substit = TextBoxPasses.Text
End Sub

Private Sub UserForm_Initialize()
    foll = True
    follNum = 0
End Sub
