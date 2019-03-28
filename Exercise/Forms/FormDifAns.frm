VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormDifAns 
   Caption         =   "Вопрос с несколькими вариантами ответа"
   ClientHeight    =   6552
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   6588
   OleObjectBlob   =   "FormDifAns.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormDifAns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonClose_Click()
    Unload FormDifAns
End Sub
' Добавляет записанные ответы в нижнее окно правильных ответов
Private Sub CommandButtonDifAdd_Click()
    Dim AllAns() As String
    
    Dim j As Integer
    j = 0
    AllAns = Split(TextBoxAllAns.Text, vbLf)
    
    
    
    '  ++++ Проверка на одинаковые варианты ответов ++++
    If Not Duty.NoIdentical(AllAns) Then
        MsgBox "Не может быть двух одинаковых ответов!"
        TextBoxAllAns.SetFocus
        Exit Sub
    End If
    For Each h In AllAns
        ' Проверка на пустой вариант ответа
'        If h = "" Then
'            MsgBox "В добавляемых вариантах присутствует пустой вариант ответа, удалите его, или допишите ответ!"
'            TextBoxAllAns.SetFocus
'            Exit Sub
'        End If
        For t = 0 To ListBoxRight.ListCount - 1
            If h = ListBoxRight.List(t) Then
                MsgBox "Не может быть двух одинаковых ответов!"
                TextBoxAllAns.SetFocus
                Exit Sub
            End If
        Next t
    Next h
    ' ++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    ' Заполняем ListBox для выбора правильных ответов
    For Each Item In AllAns
        If j < UBound(AllAns) Then
            If Left(Item, Len(Item) - 1) <> "" Then  ' пропускаем пустые строки в массиве
            If Left(Item, Len(Item) - 1) <> " " Then
                ListBoxRight.AddItem Left(Item, Len(Item) - 1)
            End If
            End If
        Else
            If Item <> "" Then
            If Item <> " " Then
                ListBoxRight.AddItem Item
            End If
            End If
        End If
        j = j + 1
    Next Item
    
    ' Очищаем поле ввода
    TextBoxAllAns.Value = ""
End Sub
' удаление элементов из списка
Private Sub CommandButtonDivRemove_Click()
    Dim f As Integer
    f = 0
    Do While (f < ListBoxRight.ListCount)
        If ListBoxRight.Selected(f) Then
            ListBoxRight.RemoveItem (f)
            f = -1
        End If
        f = f + 1
    Loop
End Sub

Private Sub CommandButtonOK_Click()
    Dim ans As String   'строка для ответов
    ans = ""
    
    ' проверка на пустое значение полей
    If TextBoxNumDifAns.Value = "" Or ListBoxRight.ListCount < 1 Then
        MsgBox "В данной форме поля не могут быть пустыми!"
        Exit Sub
    End If
    
    ' Проверка на количество вариантов ответов, их должно быть от 2 до 7
    If ListBoxRight.ListCount < 2 Then
        MsgBox "Вариантов ответов должно быть больше одного!"
        TextBoxAllAns.SetFocus
        Exit Sub
    ElseIf ListBoxRight.ListCount > 7 Then
        MsgBox "Количество вариантов ответов не должно превышать семи!"
        TextBoxAllAns.SetFocus
        Exit Sub
    End If
    
    ' Проверка на выбор хотя бы одного правильного ответа
    For y = 0 To ListBoxRight.ListCount - 1 Step 1
        If ListBoxRight.Selected(y) Then
            y = -1
            Exit For
        End If
    Next y
    If y <> -1 Then
        MsgBox "Невыбрано ни одного правильного ответа!"
        Exit Sub
    End If
    

    ' Если ошибок не было
    Dim j As Integer
    j = 0
    For k = 0 To ListBoxRight.ListCount - 1 Step 1
        If ListBoxRight.Selected(k) Then
            'создаем на форме выбора вариантов список правильных ответов
            FormDifVarAns.ListBoxAllRightAns.AddItem ListBoxRight.List(k, 0)
            'помечаем ответ как правильный
            ans = ans + "## " + ListBoxRight.List(k, 0) + vbLf
        Else
            'помечаем ответ как не правильны
            ans = ans + "# " + ListBoxRight.List(k, 0) + vbLf
        End If
    Next k
    
    'передаем на форму выбора варинатов ответа правильные ответы и номер задачи
    FormDifVarAns.ans = ans
    FormDifVarAns.numAns = TextBoxNumDifAns.Value
    
    'скрываем текущую форму и активируем форму выбора вариантов
    FormDifAns.Hide
    FormDifVarAns.Show
    
End Sub


Private Sub help_Click()
    HelpForm.refer = "Введите порядковый номер вопроса." + vbLf + vbLf _
                   + "Затем заполните текстовое поле возможными ответами, вводя каждый ответ в новой строке. " _
                   + "(Переход на новую строку осуществляется нажатием сочетания клавиш Shift+Enter.)" + vbLf + vbLf _
                   + "При нажатии на кнопку ""Добавить варианты"" ответы из верхнего окна добавятся к списку в нижнем." _
                   + "Кнопка ""Удалить выделенное"" удаляет ответы из нижнего поля отмеченные галочками." + vbLf + vbLf _
                   + "Далее в нижнем окне нужно выбрать правильные варианты ответов, щелкнув на них мышью." + vbLf + vbLf _
                   + "Нажатие кнопки ОК откроет новое окно со списком только что выбранных правильных ответов."
    HelpForm.Show
End Sub

Private Sub Label2_Click()

End Sub

Private Sub UserForm_Click()

End Sub
