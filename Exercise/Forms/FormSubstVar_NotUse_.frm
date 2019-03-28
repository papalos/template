VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSubstVar_NotUse_ 
   Caption         =   "Добавление правильного ответа"
   ClientHeight    =   4656
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8280.001
   OleObjectBlob   =   "FormSubstVar_NotUse_.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSubstVar_NotUse_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonIns_Click()
    Selection.TypeText TextBoxAnswers.Text
    TextBoxAnswers.Text = ""
End Sub

Private Sub CommandButtonSubstCancel_Click()
    Unload FormSubstVar
End Sub

Private Sub CommandButtonSubstOK_Click()

    Dim ArrText() As String
    Dim ins As String
    Dim counting As Integer
    
    counting = 0
    ArrText = Split(TextBoxSubst.Text, vbLf)
    
    For Each s In ArrText
        counting = counting + 1
        If counting = 1 Then
            If counting = UBound(ArrText) + 1 Then
                ins = ins + CStr(counting) + "- " + s
            Else
                ins = CStr(counting) + "- " + Left(s, Len(s) - 1)
            End If
        ElseIf counting = UBound(ArrText) + 1 Then
            ins = ins + ";" + CStr(counting) + "- " + s
        Else
            ins = ins + ";" + CStr(counting) + "- " + Left(s, Len(s) - 1)
        End If
    Next s

    
    ' проверка на пустое значение полей
    If TextBoxSubst.Value = "" Or TextBoxSubstScore.Value = "" Then
        MsgBox "В данной форме поля не могут быть пустыми!"
        Exit Sub
    End If
    
    ' Проверка на корректный ввод баллов
    If Num(TextBoxSubstScore.Value) = "error" Then
        MsgBox "Количество баллов не представленно числом!"
        Exit Sub
    End If
    
    TextBoxAnswers.Text = TextBoxAnswers.Text + "### (" + TextBoxSubstScore.Value + " " + Num(TextBoxSubstScore.Value) + ") " + ins + vbLf
    TextBoxSubstScore.Value = ""
    TextBoxSubst.Value = ""
End Sub

Private Sub help_Click()
    HelpForm.refer = "Так как в предыдущем окне не было задано вариантов подстановки, в этом окне необходимо их сформировать." + vbLf + vbLf _
                   + "В верхнем окне впишите последовательность подстановок, укажите балл за нее и нажмите ""Добавить ответ""." + vbLf + vbLf _
                   + "Последовательность отобразится в нижнем окне ""Варианты ответов""" + vbLf + vbLf _
                   + "Добавьте необходимое количество последовательностей, указывая балл за каждую, подтверждая ввод нажатием кнопки ""Добавить ответ""" + vbLf + vbLf _
                   + "Нажатие на кнопку ""Вывести шаблон"" приведет к выводу последовательностей указанных в окне ""Варианты ответов"" в текст документа Word  и закрытию окна." + vbLf + vbLf _
                   + "Кнопкой ""Ввод правильных ответов"" на вкладке ""Разработка заданий"" можно воспользоваться отдельно для открытия формы ""Ввод правильных ответов"", чтобы дописать недостающие варианты ответов." _
                   + "Кнопка ""Добавить ответ"" записывает вариант ответа в документ в место положения курсора (следите за тем, чтобы он находился в нужном месте после заголовка ===Оценки===."
    HelpForm.Show
End Sub

Private Sub UserForm_Click()

End Sub
