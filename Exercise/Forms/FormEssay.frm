VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormEssay 
   Caption         =   "Эссе"
   ClientHeight    =   3192
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3768
   OleObjectBlob   =   "FormEssay.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormEssay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonEssayExit_Click()
    Unload FormEssay
End Sub

Private Sub CommandButtonOkOneAns_Click()
    ' проверка на пустое значение полей
    If TextBoxNumEss.Value = "" Or TextBoxLongAns.Value = "" Or TextBoxScoreEss.Value = "" Then
        MsgBox "В данной форме поля не могут быть пустыми!"
        Exit Sub
    End If
    
    ' проверка на пустое значение полей
    If TextBoxLongAns.Value < 1 Or TextBoxLongAns.Value > 1024 Then
        MsgBox "Значение поля ""Длина ответа"" в недопустимом диапазоне!"
        Exit Sub
    End If
    
    ' Проверка на корректный ввод баллов
    If Num(TextBoxScoreEss.Value) = "error" Then
        MsgBox "Количество баллов не представленно числом!"
        Exit Sub
    End If
    
    Selection.TypeText _
    "== Задание " + TextBoxNumEss.Value + " ==" + vbLf _
    + "Сюда вписывается текст задания." + vbLf _
    + "=== Свойства ===" + vbLf _
    + "### Длина ответа: " + TextBoxLongAns.Value + vbLf _
    + "=== Оценки ===" + vbLf _
    + "### (" + TextBoxScoreEss.Value + " " + Num(TextBoxScoreEss.Value) + ") "
    
    Selection.TypeText vbLf + vbLf

    
    Unload FormEssay
End Sub

Private Sub UserForm_Click()

End Sub
