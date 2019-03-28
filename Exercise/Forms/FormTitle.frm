VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormTitle 
   Caption         =   "Заголовок"
   ClientHeight    =   2880
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   6456
   OleObjectBlob   =   "FormTitle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButtonCancel_Click()
TextBoxProf.Value = ""
TextBoxGrade.Value = ""
TextBoxVar.Value = ""
FormTitle.Hide
End Sub

Private Sub CommandButtonOK_Click()
' проверка на пустое значение полей
    If TextBoxProf.Value = "" Or TextBoxGrade.Value = "" Or TextBoxVar.Value = "" Then
        MsgBox "В данной форме поля не могут быть пустыми!"
        Exit Sub
    End If
    
    ' Проверка на то, что введены числа
    If Num(TextBoxVar.Value) = "error" Then
        MsgBox "В полях предполагающие числа, некоретный ввод!"
        Exit Sub
    End If
    
    ' Проверка на то, что введены числа (не используется в этой версии)
'    If CInt(TextBoxGrade.Value) < 7 Or CInt(TextBoxGrade.Value) > 12 Then
'        MsgBox "Класс может быть от 7 до 11-ого!"
'        Exit Sub
'    End If
    
    ' Проверка на то, что введены числа
    If CInt(TextBoxVar.Value) < 1 Or CInt(TextBoxVar.Value) > 10 Then
        MsgBox "Не более 10 вариантов!"
        Exit Sub
    End If
    
    
Selection.TypeText "= " + TextBoxProf.Value + ", " + TextBoxGrade.Value + " класс" + ", " + "Вариант " + TextBoxVar.Value + " =" + vbLf
TextBoxProf.Value = ""
TextBoxGrade.Value = ""
TextBoxVar.Value = ""
FormTitle.Hide
End Sub

Private Sub help_Click()
    'Вызов справки
    HelpForm.refer = "Введите название профиля по которому составляете задание, " _
               + "числом укажите класс для которого составляется задание и проставьте числом номер варианта." + vbLf + vbLf _
               + "Если предполагается один вариант задания, все равно внесите в поле варианта единицу (цифрами)." + vbLf + vbLf _
               + "После чего нажмите кнопку ""ОК"", в документе отобразится введенная информация в требуемом формате." + vbLf + vbLf _
               + "Нажатие по кнопке ""Отмена"" закроет форму без сохранения информации."

    HelpForm.Show
End Sub

