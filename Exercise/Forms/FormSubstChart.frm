VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSubstChart 
   Caption         =   "ѕоследовательность подстановок"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   9036.001
   OleObjectBlob   =   "FormSubstChart.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSubstChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If ListBoxAns.ListIndex > -1 Then
        ListBoxRightSec.AddItem ListBoxAns.Value
    End If
End Sub

' удаление варианта из правой колонки
Private Sub CommandButton2_Click()
    If ListBoxRightSec.ListIndex > -1 Then
        Dim remove As Integer
        remove = ListBoxRightSec.ListIndex
        ListBoxRightSec.ListIndex = -1
        ListBoxRightSec.RemoveItem remove
    End If
End Sub

Private Sub CommandButtonExit_Click()
    Unload FormSubstChart
End Sub

Private Sub CommandButtonVar_Click()
    For i = 0 To ListBoxRightSec.ListCount - 1 Step 1
        FormSelect.ListBoxLeft.AddItem ListBoxRightSec.List(i)
    Next i
    Unload FormSubstChart
    FormSelect.Show
End Sub

Private Sub help_Click()
    HelpForm.refer = "Ќа форме необходимо установить правильный пор€док верных ответов в соответствии с пропусками в задании." + vbLf + vbLf _
                   + "ўелчок мыши на варианте ответа в левом окне перемещает его в правое." + vbLf + vbLf _
                   + "—формируйте правильную последовательность ответов в правом окне." + vbLf + vbLf _
                   + "ўелчок мышью на варианте в правом окне удал€ет запись и возвращает ее обратно в левое окно." + vbLf + vbLf _
                   + "Ќажатие кнопки ""«адать варианты"" открывает форму ""‘ормирование вариантов ответов"""
    HelpForm.Show
End Sub

'Private Sub ListBoxAns_Click()
'    Dim remove As Integer
'    remove = ListBoxAns.ListIndex
'    ListBoxRightSec.AddItem ListBoxAns.Value
'    ListBoxAns.ListIndex = -1
'    ListBoxAns.RemoveItem remove     'закомментированные удалени€ в данной форме не требуютс€
'End Sub
'
'Private Sub ListBoxRightSec_Click()
'    Dim remove As Integer
'    remove = ListBoxRightSec.ListIndex
'    ListBoxAns.AddItem ListBoxRightSec.Value  ' закоментированный перенос в другое окно не требуетс€
'    ListBoxRightSec.ListIndex = -1
'    ListBoxRightSec.RemoveItem remove
'End Sub

Private Sub UserForm_Click()

End Sub
