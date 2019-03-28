Attribute VB_Name = "Duty"

' функция отрезает последний символ у строки (предполагается удаление символа конца строки)
Public Function NotEndSimbol(str As String)
    NotEndSimbol = Left(str, Len(str) - 1)
End Function

' Не допускаются одинаковые варианты ответов
Public Function NoIdentical(Arr() As String) As Boolean
Dim rt As Boolean

    For k = 0 To UBound(Arr) Step 1
        rt = False
        If k = UBound(Arr) Then
            If Arr(k) <> "" Then
                rt = True
            End If
        Else
            If Duty.NotEndSimbol(Arr(k)) <> "" Then
                rt = True
            End If
        End If
        
        If rt Then
        For m = k + 1 To UBound(Arr) Step 1
            If m = UBound(Arr) Then
                If Duty.NotEndSimbol(Arr(k)) = Arr(m) Then
                    NoIdentical = False
                    Exit Function
                End If
            Else
                    
                If Arr(k) = Arr(m) Then
                    NoIdentical = False
                    Exit Function
                End If
            End If
        Next m
        End If
    
    Next k
    NoIdentical = True
End Function


        
