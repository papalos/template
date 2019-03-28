Attribute VB_Name = "Test"
'Альтернативный алгоритм поиска и подсчета баллов в документе

Function Search(ByVal lookFor As String, ByVal startDoc As Long, ByVal endDoc As Long, ByRef findedStr As Range) As Boolean
    Set r = ActiveDocument.Range(startDoc, endDoc)  'Задаем позицию с которой начинаем поиск
    With r.Find
        .ClearFormatting
        .Text = lookFor                  ' Любые символы между указанных слов "Слово1*Слово2"
        .Forward = True
        .Wrap = wdFindStop               ' Завершаем поиск достигая конца диапазона поиска
        .Format = False
        .MatchCase = True                ' Соблюдаем регистр букв
        .MatchWholeWord = False          ' Ищем весь текст содержащий эти слова, а не только сами слова
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        
        If .Execute Then
            Search = True
            Set findedStr = r
            Exit Function
        Else
            Search = False
        End If
    End With
End Function

Sub lookFor()
    Dim TotalScore As Integer
    Dim TotalDocStart As Long
    Dim TotalDocEnd As Long
    Dim RangeAllDoc As Range
    Dim score As Integer
    
    With ActiveDocument.Range
        TotalDocStart = .start
        TotalDocEnd = .End
    End With
    
    Dim cicle As Boolean, fined As Boolean
    cicle = True: fined = False
    
    Do While cicle
        fined = Search("Оценки*Задание", TotalDocStart, TotalDocEnd, RangeAllDoc)
        TotalDocStart = RangeAllDoc.End
        cicle = fined
        'Debug.Print cicle                                 'TEST
        'Debug.Print RangeAllDoc                           'TEST
        
        'определяем переменные для начала и конца найденного участка текста
        Dim FinedPartDocStart As Long
        Dim FinedPartDocEnd As Long
                        
        ' Переменная для хранения найденой строки с баллами
        Dim RangePart As Range
            
        'Суммировать баллы или нет
        Dim YesNo As Boolean
        
        If fined Then
            'если найден, ищем в найденом
            
            'Инициализируем эти переменные началом и концом найденного участка документа
            FinedPartDocStart = RangeAllDoc.start
            FinedPartDocEnd = RangeAllDoc.End
            
            YesNo = isSumming(RangeAllDoc, "суммировать")
            score = 0
            
            Do While cicle
                fined = Search("###*б", FinedPartDocStart, FinedPartDocEnd, RangePart)
                If fined Then
                    If YesNo Then
                        score = score + sumScore(fined, RangePart)
                    Else
                        If score < sumScore(fined, RangePart) Then
                            score = sumScore(fined, RangePart)
                        End If
                    End If
                    FinedPartDocStart = RangePart.End
                    'Debug.Print RangePart
                End If
                cicle = fined
            Loop
            TotalScore = TotalScore + score
            cicle = True
        Else
            'если не найден проводим заключительный поиск до конца документа
            
            'Инициализируем эти переменные началом и концом найденного участка документа
            FinedPartDocStart = RangeAllDoc.End + 6
            
            
            Call Search("Оценки", FinedPartDocStart, TotalDocEnd, RangePart)
            Debug.Print RangePart
            FinedPartDocStart = RangePart.End
            FinedPartDocEnd = TotalDocEnd
            
            RangeAllDoc.start = FinedPartDocStart
                      
            'Суммировать баллы или нет
            YesNo = isSumming(RangeAllDoc, "суммировать")
            'Debug.Print RangePart            'TEST
            score = 0
            cicle = True
            Do While cicle
                fined = Search("###*б", FinedPartDocStart, FinedPartDocEnd, RangePart)
                If fined Then
                    If YesNo Then
                        score = score + sumScore(fined, RangePart)
                    Else
                        If score < sumScore(fined, RangePart) Then
                            score = sumScore(fined, RangePart)
                        End If
                    End If
                    
                    FinedPartDocStart = RangePart.End
                    'Debug.Print RangePart                 'TEST
                End If
                cicle = fined
            Loop
            TotalScore = TotalScore + score
            MsgBox "Общий балл посчитан: " + CStr(TotalScore)
        End If
    Loop
    
End Sub

'Преобразует найденную позицию в число
Function sumScore(ByVal check As Boolean, ByRef part As Range) As Integer
    If check Then
        sumScore = CInt(Left(Right(part, Len(part) - 5), Len(part) - 5 - 2)) ' Отсекаем два символа справа и пять символов слева, полученный символ преобразуем в число

    Else
        sumScore = 0
    End If
End Function

'Возвращает найдена ли заданная строка в указанном диапазоне.
Function isSumming(ByVal isSum As Range, lookFor As String) As Boolean
    With isSum.Find
        .ClearFormatting
        .Text = lookFor                 ' Ищем слово "Суммировать"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False               ' Не соблюдаем регистр букв
        .MatchWholeWord = False          ' Ищем весь текст содержащий эти слова, а не только сами слова
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
                    
        isSumming = .Execute
    End With
End Function

