Attribute VB_Name = "Search"
Public start As Long          'переменная для хранения номера начального символа диапазона
Public endDoc As Long         'Конец документа
Public TotalScore As Integer  'Общее количество баллов за все задания
Public maxScore As Integer    'Максимальный балл за задание
Public score As Integer       'Текущий рассматриваемый балл
Public sum As Boolean         'Суммировать ли баллы в задании
Public continue As Boolean    'Продолжение цикла
Public startInn As Long       'Начало подпоиска


Sub Init()
    score = 0
    maxScore = 0
    TotalScore = 0
    
    start = ActiveDocument.Range.start ' Номер символа с которого начинается весь документ
    endDoc = ActiveDocument.Range.End
End Sub


Sub Find()
Init
Dim r
Do While start >= 0
    On Error GoTo myError
    Set r = ActiveDocument.Range(start)  'Задаем позицию с которой начинаем поиск
    With r.Find
        .ClearFormatting
        .Text = "Оценки*Задание"         ' Любые символы между указанных слов
        .Forward = True
        .Wrap = wdFindStop               ' Завершаем поиск достигая конца диапазона поиска
        .Format = False
        .MatchCase = True                ' Соблюдаем регистр букв
        .MatchWholeWord = False          ' Ищем весь текст содержащий эти слова, а не только сами слова
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        
        If .Execute Then
            score = 0                    'Обнуляем текущий балл
            maxScore = 0                 'Обнуляем максимальную оценку за задание
            
            'Debug.Print r
            start = r.End                'Конец найденной строки определяем как начало нового поиска
            
            '''' Тут будет код
            Dim inner As Range           'Переменная для подпоиска внутри найденного диапазона
            startInn = r.start           'Задаем начало подпоиска с начала найденного диапазона
            
            'Debug.Print r
            '''Задаем диапазон поиска концом служит конец ранее найденного диапазона
            Set inner = ActiveDocument.Range(startInn, r.End)
            'Debug.Print inner
            With inner.Find
                .ClearFormatting
                .Text = "суммировать"            ' Ищем слово "Суммировать"
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False               ' Не соблюдаем регистр букв
                .MatchWholeWord = False          ' Ищем весь текст содержащий эти слова, а не только сами слова
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = True
                    
                If .Execute Then
                    sum = True
                Else
                    sum = False
                End If
            End With
            
            '''Переопределяем диапазон, переодически изменяя начало поиска, концом служит конец ранее найденного диапазона
            continue = True
            Do While continue
                Set inner = ActiveDocument.Range(startInn, r.End)
                With inner.Find
                    .ClearFormatting
                    .Text = "###*б"                  ' Любые символы между указанных слов
                    .Forward = True
                    .Wrap = wdFindStop
                    .Format = False
                    .MatchCase = True                ' Соблюдаем регистр букв
                    .MatchWholeWord = False          ' Ищем весь текст содержащий эти слова, а не только сами слова
                    .MatchAllWordForms = False
                    .MatchSoundsLike = False
                    .MatchWildcards = True
                    
                    If .Execute Then
                        
                        score = CInt(Left(Right(inner, Len(inner) - 5), Len(inner) - 5 - 2)) ' Отсекаем два символа справа и пять символов слева, полученный символ преобразуем в число
                        startInn = inner.End   'новым началом будет служить конец строки ### (n б
                        If sum Then
                            maxScore = maxScore + score 'если ранее мы нашли слово суммировать складываем баллы
                        Else
                            If score > maxScore Then
                                maxScore = score     'если нет просто выбираем наибольший
                            End If
                        End If
                    Else
                        continue = False    'если более ничего не находим прерываем цикл поиска
                    End If
                End With
            Loop
            
            TotalScore = TotalScore + maxScore
            '''' Тут будет его окончание
        Else
            ' заключительный подсчет
            score = 0                    'Обнуляем текущий балл
            maxScore = 0                 'Обнуляем максимальную оценку за задание
            
            Set inner = ActiveDocument.Range(startInn - 6, endDoc) '+++
            'Ищем начальную точку заключительного поиска
            With inner.Find
                .ClearFormatting
                .Text = "Оценки"            ' Ищем слово "Оценки"
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False               ' Не соблюдаем регистр букв
                .MatchWholeWord = False          ' Ищем весь текст содержащий эти слова, а не только сами слова
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = True
                    
                If .Execute Then
                    startInn = inner.start
                End If
                'Debug.Print inner
            End With
            
            'переопределяем начало поиска
            Set inner = ActiveDocument.Range(startInn, endDoc)
            'Debug.Print inner
            With inner.Find
                .ClearFormatting
                .Text = "суммировать"            ' Ищем слово "Суммировать"
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False               ' Не соблюдаем регистр букв
                .MatchWholeWord = False          ' Ищем весь текст содержащий эти слова, а не только сами слова
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = True
                    
                If .Execute Then
                    sum = True
                Else
                    sum = False
                End If
            End With
            Do While startInn > 0
                Set inner = ActiveDocument.Range(startInn, endDoc)
                With inner.Find
                    .ClearFormatting
                    .Text = "###*б"                  ' Любые символы между указанных слов
                    .Forward = True
                    .Wrap = wdFindStop
                    .Format = False
                    .MatchCase = True                ' Соблюдаем регистр букв
                    .MatchWholeWord = False          ' Ищем весь текст содержащий эти слова, а не только сами слова
                    .MatchAllWordForms = False
                    .MatchSoundsLike = False
                    .MatchWildcards = True
                    
                    If .Execute Then
                        
                        score = CInt(Left(Right(inner, Len(inner) - 5), Len(inner) - 5 - 2)) ' Отсекаем два символа справа и пять символов слева, полученный символ преобразуем в число
                        startInn = inner.End
                        If sum Then
                            maxScore = maxScore + score
                        Else
                            If score > maxScore Then
                                maxScore = score
                            End If
                        End If
                    Else
                        startInn = 0
                    End If
                End With
            Loop
            
            TotalScore = TotalScore + maxScore
            start = -1
            
            MsgBox "Общий балл посчитан и составляет: " + CStr(TotalScore) + "!", vbExclamation
        End If
    End With
'Debug.Print r
Loop
myError:
    If Err Then
        MsgBox "Ошибка! Вероятно в документе отсутствуют баллы."
    End If
End Sub

