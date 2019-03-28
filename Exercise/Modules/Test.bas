Attribute VB_Name = "Test"
'�������������� �������� ������ � �������� ������ � ���������

Function Search(ByVal lookFor As String, ByVal startDoc As Long, ByVal endDoc As Long, ByRef findedStr As Range) As Boolean
    Set r = ActiveDocument.Range(startDoc, endDoc)  '������ ������� � ������� �������� �����
    With r.Find
        .ClearFormatting
        .Text = lookFor                  ' ����� ������� ����� ��������� ���� "�����1*�����2"
        .Forward = True
        .Wrap = wdFindStop               ' ��������� ����� �������� ����� ��������� ������
        .Format = False
        .MatchCase = True                ' ��������� ������� ����
        .MatchWholeWord = False          ' ���� ���� ����� ���������� ��� �����, � �� ������ ���� �����
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
        fined = Search("������*�������", TotalDocStart, TotalDocEnd, RangeAllDoc)
        TotalDocStart = RangeAllDoc.End
        cicle = fined
        'Debug.Print cicle                                 'TEST
        'Debug.Print RangeAllDoc                           'TEST
        
        '���������� ���������� ��� ������ � ����� ���������� ������� ������
        Dim FinedPartDocStart As Long
        Dim FinedPartDocEnd As Long
                        
        ' ���������� ��� �������� �������� ������ � �������
        Dim RangePart As Range
            
        '����������� ����� ��� ���
        Dim YesNo As Boolean
        
        If fined Then
            '���� ������, ���� � ��������
            
            '�������������� ��� ���������� ������� � ������ ���������� ������� ���������
            FinedPartDocStart = RangeAllDoc.start
            FinedPartDocEnd = RangeAllDoc.End
            
            YesNo = isSumming(RangeAllDoc, "�����������")
            score = 0
            
            Do While cicle
                fined = Search("###*�", FinedPartDocStart, FinedPartDocEnd, RangePart)
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
            '���� �� ������ �������� �������������� ����� �� ����� ���������
            
            '�������������� ��� ���������� ������� � ������ ���������� ������� ���������
            FinedPartDocStart = RangeAllDoc.End + 6
            
            
            Call Search("������", FinedPartDocStart, TotalDocEnd, RangePart)
            Debug.Print RangePart
            FinedPartDocStart = RangePart.End
            FinedPartDocEnd = TotalDocEnd
            
            RangeAllDoc.start = FinedPartDocStart
                      
            '����������� ����� ��� ���
            YesNo = isSumming(RangeAllDoc, "�����������")
            'Debug.Print RangePart            'TEST
            score = 0
            cicle = True
            Do While cicle
                fined = Search("###*�", FinedPartDocStart, FinedPartDocEnd, RangePart)
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
            MsgBox "����� ���� ��������: " + CStr(TotalScore)
        End If
    Loop
    
End Sub

'����������� ��������� ������� � �����
Function sumScore(ByVal check As Boolean, ByRef part As Range) As Integer
    If check Then
        sumScore = CInt(Left(Right(part, Len(part) - 5), Len(part) - 5 - 2)) ' �������� ��� ������� ������ � ���� �������� �����, ���������� ������ ����������� � �����

    Else
        sumScore = 0
    End If
End Function

'���������� ������� �� �������� ������ � ��������� ���������.
Function isSumming(ByVal isSum As Range, lookFor As String) As Boolean
    With isSum.Find
        .ClearFormatting
        .Text = lookFor                 ' ���� ����� "�����������"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False               ' �� ��������� ������� ����
        .MatchWholeWord = False          ' ���� ���� ����� ���������� ��� �����, � �� ������ ���� �����
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
                    
        isSumming = .Execute
    End With
End Function

