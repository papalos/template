Attribute VB_Name = "Search"
Public start As Long          '���������� ��� �������� ������ ���������� ������� ���������
Public endDoc As Long         '����� ���������
Public TotalScore As Integer  '����� ���������� ������ �� ��� �������
Public maxScore As Integer    '������������ ���� �� �������
Public score As Integer       '������� ��������������� ����
Public sum As Boolean         '����������� �� ����� � �������
Public continue As Boolean    '����������� �����
Public startInn As Long       '������ ���������


Sub Init()
    score = 0
    maxScore = 0
    TotalScore = 0
    
    start = ActiveDocument.Range.start ' ����� ������� � �������� ���������� ���� ��������
    endDoc = ActiveDocument.Range.End
End Sub


Sub Find()
Init
Dim r
Do While start >= 0
    On Error GoTo myError
    Set r = ActiveDocument.Range(start)  '������ ������� � ������� �������� �����
    With r.Find
        .ClearFormatting
        .Text = "������*�������"         ' ����� ������� ����� ��������� ����
        .Forward = True
        .Wrap = wdFindStop               ' ��������� ����� �������� ����� ��������� ������
        .Format = False
        .MatchCase = True                ' ��������� ������� ����
        .MatchWholeWord = False          ' ���� ���� ����� ���������� ��� �����, � �� ������ ���� �����
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        
        If .Execute Then
            score = 0                    '�������� ������� ����
            maxScore = 0                 '�������� ������������ ������ �� �������
            
            'Debug.Print r
            start = r.End                '����� ��������� ������ ���������� ��� ������ ������ ������
            
            '''' ��� ����� ���
            Dim inner As Range           '���������� ��� ��������� ������ ���������� ���������
            startInn = r.start           '������ ������ ��������� � ������ ���������� ���������
            
            'Debug.Print r
            '''������ �������� ������ ������ ������ ����� ����� ���������� ���������
            Set inner = ActiveDocument.Range(startInn, r.End)
            'Debug.Print inner
            With inner.Find
                .ClearFormatting
                .Text = "�����������"            ' ���� ����� "�����������"
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False               ' �� ��������� ������� ����
                .MatchWholeWord = False          ' ���� ���� ����� ���������� ��� �����, � �� ������ ���� �����
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = True
                    
                If .Execute Then
                    sum = True
                Else
                    sum = False
                End If
            End With
            
            '''�������������� ��������, ������������ ������� ������ ������, ������ ������ ����� ����� ���������� ���������
            continue = True
            Do While continue
                Set inner = ActiveDocument.Range(startInn, r.End)
                With inner.Find
                    .ClearFormatting
                    .Text = "###*�"                  ' ����� ������� ����� ��������� ����
                    .Forward = True
                    .Wrap = wdFindStop
                    .Format = False
                    .MatchCase = True                ' ��������� ������� ����
                    .MatchWholeWord = False          ' ���� ���� ����� ���������� ��� �����, � �� ������ ���� �����
                    .MatchAllWordForms = False
                    .MatchSoundsLike = False
                    .MatchWildcards = True
                    
                    If .Execute Then
                        
                        score = CInt(Left(Right(inner, Len(inner) - 5), Len(inner) - 5 - 2)) ' �������� ��� ������� ������ � ���� �������� �����, ���������� ������ ����������� � �����
                        startInn = inner.End   '����� ������� ����� ������� ����� ������ ### (n �
                        If sum Then
                            maxScore = maxScore + score '���� ����� �� ����� ����� ����������� ���������� �����
                        Else
                            If score > maxScore Then
                                maxScore = score     '���� ��� ������ �������� ����������
                            End If
                        End If
                    Else
                        continue = False    '���� ����� ������ �� ������� ��������� ���� ������
                    End If
                End With
            Loop
            
            TotalScore = TotalScore + maxScore
            '''' ��� ����� ��� ���������
        Else
            ' �������������� �������
            score = 0                    '�������� ������� ����
            maxScore = 0                 '�������� ������������ ������ �� �������
            
            Set inner = ActiveDocument.Range(startInn - 6, endDoc) '+++
            '���� ��������� ����� ��������������� ������
            With inner.Find
                .ClearFormatting
                .Text = "������"            ' ���� ����� "������"
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False               ' �� ��������� ������� ����
                .MatchWholeWord = False          ' ���� ���� ����� ���������� ��� �����, � �� ������ ���� �����
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = True
                    
                If .Execute Then
                    startInn = inner.start
                End If
                'Debug.Print inner
            End With
            
            '�������������� ������ ������
            Set inner = ActiveDocument.Range(startInn, endDoc)
            'Debug.Print inner
            With inner.Find
                .ClearFormatting
                .Text = "�����������"            ' ���� ����� "�����������"
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False               ' �� ��������� ������� ����
                .MatchWholeWord = False          ' ���� ���� ����� ���������� ��� �����, � �� ������ ���� �����
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
                    .Text = "###*�"                  ' ����� ������� ����� ��������� ����
                    .Forward = True
                    .Wrap = wdFindStop
                    .Format = False
                    .MatchCase = True                ' ��������� ������� ����
                    .MatchWholeWord = False          ' ���� ���� ����� ���������� ��� �����, � �� ������ ���� �����
                    .MatchAllWordForms = False
                    .MatchSoundsLike = False
                    .MatchWildcards = True
                    
                    If .Execute Then
                        
                        score = CInt(Left(Right(inner, Len(inner) - 5), Len(inner) - 5 - 2)) ' �������� ��� ������� ������ � ���� �������� �����, ���������� ������ ����������� � �����
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
            
            MsgBox "����� ���� �������� � ����������: " + CStr(TotalScore) + "!", vbExclamation
        End If
    End With
'Debug.Print r
Loop
myError:
    If Err Then
        MsgBox "������! �������� � ��������� ����������� �����."
    End If
End Sub

