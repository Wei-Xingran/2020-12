VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub Json����() '����UTF8�ı��ļ�
Attribute Json����.VB_ProcData.VB_Invoke_Func = "e\n14"
 'myrange = ActiveWorkbook.Names("schoolinfo").RefersToRange 'ͨ�������������ѡ������
 Dim Findname, Outname As String
 Dim wb As Workbook
 
myRange = Selection  'ͨ��ѡ��������ѡ������
cou = UBound(myRange)
For ric = 1 To cou


Findname = myRange(ric, 1)    'Ŀ�����
Outname = myRange(ric, 2)     '������
Set wb = GetObject(ThisWorkbook.Path & "\" & Findname & ".xlsx")    '��ȡĿ���������
Set st1 = wb.Worksheets("Value")
    
'used1 = Range(st1.Range("a1").End(xlDown), st1.Range("a1").End(xlToRight)) 'ͨ�������е�������������ѡ������
Total1 = st1.Range("A65535").End(xlUp).Row 'UBound(used1, 1) '��ȡ����
Fields1 = st1.Range("IV1").End(xlToLeft).Column 'UBound(used1, 2) '��ȡ����

ReDim myrange1(1 To Total1, 1 To Fields1)  '���ݸ�ֵΪtext����
For i = 1 To Total1
    For j = 1 To Fields1
        If IsNumeric(st1.Cells(i, j)) Then
            myrange1(i, j) = Replace(st1.Cells(i, j).Text, " ", "")
        Else
            myrange1(i, j) = Application.Substitute(st1.Cells(i, j).Text, Chr(10), "\n")
        End If
    Next
Next



If wb.Worksheets.Count = 2 Then            '�ж��Ƿ���Array��
Set st2 = wb.Worksheets("Array")
    'used2 = Range(st2.Range("a1").End(xlDown), st2.Range("a1").End(xlToRight)) 'ͨ�������е�������������ѡ������
    Total2 = st2.Range("A65535").End(xlUp).Row 'UBound(used2, 1) '��ȡ����
    Fields2 = st2.Range("IV1").End(xlToLeft).Column 'UBound(used2, 2) '��ȡ����
    ReDim myrange2(1 To Total2, 1 To Fields2)  '���ݸ�ֵΪtext����
    For i = 1 To Total2
        For j = 1 To Fields2
            If IsNumeric(st2.Cells(i, j)) Then
                myrange2(i, j) = Replace(st2.Cells(i, j).Text, " ", "")
            Else
                myrange2(i, j) = st2.Cells(i, j).Text
            End If
        Next
    Next
Else
    Total2 = 0
    Fields2 = 0
 
End If


wb.Close            '�ر�Ŀ���
     
     
     Dim objStream As Object
     Set objStream = CreateObject("ADODB.Stream")
     
     With objStream
            .Type = 2
            .Charset = "UTF-8"
            .Open
            .WriteText "["
    'Value��
            For i = 2 To Total1
                .WriteText vbLf & "  " & "{" & vbLf & "    "
                For j = 1 To Fields1
                    If IsNumeric(myrange1(i, j)) And Left(myrange1(1, j), 8) <> "strCount" Then
                        .WriteText """" & myrange1(1, j) & """" & ": " & Replace(myrange1(i, j), """", "\""")           '����
                    ElseIf myrange1(i, j) = "FALSE" Then
                        .WriteText """" & myrange1(1, j) & """" & ": " & Replace("false", """", "\""")          '����ֵ
                    ElseIf myrange1(i, j) = "TRUE" Then
                        .WriteText """" & myrange1(1, j) & """" & ": " & Replace("true", """", "\""")          '����ֵ
                    Else
                        .WriteText """" & myrange1(1, j) & """: """ & Replace(myrange1(i, j), """", "\""") & """"         '�ı�
                    End If
                     If j <> Fields1 Then
                        .WriteText "," & vbLf & "    "
                     End If
                Next
    'Array��
                If Total2 And Fields2 > 0 Then
                    For n = 2 To Fields2
                        fir = 0
                        .WriteText "," & vbLf & "    " & """" & Replace(myrange2(1, n), """", "\""") & """" & ": " & "["
                        For m = 2 To Total2
                          If myrange2(m, n) <> "" Then
                            If myrange2(m, 1) = myrange1(i, 1) And fir = 0 Then
                          
                            
                                fir = 1
                                If IsNumeric(myrange2(m, n)) Then
                                    .WriteText vbLf & "      " & Replace(myrange2(m, n), """", "\""")
                                ElseIf myrange2(m, n) = "FALSE" Then
                                    .WriteText vbLf & "      " & Replace("false", """", "\""")
                                ElseIf myrange2(m, n) = "TRUE" Then
                                    .WriteText vbLf & "      " & Replace("true", """", "\""")
                                Else
                                    .WriteText vbLf & "      " & """" & Replace(myrange2(m, n), """", "\""") & """"
                                End If
                            ElseIf myrange2(m, 1) = myrange1(i, 1) And fir = 1 Then
                                If IsNumeric(myrange2(m, n)) Then
                                    .WriteText "," & vbLf & "      " & Replace(myrange2(m, n), """", "\""")
                                ElseIf myrange2(m, n) = "FALSE" Then
                                    .WriteText "," & vbLf & "      " & Replace("false", """", "\""")
                                ElseIf myrange2(m, n) = "TRUE" Then
                                    .WriteText "," & vbLf & "      " & Replace("true", """", "\""")
                                Else
                                    .WriteText "," & vbLf & "      " & """" & Replace(myrange2(m, n), """", "\""") & """"
                                End If
                            End If
                            End If
                        Next
                        If n = Fields2 Then
                            .WriteText vbLf & "    " & "]"
                        Else
                            .WriteText vbLf & "    " & "]"
                        End If
                    Next
                End If
                
            
                
                If i = Total1 Then
                        .WriteText vbLf & "  " & "}"
                Else
                        .WriteText vbLf & "  " & "},"
                End If
            Next
 
            .WriteText vbLf & "]"
            filename = ThisWorkbook.Path & "\" & Outname
            tempFilename = filename & ".json"
            .SaveToFile tempFilename, 2
            .Flush
            .Close
            MsgBox Findname & vbCrLf & "�����ɹ�"
     End With
     Set objStream = Nothing
     ' ȥ��BOMͷ
    Call Utf8WithoutBom(tempFilename, filename)
    
    ' ɾ����ʱ�ļ�
    Kill tempFilename
    Name filename As filename & ".json"
    
Next
End Sub


Sub ����Json����() '����UTF8�ı��ļ�
 Dim Findname, Outname As String
 Dim wb As Workbook
 
myRange = Selection  'ͨ��ѡ��������ѡ������
Findname = myRange(1, 1)    'Ŀ�����
Outname = myRange(1, 2)     '������
 
 Set wb = GetObject(ThisWorkbook.Path & "\" & Findname & ".xlsx")    '��ȡĿ���������
 Set st1 = wb.Worksheets("Value")
    
'used1 = Range(st1.Range("a1").End(xlDown), st1.Range("a1").End(xlToRight)) 'ͨ�������е�������������ѡ������
Total1 = st1.Range("A65535").End(xlUp).Row 'UBound(used1, 1) '��ȡ����
Fields1 = st1.Range("IV1").End(xlToLeft).Column 'UBound(used1, 2) '��ȡ����

ReDim myrange1(1 To Total1, 1 To Fields1)  '���ݸ�ֵΪtext����
For i = 1 To Total1
    For j = 1 To Fields1
        myrange1(i, j) = st1.Cells(i, j).Text
    Next
Next

If wb.Worksheets.Count = 2 Then            '�ж��Ƿ���Array��
Set st2 = wb.Worksheets("Array")
    'used2 = Range(st2.Range("a1").End(xlDown), st2.Range("a1").End(xlToRight)) 'ͨ�������е�������������ѡ������
    Total2 = st2.Range("A65535").End(xlUp).Row 'UBound(used2, 1) '��ȡ����
    Fields2 = st2.Range("IV1").End(xlToLeft).Column 'UBound(used2, 2) '��ȡ����
    ReDim myrange2(1 To Total2, 1 To Fields2)  '���ݸ�ֵΪtext����
    For i = 1 To Total2
        For j = 1 To Fields2
            myrange2(i, j) = st2.Cells(i, j).Text
        Next
    Next
Else
    Total2 = 0
    Fields2 = 0
 
End If


wb.Close


     Dim objStream As Object
     Set objStream = CreateObject("ADODB.Stream")
     
     With objStream
            .Type = 2
            .Charset = "UTF-8"
            .Open
            .WriteText "["
    'Value��
             For i = 2 To Total1
                .WriteText vbLf & "  " & "{" & vbLf & "    "
                If IsNumeric(myrange2(i, 1)) Then
                    .WriteText """" & myrange1(1, 1) & """" & ": " & Replace(myrange1(i, 1), """", "\""")           '����
                Else
                    .WriteText """" & myrange1(1, 1) & """: """ & Replace(myrange1(i, 1), """", "\""") & """"         '�ı�
                End If
                .WriteText "," & vbLf & "    " & """drops""" & ":" & "["
                For j = 2 To Total2
                    If myrange2(j, 1) = myrange1(i, 1) Then
                        If once = 0 Then
                            .WriteText vbLf & "      ["
                            once = 1
                        End If
                        .WriteText vbLf & "       {"
                        .WriteText vbLf & "          " & """id""" & ": " & Replace(myrange2(j, 3), """", "\""") & ","         'id
                        .WriteText vbLf & "          " & """count""" & ": " & Replace(myrange2(j, 4), """", "\""") & ","         'count
                        .WriteText vbLf & "          " & """rate""" & ": " & Replace(myrange2(j, 5), """", "\""")                'rate
                        
                        If j + 1 > Total2 Then
                            nextj = j + 1 - Total2
                        Else
                            nextj = j + 1
                        End If
                        
                        If myrange2(j, 1) = myrange2(nextj, 1) Then
                            If myrange2(j, 2) = myrange2(nextj, 2) Then
                                .WriteText vbLf & "        },"
                            Else
                                .WriteText vbLf & "        }" & vbLf & "      ],"
                                once = 0
                            End If
                        Else
                            .WriteText vbLf & "        }" & vbLf & "      ]"
                            once = 0
                        End If

                    End If
                Next
                     
                .WriteText vbLf & "    ]"
                If i = Total1 Then
                    .WriteText vbLf & "  }"
                Else
                    .WriteText vbLf & "  },"
                End If
            Next
              
 
            .WriteText vbLf & "]"
            filename = ThisWorkbook.Path & "\" & Outname
            tempFilename = filename & ".json"
            .SaveToFile tempFilename, 2
            .Flush
            .Close
            MsgBox "�����ɹ�"
     End With
     Set objStream = Nothing
     ' ȥ��BOMͷ
    Call Utf8WithoutBom(tempFilename, filename)
    
    ' ɾ����ʱ�ļ�
    Kill tempFilename
    Name filename As filename & ".json"
End Sub

Sub ����Json����() '����UTF8�ı��ļ�
 'myrange = ActiveWorkbook.Names("schoolinfo").RefersToRange 'ͨ�������������ѡ������
 Dim Findname, Outname As String
 Dim wb As Workbook
 
myRange = Selection  'ͨ��ѡ��������ѡ������
cou = UBound(myRange)
For ric = 1 To cou


Findname = myRange(ric, 1)    'Ŀ�����
Outname = myRange(ric, 2)     '������
Set wb = GetObject(ThisWorkbook.Path & "\" & Findname & ".xlsx")    '��ȡĿ���������
Set st1 = wb.Worksheets("Value")
    
'used1 = Range(st1.Range("a1").End(xlDown), st1.Range("a1").End(xlToRight)) 'ͨ�������е�������������ѡ������
Total1 = st1.Range("A65535").End(xlUp).Row 'UBound(used1, 1) '��ȡ����
Fields1 = st1.Range("IV1").End(xlToLeft).Column 'UBound(used1, 2) '��ȡ����

ReDim myrange1(1 To Total1, 1 To Fields1)  '���ݸ�ֵΪtext����
For i = 1 To Total1
    For j = 1 To Fields1
        If IsNumeric(st1.Cells(i, j)) Then
            myrange1(i, j) = Replace(st1.Cells(i, j).Text, " ", "")
        Else
            myrange1(i, j) = Application.Substitute(st1.Cells(i, j).Text, Chr(10), "\n")
        End If
    Next
Next


For Each sht In wb.Worksheets

    If sht.Name = "TraditionalChinese" Then     'д�뷱��
        Set st2 = wb.Worksheets("TraditionalChinese")
        'used2 = Range(st2.Range("a1").End(xlDown), st2.Range("a1").End(xlToRight)) 'ͨ�������е�������������ѡ������
        Total2 = st2.Range("A65535").End(xlUp).Row 'UBound(used2, 1) '��ȡ����
        Fields2 = st2.Range("IV1").End(xlToLeft).Column 'UBound(used2, 2) '��ȡ����
        ReDim myrange2(1 To Total2, 1 To Fields2)  '���ݸ�ֵΪtext����
        For i = 1 To Total2
            For j = 1 To Fields2
                If IsNumeric(st2.Cells(i, j)) Then
                    myrange2(i, j) = Replace(st2.Cells(i, j).Text, " ", "")
                Else
                    myrange2(i, j) = Application.Substitute(st2.Cells(i, j).Text, Chr(10), "\n")
                End If
            Next
        Next
    Else
        Total2 = 0
        Fields2 = 0
    End If
    
    If sht.Name = "English" Then            'д��Ӣ��
        Set st3 = wb.Worksheets("English")
        'used2 = Range(st2.Range("a1").End(xlDown), st2.Range("a1").End(xlToRight)) 'ͨ�������е�������������ѡ������
        Total3 = st3.Range("A65535").End(xlUp).Row 'UBound(used2, 1) '��ȡ����
        Fields3 = st3.Range("IV1").End(xlToLeft).Column 'UBound(used2, 2) '��ȡ����
        ReDim myrange3(1 To Total3, 1 To Fields3)  '���ݸ�ֵΪtext����
        For i = 1 To Total3
            For j = 1 To Fields3
                If IsNumeric(st3.Cells(i, j)) Then
                    myrange3(i, j) = Replace(st3.Cells(i, j).Text, " ", "")
                Else
                    myrange3(i, j) = Application.Substitute(st3.Cells(i, j).Text, Chr(10), "\n")
                End If
            Next
        Next
    Else
        Total3 = 0
        Fields3 = 0
    End If
    
    If sht.Name = "Array" Then
        Set st0 = wb.Worksheets("Array")
        'used2 = Range(st2.Range("a1").End(xlDown), st2.Range("a1").End(xlToRight)) 'ͨ�������е�������������ѡ������
        Total0 = st0.Range("A65535").End(xlUp).Row 'UBound(used2, 1) '��ȡ����
        Fields0 = st0.Range("IV1").End(xlToLeft).Column 'UBound(used2, 2) '��ȡ����
        ReDim myrange0(1 To Total0, 1 To Fields0)  '���ݸ�ֵΪtext����
        For i = 1 To Total0
            For j = 1 To Fields0
                If IsNumeric(st0.Cells(i, j)) Then
                    myrange0(i, j) = Replace(st0.Cells(i, j).Text, " ", "")
                Else
                    myrange0(i, j) = Application.Substitute(st0.Cells(i, j).Text, Chr(10), "\n")
                End If
            Next
        Next
    Else
        Total0 = 0
        Fields0 = 0
    End If
    
Next
wb.Close            '�ر�Ŀ���
     
     
     Dim objStream As Object
     Set objStream = CreateObject("ADODB.Stream")
     
     With objStream
            .Type = 2
            .Charset = "UTF-8"
            .Open
            .WriteText "{"
    '����
            .WriteText vbLf & "  """ & "SimplifiedChinese" & """" & ":["
            For i = 2 To Total1
                .WriteText vbLf & "    " & "{" & vbLf & "      "
                For j = 1 To Fields1
                    If IsNumeric(myrange1(i, j)) And Left(myrange1(1, j), 8) <> "strCount" Then
                        .WriteText """" & myrange1(1, j) & """" & ": " & Replace(myrange1(i, j), """", "\""")           '����
                    ElseIf myrange1(i, j) = "FALSE" Then
                        .WriteText """" & myrange1(1, j) & """" & ": " & Replace("false", """", "\""")          '����ֵ
                    ElseIf myrange1(i, j) = "TRUE" Then
                        .WriteText """" & myrange1(1, j) & """" & ": " & Replace("true", """", "\""")          '����ֵ
                    Else
                        .WriteText """" & myrange1(1, j) & """: """ & Replace(myrange1(i, j), """", "\""") & """"         '�ı�
                    End If
                     If j <> Fields1 Then
                        .WriteText "," & vbLf & "      "
                     End If
                Next
                
                If Total0 And Fields0 > 0 Then
                    For n = 2 To Fields0
                        fir = 0
                        .WriteText "," & vbLf & "    " & """" & Replace(myrange0(1, n), """", "\""") & """" & ": " & "["
                        For m = 2 To Total2
                          If myrange0(m, n) <> "" Then
                            If myrange0(m, 1) = myrange1(i, 1) And fir = 0 Then
                                fir = 1
                                If IsNumeric(myrange0(m, n)) Then
                                    .WriteText vbLf & "      " & Replace(myrange0(m, n), """", "\""")
                                ElseIf myrange0(m, n) = "FALSE" Then
                                    .WriteText vbLf & "      " & Replace("false", """", "\""")
                                ElseIf myrange0(m, n) = "TRUE" Then
                                    .WriteText vbLf & "      " & Replace("true", """", "\""")
                                Else
                                    .WriteText vbLf & "      " & """" & Replace(myrange0(m, n), """", "\""") & """"
                                End If
                            ElseIf myrange0(m, 1) = myrange1(i, 1) And fir = 1 Then
                                If IsNumeric(myrange0(m, n)) Then
                                    .WriteText "," & vbLf & "      " & Replace(myrange0(m, n), """", "\""")
                                ElseIf myrange0(m, n) = "FALSE" Then
                                    .WriteText "," & vbLf & "      " & Replace("false", """", "\""")
                                ElseIf myrange0(m, n) = "TRUE" Then
                                    .WriteText "," & vbLf & "      " & Replace("true", """", "\""")
                                Else
                                    .WriteText "," & vbLf & "      " & """" & Replace(myrange0(m, n), """", "\""") & """"
                                End If
                            End If
                            End If
                        Next
                        If n = Fields0 Then
                            .WriteText vbLf & "    " & "]"
                        Else
                            .WriteText vbLf & "    " & "]"
                        End If
                    Next
                End If
                
                If i = Total1 Then
                        .WriteText vbLf & "    " & "}"
                Else
                        .WriteText vbLf & "    " & "},"
                End If
            Next
            .WriteText vbLf & "  ]" & ","
                
    '����
            .WriteText vbLf & "  """ & "TraditionalChinese" & """" & ":["
            If Total2 And Fields2 > 0 Then
                For i = 2 To Total2
                    .WriteText vbLf & "    " & "{" & vbLf & "      "
                    For j = 1 To Fields2
                        If IsNumeric(myrange2(i, j)) And Left(myrange2(1, j), 8) <> "strCount" Then
                            .WriteText """" & myrange2(1, j) & """" & ": " & Replace(myrange2(i, j), """", "\""")           '����
                        ElseIf myrange2(i, j) = "FALSE" Then
                            .WriteText """" & myrange2(1, j) & """" & ": " & Replace("false", """", "\""")          '����ֵ
                        ElseIf myrange2(i, j) = "TRUE" Then
                            .WriteText """" & myrange2(1, j) & """" & ": " & Replace("true", """", "\""")          '����ֵ
                        Else
                            .WriteText """" & myrange2(1, j) & """: """ & Replace(myrange2(i, j), """", "\""") & """"         '�ı�
                        End If
                        If j <> Fields2 Then
                            .WriteText "," & vbLf & "      "
                        End If
                    Next
                    
                    If Total0 And Fields0 > 0 Then
                        For n = 2 To Fields0
                            fir = 0
                            .WriteText "," & vbLf & "    " & """" & Replace(myrange0(1, n), """", "\""") & """" & ": " & "["
                            For m = 2 To Total2
                              If myrange0(m, n) <> "" Then
                                If myrange0(m, 1) = myrange1(i, 1) And fir = 0 Then
                                    fir = 1
                                    If IsNumeric(myrange0(m, n)) Then
                                        .WriteText vbLf & "      " & Replace(myrange0(m, n), """", "\""")
                                    ElseIf myrange0(m, n) = "FALSE" Then
                                        .WriteText vbLf & "      " & Replace("false", """", "\""")
                                    ElseIf myrange0(m, n) = "TRUE" Then
                                        .WriteText vbLf & "      " & Replace("true", """", "\""")
                                    Else
                                        .WriteText vbLf & "      " & """" & Replace(myrange0(m, n), """", "\""") & """"
                                    End If
                                ElseIf myrange0(m, 1) = myrange1(i, 1) And fir = 1 Then
                                    If IsNumeric(myrange0(m, n)) Then
                                        .WriteText "," & vbLf & "      " & Replace(myrange0(m, n), """", "\""")
                                    ElseIf myrange0(m, n) = "FALSE" Then
                                        .WriteText "," & vbLf & "      " & Replace("false", """", "\""")
                                    ElseIf myrange0(m, n) = "TRUE" Then
                                        .WriteText "," & vbLf & "      " & Replace("true", """", "\""")
                                    Else
                                        .WriteText "," & vbLf & "      " & """" & Replace(myrange0(m, n), """", "\""") & """"
                                    End If
                                End If
                                End If
                            Next
                            If n = Fields0 Then
                                .WriteText vbLf & "    " & "]"
                            Else
                                .WriteText vbLf & "    " & "]"
                            End If
                        Next
                    End If
                
                    If i = Total2 Then
                            .WriteText vbLf & "    " & "}"
                    Else
                            .WriteText vbLf & "    " & "},"
                    End If
                Next
            End If
            .WriteText vbLf & "  ]" & ","
      
      
      'Ӣ��
            .WriteText vbLf & "  """ & "English" & """" & ":["
            If Total3 And Fields3 > 0 Then
                For i = 2 To Total3
                    .WriteText vbLf & "    " & "{" & vbLf & "      "
                    For j = 1 To Fields3
                        If IsNumeric(myrange3(i, j)) And Left(myrange3(1, j), 8) <> "strCount" Then
                            .WriteText """" & myrange3(1, j) & """" & ": " & Replace(myrange3(i, j), """", "\""")           '����
                        ElseIf myrange3(i, j) = "FALSE" Then
                            .WriteText """" & myrange3(1, j) & """" & ": " & Replace("false", """", "\""")          '����ֵ
                        ElseIf myrange3(i, j) = "TRUE" Then
                            .WriteText """" & myrange3(1, j) & """" & ": " & Replace("true", """", "\""")          '����ֵ
                        Else
                            .WriteText """" & myrange3(1, j) & """: """ & Replace(myrange3(i, j), """", "\""") & """"         '�ı�
                        End If
                        If j <> Fields3 Then
                            .WriteText "," & vbLf & "      "
                        End If
                    Next
                    
                    If Total0 And Fields0 > 0 Then
                        For n = 2 To Fields0
                            fir = 0
                            .WriteText "," & vbLf & "    " & """" & Replace(myrange0(1, n), """", "\""") & """" & ": " & "["
                            For m = 2 To Total2
                              If myrange0(m, n) <> "" Then
                                If myrange0(m, 1) = myrange1(i, 1) And fir = 0 Then
                                    fir = 1
                                    If IsNumeric(myrange0(m, n)) Then
                                        .WriteText vbLf & "      " & Replace(myrange0(m, n), """", "\""")
                                    ElseIf myrange0(m, n) = "FALSE" Then
                                        .WriteText vbLf & "      " & Replace("false", """", "\""")
                                    ElseIf myrange0(m, n) = "TRUE" Then
                                        .WriteText vbLf & "      " & Replace("true", """", "\""")
                                    Else
                                        .WriteText vbLf & "      " & """" & Replace(myrange0(m, n), """", "\""") & """"
                                    End If
                                ElseIf myrange0(m, 1) = myrange1(i, 1) And fir = 1 Then
                                    If IsNumeric(myrange0(m, n)) Then
                                        .WriteText "," & vbLf & "      " & Replace(myrange0(m, n), """", "\""")
                                    ElseIf myrange0(m, n) = "FALSE" Then
                                        .WriteText "," & vbLf & "      " & Replace("false", """", "\""")
                                    ElseIf myrange0(m, n) = "TRUE" Then
                                        .WriteText "," & vbLf & "      " & Replace("true", """", "\""")
                                    Else
                                        .WriteText "," & vbLf & "      " & """" & Replace(myrange0(m, n), """", "\""") & """"
                                    End If
                                End If
                                End If
                            Next
                            If n = Fields0 Then
                                .WriteText vbLf & "    " & "]"
                            Else
                                .WriteText vbLf & "    " & "]"
                            End If
                        Next
                    End If
                    
                    If i = Total3 Then
                            .WriteText vbLf & "    " & "}"
                    Else
                            .WriteText vbLf & "    " & "},"
                    End If
                Next
            End If
            .WriteText vbLf & "  ]"
            
            .WriteText vbLf & "}"
            filename = ThisWorkbook.Path & "\" & Outname
            tempFilename = filename & ".json"
            .SaveToFile tempFilename, 2
            .Flush
            .Close
            MsgBox Findname & vbCrLf & "�����ɹ�"
     End With
     Set objStream = Nothing
     ' ȥ��BOMͷ
    Call Utf8WithoutBom(tempFilename, filename)
    
    ' ɾ����ʱ�ļ�
    Kill tempFilename
    Name filename As filename & ".json"
    
Next
End Sub
' ȥ��BOMͷ
' �ο�http://club.excelhome.net/forum.php?mod=viewthread&tid=1240816
Function Utf8WithoutBom(ByVal getPath As String, ByVal putPath As String)
    
    Dim getFileNum As Integer
    Dim putFileNum As Integer
    getFileNum = 1
    putFileNum = 2

    Open getPath For Binary As #getFileNum
    Open putPath For Binary As #putFileNum
    
    Dim fileByte As Byte
    Seek #getFileNum, 4
    
    For i = 1 To LOF(getFileNum) - 3
        Get #getFileNum, , fileByte
        Put #putFileNum, , fileByte
    Next i
    
    Close #getFileNum
    Close #putFileNum
End Function


Sub ���������޸�()
   p = ThisWorkbook.Path & "\"
   f = Dir(p & "*.xlsx")
   Application.DisplayAlerts = False
   Do While f <> ""
       If f <> ThisWorkbook.Name Then
           Set wb = GetObject(p & f)
           wb.Sheets("Value").Name = "SimplifiedChinese"
           Application.Windows(wb.Name).Visible = True
           Workbooks(f).Close True
      End If
      f = Dir
   Loop
   Set wb = Nothing
   Application.DisplayAlerts = True
   MsgBox "�޸���ɣ�"
End Sub
'ͨ��getobject��dao��Excel�ļ�ֻҪ���޸ģ�д����dao����󣬾�ֻר����VBE�п��������û�������ȴ��������
'����������Excel����ȥ�ֶ��򿪴��ļ���Ҳ��ʲô�����������������û��������⣡���Ҫ���������⣬������wb.close ǰ��һ��Application.Windows(wb.name).Visible = True
'�����Ϊʹ��with workbooks.open(path)���������getobject��ͨ����̨�����ġ�