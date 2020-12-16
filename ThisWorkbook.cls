VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub Json导出() '创建UTF8文本文件
Attribute Json导出.VB_ProcData.VB_Invoke_Func = "e\n14"
 'myrange = ActiveWorkbook.Names("schoolinfo").RefersToRange '通过定义的名称来选择数据
 Dim Findname, Outname As String
 Dim wb As Workbook
 
myRange = Selection  '通过选中区域来选择数据
cou = UBound(myRange)
For ric = 1 To cou


Findname = myRange(ric, 1)    '目标表名
Outname = myRange(ric, 2)     '导出名
Set wb = GetObject(ThisWorkbook.Path & "\" & Findname & ".xlsx")    '获取目标表内数据
Set st1 = wb.Worksheets("Value")
    
'used1 = Range(st1.Range("a1").End(xlDown), st1.Range("a1").End(xlToRight)) '通过标题行的最大行最大列来选择数据
Total1 = st1.Range("A65535").End(xlUp).Row 'UBound(used1, 1) '获取行数
Fields1 = st1.Range("IV1").End(xlToLeft).Column 'UBound(used1, 2) '获取列数

ReDim myrange1(1 To Total1, 1 To Fields1)  '数据赋值为text数组
For i = 1 To Total1
    For j = 1 To Fields1
        If IsNumeric(st1.Cells(i, j)) Then
            myrange1(i, j) = Replace(st1.Cells(i, j).Text, " ", "")
        Else
            myrange1(i, j) = Application.Substitute(st1.Cells(i, j).Text, Chr(10), "\n")
        End If
    Next
Next



If wb.Worksheets.Count = 2 Then            '判断是否有Array表
Set st2 = wb.Worksheets("Array")
    'used2 = Range(st2.Range("a1").End(xlDown), st2.Range("a1").End(xlToRight)) '通过标题行的最大行最大列来选择数据
    Total2 = st2.Range("A65535").End(xlUp).Row 'UBound(used2, 1) '获取行数
    Fields2 = st2.Range("IV1").End(xlToLeft).Column 'UBound(used2, 2) '获取列数
    ReDim myrange2(1 To Total2, 1 To Fields2)  '数据赋值为text数组
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


wb.Close            '关闭目标表
     
     
     Dim objStream As Object
     Set objStream = CreateObject("ADODB.Stream")
     
     With objStream
            .Type = 2
            .Charset = "UTF-8"
            .Open
            .WriteText "["
    'Value表
            For i = 2 To Total1
                .WriteText vbLf & "  " & "{" & vbLf & "    "
                For j = 1 To Fields1
                    If IsNumeric(myrange1(i, j)) And Left(myrange1(1, j), 8) <> "strCount" Then
                        .WriteText """" & myrange1(1, j) & """" & ": " & Replace(myrange1(i, j), """", "\""")           '数字
                    ElseIf myrange1(i, j) = "FALSE" Then
                        .WriteText """" & myrange1(1, j) & """" & ": " & Replace("false", """", "\""")          '布尔值
                    ElseIf myrange1(i, j) = "TRUE" Then
                        .WriteText """" & myrange1(1, j) & """" & ": " & Replace("true", """", "\""")          '布尔值
                    Else
                        .WriteText """" & myrange1(1, j) & """: """ & Replace(myrange1(i, j), """", "\""") & """"         '文本
                    End If
                     If j <> Fields1 Then
                        .WriteText "," & vbLf & "    "
                     End If
                Next
    'Array表
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
            MsgBox Findname & vbCrLf & "导出成功"
     End With
     Set objStream = Nothing
     ' 去除BOM头
    Call Utf8WithoutBom(tempFilename, filename)
    
    ' 删除临时文件
    Kill tempFilename
    Name filename As filename & ".json"
    
Next
End Sub


Sub 掉落Json导出() '创建UTF8文本文件
 Dim Findname, Outname As String
 Dim wb As Workbook
 
myRange = Selection  '通过选中区域来选择数据
Findname = myRange(1, 1)    '目标表名
Outname = myRange(1, 2)     '导出名
 
 Set wb = GetObject(ThisWorkbook.Path & "\" & Findname & ".xlsx")    '获取目标表内数据
 Set st1 = wb.Worksheets("Value")
    
'used1 = Range(st1.Range("a1").End(xlDown), st1.Range("a1").End(xlToRight)) '通过标题行的最大行最大列来选择数据
Total1 = st1.Range("A65535").End(xlUp).Row 'UBound(used1, 1) '获取行数
Fields1 = st1.Range("IV1").End(xlToLeft).Column 'UBound(used1, 2) '获取列数

ReDim myrange1(1 To Total1, 1 To Fields1)  '数据赋值为text数组
For i = 1 To Total1
    For j = 1 To Fields1
        myrange1(i, j) = st1.Cells(i, j).Text
    Next
Next

If wb.Worksheets.Count = 2 Then            '判断是否有Array表
Set st2 = wb.Worksheets("Array")
    'used2 = Range(st2.Range("a1").End(xlDown), st2.Range("a1").End(xlToRight)) '通过标题行的最大行最大列来选择数据
    Total2 = st2.Range("A65535").End(xlUp).Row 'UBound(used2, 1) '获取行数
    Fields2 = st2.Range("IV1").End(xlToLeft).Column 'UBound(used2, 2) '获取列数
    ReDim myrange2(1 To Total2, 1 To Fields2)  '数据赋值为text数组
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
    'Value表
             For i = 2 To Total1
                .WriteText vbLf & "  " & "{" & vbLf & "    "
                If IsNumeric(myrange2(i, 1)) Then
                    .WriteText """" & myrange1(1, 1) & """" & ": " & Replace(myrange1(i, 1), """", "\""")           '数字
                Else
                    .WriteText """" & myrange1(1, 1) & """: """ & Replace(myrange1(i, 1), """", "\""") & """"         '文本
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
            MsgBox "导出成功"
     End With
     Set objStream = Nothing
     ' 去除BOM头
    Call Utf8WithoutBom(tempFilename, filename)
    
    ' 删除临时文件
    Kill tempFilename
    Name filename As filename & ".json"
End Sub

Sub 描述Json导出() '创建UTF8文本文件
 'myrange = ActiveWorkbook.Names("schoolinfo").RefersToRange '通过定义的名称来选择数据
 Dim Findname, Outname As String
 Dim wb As Workbook
 
myRange = Selection  '通过选中区域来选择数据
cou = UBound(myRange)
For ric = 1 To cou


Findname = myRange(ric, 1)    '目标表名
Outname = myRange(ric, 2)     '导出名
Set wb = GetObject(ThisWorkbook.Path & "\" & Findname & ".xlsx")    '获取目标表内数据
Set st1 = wb.Worksheets("Value")
    
'used1 = Range(st1.Range("a1").End(xlDown), st1.Range("a1").End(xlToRight)) '通过标题行的最大行最大列来选择数据
Total1 = st1.Range("A65535").End(xlUp).Row 'UBound(used1, 1) '获取行数
Fields1 = st1.Range("IV1").End(xlToLeft).Column 'UBound(used1, 2) '获取列数

ReDim myrange1(1 To Total1, 1 To Fields1)  '数据赋值为text数组
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

    If sht.Name = "TraditionalChinese" Then     '写入繁中
        Set st2 = wb.Worksheets("TraditionalChinese")
        'used2 = Range(st2.Range("a1").End(xlDown), st2.Range("a1").End(xlToRight)) '通过标题行的最大行最大列来选择数据
        Total2 = st2.Range("A65535").End(xlUp).Row 'UBound(used2, 1) '获取行数
        Fields2 = st2.Range("IV1").End(xlToLeft).Column 'UBound(used2, 2) '获取列数
        ReDim myrange2(1 To Total2, 1 To Fields2)  '数据赋值为text数组
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
    
    If sht.Name = "English" Then            '写入英文
        Set st3 = wb.Worksheets("English")
        'used2 = Range(st2.Range("a1").End(xlDown), st2.Range("a1").End(xlToRight)) '通过标题行的最大行最大列来选择数据
        Total3 = st3.Range("A65535").End(xlUp).Row 'UBound(used2, 1) '获取行数
        Fields3 = st3.Range("IV1").End(xlToLeft).Column 'UBound(used2, 2) '获取列数
        ReDim myrange3(1 To Total3, 1 To Fields3)  '数据赋值为text数组
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
        'used2 = Range(st2.Range("a1").End(xlDown), st2.Range("a1").End(xlToRight)) '通过标题行的最大行最大列来选择数据
        Total0 = st0.Range("A65535").End(xlUp).Row 'UBound(used2, 1) '获取行数
        Fields0 = st0.Range("IV1").End(xlToLeft).Column 'UBound(used2, 2) '获取列数
        ReDim myrange0(1 To Total0, 1 To Fields0)  '数据赋值为text数组
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
wb.Close            '关闭目标表
     
     
     Dim objStream As Object
     Set objStream = CreateObject("ADODB.Stream")
     
     With objStream
            .Type = 2
            .Charset = "UTF-8"
            .Open
            .WriteText "{"
    '简中
            .WriteText vbLf & "  """ & "SimplifiedChinese" & """" & ":["
            For i = 2 To Total1
                .WriteText vbLf & "    " & "{" & vbLf & "      "
                For j = 1 To Fields1
                    If IsNumeric(myrange1(i, j)) And Left(myrange1(1, j), 8) <> "strCount" Then
                        .WriteText """" & myrange1(1, j) & """" & ": " & Replace(myrange1(i, j), """", "\""")           '数字
                    ElseIf myrange1(i, j) = "FALSE" Then
                        .WriteText """" & myrange1(1, j) & """" & ": " & Replace("false", """", "\""")          '布尔值
                    ElseIf myrange1(i, j) = "TRUE" Then
                        .WriteText """" & myrange1(1, j) & """" & ": " & Replace("true", """", "\""")          '布尔值
                    Else
                        .WriteText """" & myrange1(1, j) & """: """ & Replace(myrange1(i, j), """", "\""") & """"         '文本
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
                
    '繁中
            .WriteText vbLf & "  """ & "TraditionalChinese" & """" & ":["
            If Total2 And Fields2 > 0 Then
                For i = 2 To Total2
                    .WriteText vbLf & "    " & "{" & vbLf & "      "
                    For j = 1 To Fields2
                        If IsNumeric(myrange2(i, j)) And Left(myrange2(1, j), 8) <> "strCount" Then
                            .WriteText """" & myrange2(1, j) & """" & ": " & Replace(myrange2(i, j), """", "\""")           '数字
                        ElseIf myrange2(i, j) = "FALSE" Then
                            .WriteText """" & myrange2(1, j) & """" & ": " & Replace("false", """", "\""")          '布尔值
                        ElseIf myrange2(i, j) = "TRUE" Then
                            .WriteText """" & myrange2(1, j) & """" & ": " & Replace("true", """", "\""")          '布尔值
                        Else
                            .WriteText """" & myrange2(1, j) & """: """ & Replace(myrange2(i, j), """", "\""") & """"         '文本
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
      
      
      '英语
            .WriteText vbLf & "  """ & "English" & """" & ":["
            If Total3 And Fields3 > 0 Then
                For i = 2 To Total3
                    .WriteText vbLf & "    " & "{" & vbLf & "      "
                    For j = 1 To Fields3
                        If IsNumeric(myrange3(i, j)) And Left(myrange3(1, j), 8) <> "strCount" Then
                            .WriteText """" & myrange3(1, j) & """" & ": " & Replace(myrange3(i, j), """", "\""")           '数字
                        ElseIf myrange3(i, j) = "FALSE" Then
                            .WriteText """" & myrange3(1, j) & """" & ": " & Replace("false", """", "\""")          '布尔值
                        ElseIf myrange3(i, j) = "TRUE" Then
                            .WriteText """" & myrange3(1, j) & """" & ": " & Replace("true", """", "\""")          '布尔值
                        Else
                            .WriteText """" & myrange3(1, j) & """: """ & Replace(myrange3(i, j), """", "\""") & """"         '文本
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
            MsgBox Findname & vbCrLf & "导出成功"
     End With
     Set objStream = Nothing
     ' 去除BOM头
    Call Utf8WithoutBom(tempFilename, filename)
    
    ' 删除临时文件
    Kill tempFilename
    Name filename As filename & ".json"
    
Next
End Sub
' 去除BOM头
' 参考http://club.excelhome.net/forum.php?mod=viewthread&tid=1240816
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


Sub 工作表名修改()
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
   MsgBox "修改完成！"
End Sub
'通过getobject打开dao的Excel文件只要被修改（写）并dao保存后，就只专能在VBE中看到，但用户界属面却看不到。
'就算你重启Excel，再去手动打开此文件，也是什么都看不到。不保存就没有这个问题！如果要解决这个问题，必须在wb.close 前加一句Application.Windows(wb.name).Visible = True
'建议改为使用with workbooks.open(path)这个函数，getobject是通过后台处理的。
