' Excel文件比较工具
' 功能：比较两个Excel文件，标记出不同的行和单元格
' 使用方法：在Excel中运行此宏

Option Explicit

Sub CompareExcelFiles()
    Dim filePathA As String, filePathB As String
    Dim wbA As Workbook, wbB As Workbook
    Dim wsA As Worksheet, wsB As Worksheet
    Dim lastRowA As Long, lastRowB As Long, lastColA As Long, lastColB As Long
    Dim maxRow As Long, maxCol As Long
    Dim i As Long, j As Long, k As Long
    Dim idA As Variant, idB As Variant
    Dim foundMatch As Boolean
    Dim matchedRowsA() As Boolean, matchedRowsB() As Boolean
    
    ' 检测操作系统类型
    Dim isMac As Boolean
    #If Mac Then
        isMac = True
    #Else
        isMac = False
    #End If
    
    ' 提示用户选择第一个Excel文件
    If isMac Then
        filePathA = MacScript("choose file of type {""org.openxmlformats.spreadsheetml.sheet"", ""com.microsoft.Excel.Sheet""} with prompt ""请选择第一个Excel文件""")
        ' 转换Mac路径格式为VBA可识别的格式
        filePathA = ConvertMacPath(filePathA)
        If filePathA = "" Then
            MsgBox "未选择文件或文件路径无效", vbExclamation
            Exit Sub
        End If
    Else
        filePathA = Application.GetOpenFilename("Excel文件 (*.xlsx;*.xls),*.xlsx;*.xls", , "请选择第一个Excel文件")
        If filePathA = "False" Then Exit Sub
    End If
    
    ' 提示用户选择第二个Excel文件
    If isMac Then
        filePathB = MacScript("choose file of type {""org.openxmlformats.spreadsheetml.sheet"", ""com.microsoft.Excel.Sheet""} with prompt ""请选择第二个Excel文件""")
        ' 转换Mac路径格式为VBA可识别的格式
        filePathB = ConvertMacPath(filePathB)
        If filePathB = "" Then
            MsgBox "未选择文件或文件路径无效", vbExclamation
            Exit Sub
        End If
    Else
        filePathB = Application.GetOpenFilename("Excel文件 (*.xlsx;*.xls),*.xlsx;*.xls", , "请选择第二个Excel文件")
        If filePathB = "False" Then Exit Sub
    End If
    
    ' 打开选择的Excel文件
    On Error Resume Next
    Set wbA = Workbooks.Open(filePathA, ReadOnly:=False)
    If Err.Number <> 0 Then
        MsgBox "无法以可写模式打开第一个文件: " & Err.Description, vbExclamation
        Err.Clear
        Set wbA = Workbooks.Open(filePathA, ReadOnly:=True)
    End If
    
    Set wbB = Workbooks.Open(filePathB, ReadOnly:=False)
    If Err.Number <> 0 Then
        MsgBox "无法以可写模式打开第二个文件: " & Err.Description, vbExclamation
        Err.Clear
        Set wbB = Workbooks.Open(filePathB, ReadOnly:=True)
    End If
    On Error GoTo 0
    
    ' 假设比较第一个工作表
    Set wsA = wbA.Worksheets(1)
    Set wsB = wbB.Worksheets(1)
    
    ' 获取数据范围
    lastRowA = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
    lastRowB = wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row
    lastColA = wsA.Cells(1, wsA.Columns.Count).End(xlToLeft).Column
    lastColB = wsB.Cells(1, wsB.Columns.Count).End(xlToLeft).Column
    
    maxRow = Application.WorksheetFunction.Max(lastRowA, lastRowB)
    maxCol = Application.WorksheetFunction.Max(lastColA, lastColB)
    
    ' 初始化匹配行的跟踪数组
    ReDim matchedRowsA(1 To lastRowA) As Boolean
    ReDim matchedRowsB(1 To lastRowB) As Boolean
    
    ' 比较行并标记差异
    
    ' 首先处理文件A中的每一行
    For i = 2 To lastRowA ' 从第2行开始，跳过标题行
        idA = wsA.Cells(i, 1).Value ' 获取标识符
        foundMatch = False
        
        ' 在文件B中查找匹配的行
        For j = 2 To lastRowB
            idB = wsB.Cells(j, 1).Value
            
            If idA = idB Then ' 找到匹配的行
                foundMatch = True
                matchedRowsA(i) = True
                matchedRowsB(j) = True
                
                ' 比较单元格并标记不同的单元格为黄色
                For k = 2 To Application.WorksheetFunction.Min(lastColA, lastColB) ' 从第2列开始，跳过标识符列
                    If wsA.Cells(i, k).Value <> wsB.Cells(j, k).Value Then
                        wsA.Cells(i, k).Interior.Color = RGB(255, 255, 0) ' 在文件A中标记为黄色
                        wsB.Cells(j, k).Interior.Color = RGB(255, 255, 0) ' 在文件B中标记为黄色
                    End If
                Next k
                Exit For
            End If
        Next j
        
        ' 如果在文件B中没有找到匹配的行，在文件A中标记为红色
        If Not foundMatch Then
            ' 标记整行为红色
            wsA.Range(wsA.Cells(i, 1), wsA.Cells(i, lastColA)).Interior.Color = RGB(255, 0, 0) ' 红色
        End If
    Next i
    
    ' 处理文件B中未匹配的行
    For j = 2 To lastRowB
        If Not matchedRowsB(j) Then
            ' 标记整行为红色
            wsB.Range(wsB.Cells(j, 1), wsB.Cells(j, lastColB)).Interior.Color = RGB(255, 0, 0) ' 红色
        End If
    Next j
    
    ' 自动调整列宽以提高可读性
    wsA.Columns.AutoFit
    wsB.Columns.AutoFit
    
    ' 保存并关闭源文件
    On Error Resume Next
    wbA.Save
    If Err.Number <> 0 Then
        MsgBox "无法保存第一个文件: " & Err.Description & "。文件可能是只读的。", vbExclamation
        Err.Clear
    End If
    
    wbB.Save
    If Err.Number <> 0 Then
        MsgBox "无法保存第二个文件: " & Err.Description & "。文件可能是只读的。", vbExclamation
        Err.Clear
    End If
    
    wbA.Close SaveChanges:=(Err.Number = 0)
    wbB.Close SaveChanges:=(Err.Number = 0)
    On Error GoTo 0
    
    ' 显示结果
    MsgBox "比较完成！红色行表示仅在一个文件中存在的行，黄色单元格表示内容不一致的单元格。已在原文件中标记。", vbInformation
End Sub

' 将Mac路径格式转换为VBA可识别的格式
Function ConvertMacPath(macPath As String) As String
    Dim vbaPath As String
    
    ' 检查路径是否为空
    If Len(macPath) = 0 Then
        ConvertMacPath = ""
        Exit Function
    End If
    
    ' 移除可能存在的前缀
    If Left(macPath, 4) = "Macintosh HD:" Then
        macPath = Mid(macPath, 14)
    End If
    
    ' 替换冒号为斜杠
    vbaPath = Replace(macPath, ":", "/")
    
    ' 确保路径以斜杠开始
    If Left(vbaPath, 1) <> "/" Then
        vbaPath = "/" & vbaPath
    End If
    
    ' 修复Mac系统下路径可能包含电脑名的问题
    ' 检查路径是否包含类似 /电脑名/Users/ 的格式
    Dim usersPos As Long
    usersPos = InStr(1, vbaPath, "/Users/", vbTextCompare)
    
    ' 如果找到 /Users/ 且不是在路径开头
    If usersPos > 1 Then
        ' 移除电脑名部分，只保留 /Users/ 及之后的部分
        vbaPath = Mid(vbaPath, usersPos)
    End If
    
    ConvertMacPath = vbaPath
End Function