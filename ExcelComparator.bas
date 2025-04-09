' Excel File Comparison Tool
' Function: Compare two Excel files and mark different rows and cells
' Usage: Run this macro in Excel

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
    
    ' Detect operating system type
    Dim isMac As Boolean
    #If Mac Then
        isMac = True
    #Else
        isMac = False
    #End If
    
    ' Prompt user to select the first Excel file
    If isMac Then
        filePathA = MacScript("choose file of type {""org.openxmlformats.spreadsheetml.sheet"", ""com.microsoft.Excel.Sheet""} with prompt ""Select the first Excel file""")
        ' Convert Mac path format to VBA recognizable format
        filePathA = ConvertMacPath(filePathA)
        If filePathA = "" Then
            MsgBox "No file selected or invalid file path", vbExclamation
            Exit Sub
        End If
    Else
        filePathA = Application.GetOpenFilename("Excel Files (*.xlsx;*.xls),*.xlsx;*.xls", , "Select the first Excel file")
        If filePathA = "False" Then Exit Sub
    End If
    
    ' Prompt user to select the second Excel file
    If isMac Then
        filePathB = MacScript("choose file of type {""org.openxmlformats.spreadsheetml.sheet"", ""com.microsoft.Excel.Sheet""} with prompt ""Select the second Excel file""")
        ' Convert Mac path format to VBA recognizable format
        filePathB = ConvertMacPath(filePathB)
        If filePathB = "" Then
            MsgBox "No file selected or invalid file path", vbExclamation
            Exit Sub
        End If
    Else
        filePathB = Application.GetOpenFilename("Excel Files (*.xlsx;*.xls),*.xlsx;*.xls", , "Select the second Excel file")
        If filePathB = "False" Then Exit Sub
    End If
    
    ' Open selected Excel files
    On Error Resume Next
    Set wbA = Workbooks.Open(filePathA, ReadOnly:=False)
    If Err.Number <> 0 Then
        MsgBox "Cannot open the first file in write mode: " & Err.Description, vbExclamation
        Err.Clear
        Set wbA = Workbooks.Open(filePathA, ReadOnly:=True)
    End If
    
    Set wbB = Workbooks.Open(filePathB, ReadOnly:=False)
    If Err.Number <> 0 Then
        MsgBox "Cannot open the second file in write mode: " & Err.Description, vbExclamation
        Err.Clear
        Set wbB = Workbooks.Open(filePathB, ReadOnly:=True)
    End If
    On Error GoTo 0
    
    ' Compare the first worksheet
    Set wsA = wbA.Worksheets(1)
    Set wsB = wbB.Worksheets(1)
    
    ' Get data range
    lastRowA = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
    lastRowB = wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row
    lastColA = wsA.Cells(1, wsA.Columns.Count).End(xlToLeft).Column
    lastColB = wsB.Cells(1, wsB.Columns.Count).End(xlToLeft).Column
    
    maxRow = Application.WorksheetFunction.Max(lastRowA, lastRowB)
    maxCol = Application.WorksheetFunction.Max(lastColA, lastColB)
    
    ' Initialize arrays to track matched rows
    ReDim matchedRowsA(1 To lastRowA) As Boolean
    ReDim matchedRowsB(1 To lastRowB) As Boolean
    
    ' Compare rows and mark differences
    
    ' First process each row in file A
    For i = 2 To lastRowA ' Start from row 2, skip header row
        idA = wsA.Cells(i, 1).Value ' Get identifier
        foundMatch = False
        
        ' Find matching row in file B
        For j = 2 To lastRowB
            idB = wsB.Cells(j, 1).Value
            
            If idA = idB Then ' Found matching row
                foundMatch = True
                matchedRowsA(i) = True
                matchedRowsB(j) = True
                
                ' Compare cells and mark different cells as yellow
                For k = 2 To Application.WorksheetFunction.Min(lastColA, lastColB) ' Start from column 2, skip identifier column
                    If wsA.Cells(i, k).Value <> wsB.Cells(j, k).Value Then
                        wsA.Cells(i, k).Interior.Color = RGB(255, 255, 0) ' Mark as yellow in file A
                        wsB.Cells(j, k).Interior.Color = RGB(255, 255, 0) ' Mark as yellow in file B
                    End If
                Next k
                Exit For
            End If
        Next j
        
        ' If no matching row found in file B, mark as red in file A
        If Not foundMatch Then
            ' Mark entire row as red
            wsA.Range(wsA.Cells(i, 1), wsA.Cells(i, lastColA)).Interior.Color = RGB(255, 0, 0) ' Red
        End If
    Next i
    
    ' Process unmatched rows in file B
    For j = 2 To lastRowB
        If Not matchedRowsB(j) Then
            ' Mark entire row as red
            wsB.Range(wsB.Cells(j, 1), wsB.Cells(j, lastColB)).Interior.Color = RGB(255, 0, 0) ' Red
        End If
    Next j
    
    ' Auto-fit columns to improve readability
    wsA.Columns.AutoFit
    wsB.Columns.AutoFit
    
    ' Save and close source files
    On Error Resume Next
    wbA.Save
    If Err.Number <> 0 Then
        MsgBox "Cannot save the first file: " & Err.Description & ". The file may be read-only.", vbExclamation
        Err.Clear
    End If
    
    wbB.Save
    If Err.Number <> 0 Then
        MsgBox "Cannot save the second file: " & Err.Description & ". The file may be read-only.", vbExclamation
        Err.Clear
    End If
    
    wbA.Close SaveChanges:=(Err.Number = 0)
    wbB.Close SaveChanges:=(Err.Number = 0)
    On Error GoTo 0
    
    ' Display results
    MsgBox "Comparison completed! Red rows indicate rows that exist in only one file, yellow cells indicate cells with inconsistent content. Marked in the original files.", vbInformation
End Sub

' Convert Mac path format to VBA recognizable format
Function ConvertMacPath(macPath As String) As String
    Dim vbaPath As String
    
    ' Check if path is empty
    If Len(macPath) = 0 Then
        ConvertMacPath = ""
        Exit Function
    End If
    
    ' Remove possible prefix
    If Left(macPath, 4) = "Macintosh HD:" Then
        macPath = Mid(macPath, 14)
    End If
    
    ' Replace colons with slashes
    vbaPath = Replace(macPath, ":", "/")
    
    ' Ensure path starts with a slash
    If Left(vbaPath, 1) <> "/" Then
        vbaPath = "/" & vbaPath
    End If
    
    ' Fix issue with Mac paths possibly containing computer name
    ' Check if path contains format like /ComputerName/Users/
    Dim usersPos As Long
    usersPos = InStr(1, vbaPath, "/Users/", vbTextCompare)
    
    ' If /Users/ is found and not at the beginning of the path
    If usersPos > 1 Then
        ' Remove computer name part, keep only /Users/ and after
        vbaPath = Mid(vbaPath, usersPos)
    End If
    
    ConvertMacPath = vbaPath
End Function
