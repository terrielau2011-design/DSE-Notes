Attribute VB_Name = "ExportModule"
'=============================================================
' DSE筆記系統 — 匯出模組
' 功能：將筆記匯出為 PDF 或 Word 文件
' 安裝：Excel → Alt+F11 → 檔案 → 匯入檔案 → 選擇此檔案
'=============================================================
Option Explicit

' ===== 主入口：匯出 PDF =====
Sub ExportToPDF()
    Dim wsSettings As Worksheet
    Dim wsTemp As Worksheet
    Dim subjectFilter As String
    Dim dataFilter As String
    Dim savePath As String
    Dim outputRow As Long
    
    On Error GoTo ErrorHandler
    
    ' 讀取設定
    Set wsSettings = ThisWorkbook.Sheets("匯出筆記")
    subjectFilter = wsSettings.Range("B4").Value
    dataFilter = wsSettings.Range("B6").Value
    
    If subjectFilter = "" Or InStr(subjectFilter, "下拉") > 0 Then subjectFilter = "全部"
    If dataFilter = "" Or InStr(dataFilter, "可選") > 0 Then dataFilter = "全部"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 建立臨時工作表
    Set wsTemp = ThisWorkbook.Sheets.Add
    wsTemp.Name = "PDF_Export_" & Format(Now, "hhmmss")
    
    outputRow = 1
    
    ' === 標題 ===
    With wsTemp.Cells(outputRow, 1)
        .Value = "DSE 筆記匯出"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(27, 58, 92)
    End With
    outputRow = outputRow + 1
    With wsTemp.Cells(outputRow, 1)
        .Value = "匯出日期：" & Format(Now, "yyyy-mm-dd hh:mm") & "    科目：" & subjectFilter & "    篩選：" & dataFilter
        .Font.Size = 10
        .Font.Color = RGB(150, 150, 150)
    End With
    outputRow = outputRow + 2
    
    ' === 筆記类科目 ===
    Dim noteSheets As Variant, noteNames As Variant
    noteSheets = Array("筆記_中史", "筆記_世史", "筆記_C&SD")
    noteNames = Array("中史", "世史", "C&SD")
    
    Dim i As Long
    For i = 0 To 2
        If subjectFilter = "全部" Or subjectFilter = noteNames(i) Then
            Call WriteNotesToSheet(ThisWorkbook.Sheets(noteSheets(i)), noteNames(i), wsTemp, outputRow, dataFilter)
        End If
    Next i
    
    ' === 操練类科目 ===
    Dim pracSheets As Variant, pracNames As Variant
    pracSheets = Array("操練_英文", "操練_數學", "操練_中文")
    pracNames = Array("英文", "數學", "中文")
    
    For i = 0 To 2
        If subjectFilter = "全部" Or subjectFilter = pracNames(i) Then
            Call WritePracticeToSheet(ThisWorkbook.Sheets(pracSheets(i)), pracNames(i), wsTemp, outputRow)
        End If
    Next i
    
    ' === 格式化欄寬 ===
    wsTemp.Columns("A").ColumnWidth = 22
    wsTemp.Columns("B").ColumnWidth = 65
    wsTemp.Columns("C").ColumnWidth = 35
    wsTemp.Columns("D").ColumnWidth = 25
    
    ' === 匯出 PDF ===
    savePath = ThisWorkbook.Path & Application.PathSeparator & "DSE筆記_" & subjectFilter & "_" & Format(Now, "yyyymmdd_hhmm") & ".pdf"
    
    wsTemp.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    ' 清除臨時工作表
    wsTemp.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "PDF 已匯出！" & vbNewLine & vbNewLine & "儲存位置：" & savePath, vbInformation, "匯出完成"
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "匯出時發生錯誤：" & vbNewLine & Err.Description, vbCritical, "錯誤"
End Sub

' ===== 主入口：匯出 Word =====
Sub ExportToWord()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim wsSettings As Worksheet
    Dim subjectFilter As String
    Dim dataFilter As String
    Dim savePath As String
    
    On Error GoTo WordError
    
    ' 檢查 Word 是否可用
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Add
    
    ' 讀取設定
    Set wsSettings = ThisWorkbook.Sheets("匯出筆記")
    subjectFilter = wsSettings.Range("B4").Value
    dataFilter = wsSettings.Range("B6").Value
    
    If subjectFilter = "" Or InStr(subjectFilter, "下拉") > 0 Then subjectFilter = "全部"
    If dataFilter = "" Or InStr(dataFilter, "可選") > 0 Then dataFilter = "全部"
    
    ' === 標題 ===
    Dim rng As Object
    Set rng = wdDoc.Range
    rng.Text = "DSE 筆記匯出"
    rng.Font.Size = 22
    rng.Font.Bold = True
    rng.Font.Color = RGB(27, 58, 92)
    rng.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
    rng.InsertParagraphAfter
    rng.Collapse 0 ' wdCollapseEnd
    
    rng.Text = "匯出日期：" & Format(Now, "yyyy-mm-dd hh:mm") & "    科目：" & subjectFilter & "    篩選：" & dataFilter
    rng.Font.Size = 10
    rng.Font.Color = RGB(150, 150, 150)
    rng.ParagraphFormat.Alignment = 1
    rng.InsertParagraphAfter
    rng.InsertParagraphAfter
    rng.Collapse 0
    
    ' === 筆記类科目 ===
    Dim noteSheets As Variant, noteNames As Variant
    noteSheets = Array("筆記_中史", "筆記_世史", "筆記_C&SD")
    noteNames = Array("中史", "世史", "C&SD")
    
    Dim i As Long
    For i = 0 To 2
        If subjectFilter = "全部" Or subjectFilter = noteNames(i) Then
            Call WriteNotesToWord(ThisWorkbook.Sheets(noteSheets(i)), noteNames(i), wdDoc, rng, dataFilter)
        End If
    Next i
    
    ' === 操練类科目 ===
    Dim pracSheets As Variant, pracNames As Variant
    pracSheets = Array("操練_英文", "操練_數學", "操練_中文")
    pracNames = Array("英文", "數學", "中文")
    
    For i = 0 To 2
        If subjectFilter = "全部" Or subjectFilter = pracNames(i) Then
            Call WritePracticeToWord(ThisWorkbook.Sheets(pracSheets(i)), pracNames(i), wdDoc, rng)
        End If
    Next i
    
    ' 儲存
    savePath = ThisWorkbook.Path & Application.PathSeparator & "DSE筆記_" & subjectFilter & "_" & Format(Now, "yyyymmdd_hhmm") & ".docx"
    wdDoc.SaveAs2 savePath, 16 ' wdFormatDocumentDefault
    
    MsgBox "Word 檔案已匯出！" & vbNewLine & vbNewLine & "儲存位置：" & savePath, vbInformation, "匯出完成"
    Exit Sub
    
WordError:
    If Err.Number = 429 Then
        MsgBox "無法啟動 Microsoft Word。" & vbNewLine & vbNewLine & "請確認已安裝 Microsoft Word。" & vbNewLine & "如未安裝 Word，請使用「匯出 PDF」功能。", vbCritical, "Word 未安裝"
    Else
        MsgBox "匯出時發生錯誤：" & vbNewLine & "錯誤代碼 " & Err.Number & "：" & Err.Description, vbCritical, "錯誤"
    End If
    On Error Resume Next
    If Not wdApp Is Nothing Then
        wdApp.Quit False
    End If
End Sub

' ===== 將筆記寫入臨時工作表（PDF用）=====
Private Sub WriteNotesToSheet(wsSource As Worksheet, subjectName As String, wsDest As Worksheet, ByRef outputRow As Long, dataFilter As String)
    ' 科目標題
    With wsDest.Cells(outputRow, 1)
        .Value = subjectName & " 筆記"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(139, 69, 19)
    End With
    wsDest.Cells(outputRow, 1).Resize(1, 4).Merge
    outputRow = outputRow + 1
    
    Dim r As Long
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    Dim noteIndex As Integer: noteIndex = 1
    
    For r = 2 To lastRow
        If wsSource.Cells(r, 1).Value = "" Then GoTo NextNoteRow
        
        ' 套用篩選
        If dataFilter = "只匯出DSE曾出現" And wsSource.Cells(r, 5).Value <> "是" Then GoTo NextNoteRow
        If dataFilter = "只匯出熟練度1-2" Then
            If IsNumeric(wsSource.Cells(r, 10).Value) And wsSource.Cells(r, 10).Value > 2 Then GoTo NextNoteRow
        End If
        If dataFilter = "只匯出熟練度1-3" Then
            If IsNumeric(wsSource.Cells(r, 10).Value) And wsSource.Cells(r, 10).Value > 3 Then GoTo NextNoteRow
        End If
        
        ' 課題標題
        With wsDest.Cells(outputRow, 1)
            .Value = noteIndex & ". " & wsSource.Cells(r, 1).Value & " — " & wsSource.Cells(r, 2).Value
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = RGB(27, 58, 92)
        End With
        wsDest.Cells(outputRow, 1).Resize(1, 4).Merge
        outputRow = outputRow + 1
        
        ' 筆記重點
        If wsSource.Cells(r, 3).Value <> "" Then
            wsDest.Cells(outputRow, 1).Value = "筆記重點："
            wsDest.Cells(outputRow, 1).Font.Bold = True
            wsDest.Cells(outputRow, 1).Font.Color = RGB(139, 69, 19)
            wsDest.Cells(outputRow, 2).Value = wsSource.Cells(r, 3).Value
            wsDest.Cells(outputRow, 2).WrapText = True
            outputRow = outputRow + 1
        End If
        
        ' 關鍵詞
        If wsSource.Cells(r, 4).Value <> "" Then
            wsDest.Cells(outputRow, 1).Value = "關鍵詞："
            wsDest.Cells(outputRow, 1).Font.Bold = True
            wsDest.Cells(outputRow, 2).Value = wsSource.Cells(r, 4).Value
            wsDest.Cells(outputRow, 2).Font.Color = RGB(100, 100, 100)
            outputRow = outputRow + 1
        End If
        
        ' DSE出題
        If wsSource.Cells(r, 5).Value = "是" Then
            wsDest.Cells(outputRow, 1).Value = "DSE出題："
            wsDest.Cells(outputRow, 1).Font.Bold = True
            wsDest.Cells(outputRow, 1).Font.Color = RGB(192, 57, 43)
            Dim dseInfo As String
            dseInfo = "已出現"
            If wsSource.Cells(r, 6).Value <> "" Then dseInfo = dseInfo & "  |  題型：" & wsSource.Cells(r, 6).Value
            If wsSource.Cells(r, 7).Value <> "" Then dseInfo = dseInfo & "  |  年份：" & wsSource.Cells(r, 7).Value
            wsDest.Cells(outputRow, 2).Value = dseInfo
            wsDest.Cells(outputRow, 2).Font.Color = RGB(192, 57, 43)
            outputRow = outputRow + 1
        End If
        
        ' 自測題目
        If wsSource.Cells(r, 8).Value <> "" Then
            wsDest.Cells(outputRow, 1).Value = "自測："
            wsDest.Cells(outputRow, 1).Font.Bold = True
            wsDest.Cells(outputRow, 2).Value = wsSource.Cells(r, 8).Value
            wsDest.Cells(outputRow, 2).Font.Color = RGB(41, 128, 185)
            outputRow = outputRow + 1
        End If
        
        ' 自測答案
        If wsSource.Cells(r, 9).Value <> "" Then
            wsDest.Cells(outputRow, 1).Value = "答案："
            wsDest.Cells(outputRow, 1).Font.Bold = True
            wsDest.Cells(outputRow, 2).Value = wsSource.Cells(r, 9).Value
            wsDest.Cells(outputRow, 2).Font.Color = RGB(39, 174, 96)
            outputRow = outputRow + 1
        End If
        
        ' 熟練度
        If wsSource.Cells(r, 10).Value <> "" Then
            wsDest.Cells(outputRow, 1).Value = "熟練度：" & wsSource.Cells(r, 10).Value & " / 5"
            outputRow = outputRow + 1
        End If
        
        outputRow = outputRow + 1
        noteIndex = noteIndex + 1
NextNoteRow:
    Next r
    outputRow = outputRow + 1
End Sub

' ===== 將操練記錄寫入臨時工作表（PDF用）=====
Private Sub WritePracticeToSheet(wsSource As Worksheet, subjectName As String, wsDest As Worksheet, ByRef outputRow As Long)
    With wsDest.Cells(outputRow, 1)
        .Value = subjectName & " 操練記錄"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(46, 80, 144)
    End With
    wsDest.Cells(outputRow, 1).Resize(1, 4).Merge
    outputRow = outputRow + 1
    
    ' 表頭
    Dim headers As Variant
    headers = Array("日期", "操練類型 / 題號 / 得分", "錯誤分析")
    Dim j As Long
    For j = 0 To 2
        wsDest.Cells(outputRow, j + 1).Value = headers(j)
        wsDest.Cells(outputRow, j + 1).Font.Bold = True
        wsDest.Cells(outputRow, j + 1).Interior.Color = RGB(27, 58, 92)
        wsDest.Cells(outputRow, j + 1).Font.Color = RGB(255, 255, 255)
    Next j
    outputRow = outputRow + 1
    
    Dim r As Long
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lastRow
        If wsSource.Cells(r, 1).Value = "" Then GoTo NextPracRow
        
        wsDest.Cells(outputRow, 1).Value = wsSource.Cells(r, 1).Value
        wsDest.Cells(outputRow, 2).Value = wsSource.Cells(r, 2).Value & " | " & wsSource.Cells(r, 6).Value & " | " & wsSource.Cells(r, 7).Value
        
        If wsSource.Cells(r, 8).Value <> "" Then
            wsDest.Cells(outputRow, 3).Value = wsSource.Cells(r, 8).Value & "：" & wsSource.Cells(r, 9).Value
            wsDest.Cells(outputRow, 3).Font.Color = RGB(192, 57, 43)
            If wsSource.Cells(r, 10).Value <> "" Then
                wsDest.Cells(outputRow, 4).Value = "改善：" & wsSource.Cells(r, 10).Value
                wsDest.Cells(outputRow, 4).Font.Color = RGB(39, 174, 96)
            End If
        End If
        
        outputRow = outputRow + 1
NextPracRow:
    Next r
    outputRow = outputRow + 1
End Sub

' ===== 將筆記寫入 Word 文件 =====
Private Sub WriteNotesToWord(wsSource As Worksheet, subjectName As String, wdDoc As Object, rng As Object, dataFilter As String)
    ' 科目標題
    rng.Text = subjectName & " 筆記"
    rng.Font.Size = 16
    rng.Font.Bold = True
    rng.Font.Color = RGB(139, 69, 19)
    rng.InsertParagraphAfter
    rng.Collapse 0
    
    Dim r As Long
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    Dim noteIndex As Integer: noteIndex = 1
    
    For r = 2 To lastRow
        If wsSource.Cells(r, 1).Value = "" Then GoTo NextNoteWord
        
        ' 套用篩選
        If dataFilter = "只匯出DSE曾出現" And wsSource.Cells(r, 5).Value <> "是" Then GoTo NextNoteWord
        If dataFilter = "只匯出熟練度1-2" Then
            If IsNumeric(wsSource.Cells(r, 10).Value) And wsSource.Cells(r, 10).Value > 2 Then GoTo NextNoteWord
        End If
        If dataFilter = "只匯出熟練度1-3" Then
            If IsNumeric(wsSource.Cells(r, 10).Value) And wsSource.Cells(r, 10).Value > 3 Then GoTo NextNoteWord
        End If
        
        ' 課題標題
        rng.Text = noteIndex & ". " & wsSource.Cells(r, 1).Value & " — " & wsSource.Cells(r, 2).Value
        rng.Font.Size = 13
        rng.Font.Bold = True
        rng.Font.Color = RGB(27, 58, 92)
        rng.InsertParagraphAfter
        rng.Collapse 0
        
        ' 筆記重點
        If wsSource.Cells(r, 3).Value <> "" Then
            rng.Text = "筆記重點：" & wsSource.Cells(r, 3).Value
            rng.Font.Size = 11
            rng.Font.Bold = False
            rng.Font.Color = RGB(50, 50, 50)
            rng.InsertParagraphAfter
            rng.Collapse 0
        End If
        
        ' 關鍵詞
        If wsSource.Cells(r, 4).Value <> "" Then
            rng.Text = "關鍵詞：" & wsSource.Cells(r, 4).Value
            rng.Font.Size = 10
            rng.Font.Color = RGB(100, 100, 100)
            rng.InsertParagraphAfter
            rng.Collapse 0
        End If
        
        ' DSE出題
        If wsSource.Cells(r, 5).Value = "是" Then
            Dim dseInfo As String
            dseInfo = "DSE出題：已出現"
            If wsSource.Cells(r, 6).Value <> "" Then dseInfo = dseInfo & "  題型：" & wsSource.Cells(r, 6).Value
            If wsSource.Cells(r, 7).Value <> "" Then dseInfo = dseInfo & "  年份：" & wsSource.Cells(r, 7).Value
            rng.Text = dseInfo
            rng.Font.Size = 10
            rng.Font.Bold = True
            rng.Font.Color = RGB(192, 57, 43)
            rng.InsertParagraphAfter
            rng.Collapse 0
        End If
        
        ' 自測
        If wsSource.Cells(r, 8).Value <> "" Then
            rng.Text = "自測：" & wsSource.Cells(r, 8).Value
            rng.Font.Size = 10
            rng.Font.Color = RGB(41, 128, 185)
            rng.InsertParagraphAfter
            rng.Collapse 0
        End If
        If wsSource.Cells(r, 9).Value <> "" Then
            rng.Text = "答案：" & wsSource.Cells(r, 9).Value
            rng.Font.Size = 10
            rng.Font.Color = RGB(39, 174, 96)
            rng.InsertParagraphAfter
            rng.Collapse 0
        End If
        
        ' 熟練度
        If wsSource.Cells(r, 10).Value <> "" Then
            rng.Text = "熟練度：" & wsSource.Cells(r, 10).Value & " / 5"
            rng.Font.Size = 10
            rng.Font.Color = RGB(100, 100, 100)
            rng.InsertParagraphAfter
            rng.Collapse 0
        End If
        
        rng.Text = ""
        rng.InsertParagraphAfter
        rng.Collapse 0
        noteIndex = noteIndex + 1
NextNoteWord:
    Next r
    
    rng.Text = ""
    rng.InsertParagraphAfter
    rng.Collapse 0
End Sub

' ===== 將操練記錄寫入 Word 文件 =====
Private Sub WritePracticeToWord(wsSource As Worksheet, subjectName As String, wdDoc As Object, rng As Object)
    rng.Text = subjectName & " 操練記錄"
    rng.Font.Size = 16
    rng.Font.Bold = True
    rng.Font.Color = RGB(46, 80, 144)
    rng.InsertParagraphAfter
    rng.Collapse 0
    
    Dim r As Long
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lastRow
        If wsSource.Cells(r, 1).Value = "" Then GoTo NextPracWord
        
        Dim lineText As String
        lineText = wsSource.Cells(r, 1).Value & " | " & wsSource.Cells(r, 2).Value & " | " & wsSource.Cells(r, 6).Value & " | " & wsSource.Cells(r, 7).Value
        
        If wsSource.Cells(r, 8).Value <> "" Then
            lineText = lineText & vbNewLine & "  錯誤：" & wsSource.Cells(r, 8).Value & " — " & wsSource.Cells(r, 9).Value
            If wsSource.Cells(r, 10).Value <> "" Then
                lineText = lineText & vbNewLine & "  改善：" & wsSource.Cells(r, 10).Value
            End If
        End If
        
        rng.Text = lineText
        rng.Font.Size = 10
        rng.Font.Color = RGB(50, 50, 50)
        rng.InsertParagraphAfter
        rng.Collapse 0
        
NextPracWord:
    Next r
    
    rng.Text = ""
    rng.InsertParagraphAfter
    rng.Collapse 0
End Sub
