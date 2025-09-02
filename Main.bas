Attribute VB_Name = "Main"
Option Explicit

Sub DoDoDo()

    Call PandaSheetCopy
    Call FitTable
    Call SortArc

End Sub

Sub DoneDone()
    Call SortDate
End Sub

Sub FitTable()

    Dim i, j As Integer
    Dim maxRow, maxCol As Integer
    Dim nowRow, nowCol As Integer
    
    Dim sortCol As Integer
    
    Dim maxRowHeight As Double
    Dim maxColWidth As Double
    
    Dim arr() As Variant
    Dim result As Variant
    
    Dim pandaBook As Workbook
    Dim pandaSheet As Worksheet
    Dim tempSheet As Worksheet

    ' ぱんだあぷり（このブック）
    Set pandaBook = ThisWorkbook
    Set pandaSheet = pandaBook.Sheets(1)
    Set tempSheet = pandaBook.Sheets(2)

    i = 0
    j = 0
    maxRow = 1
    maxCol = 1
    nowRow = 1
    nowCol = 1
    sortCol = 1
    maxRowHeight = 2
    maxColWidth = 25

    arr = Array("担当者", "得意先名", "邸名", "確定図", "設計担当者", "テキスト252", "テキスト253", "備考２", "見積内容", "坪数", "構造仕様")

    tempSheet.Activate
    
    maxRow = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    maxCol = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    
    Cells(1, maxCol).Select
    
    For j = maxCol To 1 Step -1
        result = Application.Match(ActiveSheet.Cells(1, j), arr, 0)
        If IsError(result) Then
            ActiveSheet.Columns(j).Delete
        End If
    Next
    
    maxRow = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    maxCol = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    Range(Cells(1, 1), Cells(maxRow, maxCol)).Select
    Range(Cells(1, 1), Cells(maxRow, maxCol)).RowHeight = 20
    Range(Cells(1, 1), Cells(maxRow, maxCol)).Columns.AutoFit
    Range(Cells(1, 1), Cells(maxRow, maxCol)).Rows.AutoFit
    
    ActiveSheet.Cells(1, 1).Select

    
    For i = 1 To maxCol
        For j = i To maxCol
            If Cells(1, j) = "担当者" Then
                   If j <> 1 Then
                       Columns(j).Cut
                       Columns(1).Insert
                   End If
            ElseIf Cells(1, j) = "得意先名" Then
                   If j <> 2 Then
                        Columns(j).Cut
                       Columns(2).Insert
                   End If
            ElseIf Cells(1, j) = "邸名" Then
                   If j <> 3 Then
                       Columns(j).Cut
                      Columns(3).Insert
                   End If
            ElseIf Cells(1, j) = "確定図" Then
                   If j <> 4 Then
                       Columns(j).Cut
                     Columns(4).Insert
                   End If
            ElseIf Cells(1, j) = "設計担当者" Then
                   If j <> 5 Then
                      Columns(j).Cut
                      Columns(5).Insert
                   End If
            ElseIf Cells(1, j) = "テキスト252" Then
                   If j <> 6 Then
                      Columns(j).Cut
                       Columns(6).Insert
                   End If
            ElseIf Cells(1, j) = "テキスト253" Then
                   If j <> 7 Then
                      Columns(j).Cut
                       Columns(7).Insert
                   End If
            ElseIf Cells(1, j) = "備考２" Then
                   If j <> 8 Then
                    Columns(j).Cut
                    Columns(8).Insert
                   End If
            ElseIf Cells(1, j) = "見積内容" Then
                   If j <> 9 Then
                    Columns(j).Cut
                       Columns(9).Insert
                   End If
            ElseIf Cells(1, j) = "坪数" Then
                   If j <> 10 Then
                    Columns(j).Cut
                    Columns(10).Insert
                   End If
            ElseIf Cells(1, j) = "構造仕様" Then
                    sortCol = j
                   If j <> 11 Then
                    Columns(j).Cut
                    Columns(11).Insert
                   End If
            End If
        Next
    Next
    
'    Call SortSel(tempSheet, 11)
    
    tempSheet.Cells(1, 1).Activate
    tempSheet.Cells(1, 1).Select
    pandaSheet.Activate
    pandaSheet.Cells(1, 1).Select

End Sub

Sub SortDate()
    Dim pandaBook As Workbook
    Dim pandaSheet As Worksheet
    Dim ws As Worksheet

    ' ぱんだあぷり（このブック）
    Set pandaBook = ThisWorkbook
    Set pandaSheet = pandaBook.Sheets(1)
    Set ws = pandaBook.Sheets(2)
    
    Call SortSel(ws, 7)
    Call SortSel(ws, 6)
End Sub

Sub SortArc()
    Dim pandaBook As Workbook
    Dim pandaSheet As Worksheet
    Dim ws As Worksheet

    ' ぱんだあぷり（このブック）
    Set pandaBook = ThisWorkbook
    Set pandaSheet = pandaBook.Sheets(1)
    Set ws = pandaBook.Sheets(2)
    
    Call SortSel(ws, 11)
End Sub

Sub SortSel(ws As Worksheet, sortKey As Long)

    ' ソートフィールドのクリア
    Dim pandaBook As Workbook
    Dim pandaSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' ぱんだあぷり（このブック）
    Set pandaBook = ThisWorkbook
    Set pandaSheet = pandaBook.Sheets(1)
    
    ws.Activate
    ws.Sort.SortFields.Clear
    
    ' データの最終行を取得
    lastRow = ActiveSheet.Cells(Rows.Count, sortKey).End(xlUp).Row

    ' 空欄のセルを "0000/00/00" に置き換える
    For i = 2 To lastRow
'        If Trim(ActiveSheet.Cells(i, sortKey).Value) = "" And Trim(ActiveSheet.Cells(i, sortKey + 1).Value) = "" Then
        If Trim(ActiveSheet.Cells(i, sortKey).Value) = "" Then
            ActiveSheet.Cells(i, sortKey).Value = "45/12/31"
'            ActiveSheet.Cells(i, sortKey + 1).Value = "45/12/31"
        End If
    Next i
    
    
    ' ソートフィールドの追加
    ws.Sort.SortFields.Add _
        Key:=ActiveSheet.Range(ActiveSheet.Cells(2, sortKey), ActiveSheet.Cells(lastRow, sortKey)), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    ' ソートの実行
    With ws.Sort
        .SetRange ws.UsedRange
        .Header = xlYes ' ヘッダーを含む場合
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' ソート後に元の状態に戻す
    For i = 2 To lastRow
        If ActiveSheet.Cells(i, sortKey).Value = "45/12/31" Then
            ActiveSheet.Cells(i, sortKey).Value = ""
'            ActiveSheet.Cells(i, sortKey + 1).Value = ""
        End If
    Next i
    
    ws.Cells(1, 1).Select
    pandaSheet.Activate
    pandaSheet.Cells(1, 1).Select
End Sub

Sub PandaSheetCopy()
    Dim pandaBook As Workbook
    Dim pandaSheet As Worksheet
    Dim targetBook As Workbook
    Dim targetSheet As Worksheet
    Dim tempSheet As Worksheet
    Dim baseName As String
    Dim newSheetName As String
    Dim sheetNum As Integer
    Dim filePath As String
    Dim folderPath As String
    Dim fileDialog As fileDialog
    
    ' ぱんだあぷり（このブック）
    Set pandaBook = ThisWorkbook
    Set pandaSheet = pandaBook.Sheets("ぱんだぱんち")

    ' 現在のExcelファイルと同じフォルダのパスを取得
    folderPath = ThisWorkbook.Path

    ' ファイル選択ダイアログ（正しく初期フォルダを設定）
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .Title = "元データを開く"
        .InitialFileName = folderPath & "\"
        .Filters.Clear
        .Filters.Add "Excelファイル", "*.xlsx; *.xlsm; *.xls"
        .AllowMultiSelect = False

        ' OKが押されたらファイルパスを取得
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub ' キャンセルされた場合は終了
        End If
    End With

    ' 選択したファイルを開く
    Set targetBook = Workbooks.Open(filePath)
    Set targetSheet = targetBook.Sheets(1)
    
    ' シート名を yyyymmdd_連番 にする
    baseName = Format(Date, "yyyymmdd")
    sheetNum = 1
    Do While SheetExists(baseName & "_" & sheetNum, pandaBook)
        sheetNum = sheetNum + 1
    Loop
'    If sheetNum = 1 Then
'        newSheetName = baseName
'    Else
        newSheetName = baseName & "_" & sheetNum
'    End If

    ' シートコピー（ぱんだぱんちの右側に追加）
    targetSheet.Copy After:=pandaBook.Sheets(1)
'    Set tempSheet = pandaBook.Sheets(pandaBook.Sheets.Count)
    Set tempSheet = pandaBook.Sheets(2)
    tempSheet.Name = newSheetName

    ' 元ブックを閉じる（保存しない）
    Application.DisplayAlerts = False
    targetBook.Close False
    Application.DisplayAlerts = True

    ' ぱんだぱんちに戻る
    pandaSheet.Activate

End Sub

' **シート名がすでに存在するかチェック**
Function SheetExists(sheetName As String, book As Workbook) As Boolean
    Dim ws As Worksheet
    For Each ws In book.Worksheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
    SheetExists = False
End Function


