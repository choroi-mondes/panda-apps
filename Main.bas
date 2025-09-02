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

    ' �ς񂾂��Ղ�i���̃u�b�N�j
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

    arr = Array("�S����", "���Ӑ於", "�@��", "�m��}", "�݌v�S����", "�e�L�X�g252", "�e�L�X�g253", "���l�Q", "���ϓ��e", "�ؐ�", "�\���d�l")

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
            If Cells(1, j) = "�S����" Then
                   If j <> 1 Then
                       Columns(j).Cut
                       Columns(1).Insert
                   End If
            ElseIf Cells(1, j) = "���Ӑ於" Then
                   If j <> 2 Then
                        Columns(j).Cut
                       Columns(2).Insert
                   End If
            ElseIf Cells(1, j) = "�@��" Then
                   If j <> 3 Then
                       Columns(j).Cut
                      Columns(3).Insert
                   End If
            ElseIf Cells(1, j) = "�m��}" Then
                   If j <> 4 Then
                       Columns(j).Cut
                     Columns(4).Insert
                   End If
            ElseIf Cells(1, j) = "�݌v�S����" Then
                   If j <> 5 Then
                      Columns(j).Cut
                      Columns(5).Insert
                   End If
            ElseIf Cells(1, j) = "�e�L�X�g252" Then
                   If j <> 6 Then
                      Columns(j).Cut
                       Columns(6).Insert
                   End If
            ElseIf Cells(1, j) = "�e�L�X�g253" Then
                   If j <> 7 Then
                      Columns(j).Cut
                       Columns(7).Insert
                   End If
            ElseIf Cells(1, j) = "���l�Q" Then
                   If j <> 8 Then
                    Columns(j).Cut
                    Columns(8).Insert
                   End If
            ElseIf Cells(1, j) = "���ϓ��e" Then
                   If j <> 9 Then
                    Columns(j).Cut
                       Columns(9).Insert
                   End If
            ElseIf Cells(1, j) = "�ؐ�" Then
                   If j <> 10 Then
                    Columns(j).Cut
                    Columns(10).Insert
                   End If
            ElseIf Cells(1, j) = "�\���d�l" Then
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

    ' �ς񂾂��Ղ�i���̃u�b�N�j
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

    ' �ς񂾂��Ղ�i���̃u�b�N�j
    Set pandaBook = ThisWorkbook
    Set pandaSheet = pandaBook.Sheets(1)
    Set ws = pandaBook.Sheets(2)
    
    Call SortSel(ws, 11)
End Sub

Sub SortSel(ws As Worksheet, sortKey As Long)

    ' �\�[�g�t�B�[���h�̃N���A
    Dim pandaBook As Workbook
    Dim pandaSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' �ς񂾂��Ղ�i���̃u�b�N�j
    Set pandaBook = ThisWorkbook
    Set pandaSheet = pandaBook.Sheets(1)
    
    ws.Activate
    ws.Sort.SortFields.Clear
    
    ' �f�[�^�̍ŏI�s���擾
    lastRow = ActiveSheet.Cells(Rows.Count, sortKey).End(xlUp).Row

    ' �󗓂̃Z���� "0000/00/00" �ɒu��������
    For i = 2 To lastRow
'        If Trim(ActiveSheet.Cells(i, sortKey).Value) = "" And Trim(ActiveSheet.Cells(i, sortKey + 1).Value) = "" Then
        If Trim(ActiveSheet.Cells(i, sortKey).Value) = "" Then
            ActiveSheet.Cells(i, sortKey).Value = "45/12/31"
'            ActiveSheet.Cells(i, sortKey + 1).Value = "45/12/31"
        End If
    Next i
    
    
    ' �\�[�g�t�B�[���h�̒ǉ�
    ws.Sort.SortFields.Add _
        Key:=ActiveSheet.Range(ActiveSheet.Cells(2, sortKey), ActiveSheet.Cells(lastRow, sortKey)), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal

    ' �\�[�g�̎��s
    With ws.Sort
        .SetRange ws.UsedRange
        .Header = xlYes ' �w�b�_�[���܂ޏꍇ
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' �\�[�g��Ɍ��̏�Ԃɖ߂�
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
    
    ' �ς񂾂��Ղ�i���̃u�b�N�j
    Set pandaBook = ThisWorkbook
    Set pandaSheet = pandaBook.Sheets("�ς񂾂ς�")

    ' ���݂�Excel�t�@�C���Ɠ����t�H���_�̃p�X���擾
    folderPath = ThisWorkbook.Path

    ' �t�@�C���I���_�C�A���O�i�����������t�H���_��ݒ�j
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .Title = "���f�[�^���J��"
        .InitialFileName = folderPath & "\"
        .Filters.Clear
        .Filters.Add "Excel�t�@�C��", "*.xlsx; *.xlsm; *.xls"
        .AllowMultiSelect = False

        ' OK�������ꂽ��t�@�C���p�X���擾
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub ' �L�����Z�����ꂽ�ꍇ�͏I��
        End If
    End With

    ' �I�������t�@�C�����J��
    Set targetBook = Workbooks.Open(filePath)
    Set targetSheet = targetBook.Sheets(1)
    
    ' �V�[�g���� yyyymmdd_�A�� �ɂ���
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

    ' �V�[�g�R�s�[�i�ς񂾂ς񂿂̉E���ɒǉ��j
    targetSheet.Copy After:=pandaBook.Sheets(1)
'    Set tempSheet = pandaBook.Sheets(pandaBook.Sheets.Count)
    Set tempSheet = pandaBook.Sheets(2)
    tempSheet.Name = newSheetName

    ' ���u�b�N�����i�ۑ����Ȃ��j
    Application.DisplayAlerts = False
    targetBook.Close False
    Application.DisplayAlerts = True

    ' �ς񂾂ς񂿂ɖ߂�
    pandaSheet.Activate

End Sub

' **�V�[�g�������łɑ��݂��邩�`�F�b�N**
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


