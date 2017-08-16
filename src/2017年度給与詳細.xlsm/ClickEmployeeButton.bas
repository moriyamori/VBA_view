Attribute VB_Name = "ClickEmployeeButton"
    Dim topRow As Integer
    Dim buttomRow As Integer
    Dim employeeSalaryItemRange As Range
    Dim monthlyTotalFomula As String
    Dim nameItemeAddress As String
    Dim positionOfTransferAmount As String
    Dim bindingResult As String
    Dim AdjustmentNumberOfPeople As String
       
'社員追加ボタン押下時の処理
Public Sub employeeAddButton()

    '社員給与詳細アクティブ
    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate

    '社員の給与欄追加する最上行を設定
    topRow = Range("C:C").Find(what:="基本給", SearchDirection:=xlPrevious).Row + ROW_TO_THE_TOP_ROW

    '社員の給与欄追加する最下行を設定
    buttomRow = topRow + 18

    '社員給与欄範囲
    Set employeeSalaryItemRange = Range(Cells(topRow, LEFTMOST_COLUMN), Cells(buttomRow, RIGHTMOST_COLUMN))

    '一人分の給与欄セルをインサート
    employeeSalaryItemRange.Insert (xlShiftDown)

    'データワークシートから給与欄テンプレを取得
    ThisWorkbook.Worksheets("EmployeeData").Range(EMPLOYEE_SALARY_ITEMS).Copy

    '給与欄テンプレを指定の位置に貼り付け
    Cells(topRow, 1).PasteSpecial (xlPasteAll)

    'コピー点線解除
    Application.CutCopyMode = False

    '***「振込額一覧」シートに追加した給与欄の振込額総計を追加する***

    '「■2017年度　社員給与詳細」シートから名前欄の位置を取得
    nameItemeAddress = Cells(topRow, 1).Address(ReferenceStyle:=xlR1C1)

    '「■2017年度　社員給与詳細」シートから振込額のある位置を取得
    positionOfTransferAmount = Cells(buttomRow, 4).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    Worksheets("■振込額一覧").Activate

    '振込額の枠追加
    With Range("A:A").Find(what:="社員")
        .Offset(0, 1).End(xlDown).End(xlDown).Offset(1).Select
        Rows(ActiveCell.Row).Select
        Selection.Insert (xlShiftDown)

        '「■振込額一覧」シートに名前欄位置の数式を追加
        bindingResult = MEIN_SHEET_NAME2 + nameItemeAddress
        bindingResult = "=IF(" & bindingResult & "="""",""""," & bindingResult & ")"
        .Offset(0, 1).End(xlDown).End(xlDown).Offset(1) = bindingResult

        '「■振込額一覧」シートの一月分
        bindingResult = MEIN_SHEET_NAME + positionOfTransferAmount
        .Offset(0, 1).End(xlDown).End(xlDown).Offset(0, 1) = bindingResult

        'オートフィルで1年間分表示
        .Offset(0, 1).End(xlDown).End(xlDown).Offset(0, 1).AutoFill Destination:= _
        .Offset(0, 1).End(xlDown).End(xlDown).Offset(0, 1).Resize(1, 12)
    End With

    '年度支給額計を追加
    Cells.Find(what:="年度支給額計").End(xlDown).End(xlDown).AutoFill Destination:=Cells.Find(what:="年度支給額計").End(xlDown).End(xlDown).Resize(2)

    'ボーナス込欄設定---
    
    Dim positionLncledingBonus As String

    positionLncledingBonus = ThisWorkbook.Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Cells(buttomRow, 17).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    bindingResult = MEIN_SHEET_NAME + positionLncledingBonus
    ThisWorkbook.Worksheets("■振込額一覧").Cells.Find(what:="ボーナス込み").End(xlDown).End(xlDown).Offset(1) = bindingResult

    '社員月合計設定---

    ' 社員月次計の数式を設定
    monthlyTotalFomula = Range(Range("C8"), Range("C8").End(xlDown)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Range("A:A").Find(what:="社員月次計").Offset(0, 2).Formula = "=SUM(" & monthlyTotalFomula & ")"
    
    '一月の精算人数分の範囲設定
    monthlyTotalFomula2 = Range(Range("C7"), Range("C7").End(xlDown)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    '一月分の精算人数
    Range("A:A").Find(what:="■社員").Offset(0, 2).Formula = "=COUNTIF(" & monthlyTotalFomula2 & " ,""<>0"")"
           
    'オートフィルで1年間分の社員月次計を表示
    Range("A:A").Find(what:="社員月次計").Offset(0, 2).AutoFill Destination:=Range("A:A").Find(what:="社員月次計").Offset(0, 2).Resize(1, 12)

    Range("A:A").Find(what:="■社員").Offset(0, 2).AutoFill Destination:=Range("A:A").Find(what:="■社員").Offset(0, 2).Resize(1, 12)
    
    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate


End Sub

'社員削除ボタン押下時の処理
Public Sub employeeDeleteButton()

    '一人分の給与欄の最上行を設定---
    
    '「■2017年度　社員給与詳細」アクティブ
    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate

    '社員の給与欄を削除する最上行を設定
    topRow = Range("C:C").Find(what:="基本給", SearchDirection:=xlPrevious).Row

    '最上行が一人しかいない場合の行位置だった場合後続処理を停止する。
    If topRow = 0 Or topRow = 14 Then
        MsgBox "これ以上削除できません。", vbCritical
        Exit Sub
    End If

    '社員の給与欄を削除する最下行を設定
    buttomRow = topRow + 18

    '社員給与欄範囲設定
    Set employeeSalaryItemRange = Range(Cells(topRow, LEFTMOST_COLUMN), Cells(buttomRow, RIGHTMOST_COLUMN))

    '社員の給与欄枠を設定
    employeeSalaryItemRange.delete (xlShiftUp)

    '振込額削除
    Worksheets("■振込額一覧").Activate
    Cells.Find(what:="社員").Offset(0, 1).End(xlDown).End(xlDown).Select
    Rows(ActiveCell.Row).Select
    Selection.delete (xlUp)

    '***社員月次計の数式を格納***
    
    '社員月次計の人数が一人の場合、処理変更
    If Range("C9") = "" Then
        monthlyTotalFomula = Range("C8").Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Else
        monthlyTotalFomula = Range(Range("C8"), Range("C8").End(xlDown)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End If
    With Range("A:A").Find(what:="社員月次計").Offset(0, 2)
        .Formula = "=SUM(" & monthlyTotalFomula & ")"
        'オートフィルで1年間分の社員月次計を表示
        .AutoFill Destination:=Range("A:A").Find(what:="社員月次計").Offset(0, 2).Resize(1, 12)
    End With

    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate

End Sub
