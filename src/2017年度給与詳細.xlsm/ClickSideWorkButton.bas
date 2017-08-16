Attribute VB_Name = "ClickSideWorkButton"
Dim parttimeSalaryItem As Range
'バイトの追加ボタン押下時の処理
Public Sub partTimeAddButton()

'「■2017年度　社員給与詳細」シートをアクティブ
    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate

'***バイトの給与欄を追加***

    'バイトの給与欄を追加する最上行を設定
    topRow = Cells.Find(what:="月給与合計", SearchDirection:=xlPrevious).Row

    'バイトの給与欄を追加する最下行を設定
    buttomRow = topRow + 18

    'バイト給与欄範囲設定
    Set parttimeSalaryItem = Range(Cells(topRow, LEFTMOST_COLUMN), Cells(buttomRow, RIGHTMOST_COLUMN))

    '一人分の給与欄枠をインサート
    parttimeSalaryItem.Insert (xlShiftDown)

    '「PartData」ワークシートから給与欄テンプレを取得
    ThisWorkbook.Worksheets("PartData").Range(RANGE_OF_PART_TIME_SALARY_ITEMS).Copy

    'バイトの給与欄を指定の位置に貼り付け
    Cells(topRow, 1).PasteSpecial (xlPasteAll)

    'コピー用点線解除
    Application.CutCopyMode = False

'***「振込額一覧」シートに、追加した給与欄の振込額の総計を追加する。***

    Dim positionOfNameItem As Range
    Dim positionOfTransferAmountItem As Range

    '給与詳細シートから名前欄のアドレスを取得
    nameItemeAddress = Cells(topRow, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)

    '給与詳細シートからひと月分の振込額のあるアドレス取得
    positionOfTransferAmount = Cells(buttomRow, 4).Address(RowAbsolute:=False, ColumnAbsolute:=False)

    Worksheets("■振込額一覧").Activate

    '振込額の枠追加---

    'もし、アルバイト・パートが一人もいない場合or一人しかいない場合の範囲変更設定
    With Cells.Find(what:=PART_TIME)
        '0人
        If .Offset(1, 2) = "" Then
            Set positionOfNameItem = .Offset(1, 1)
            Set positionOfTransferAmountItem = .Offset(1, 2)
            '1人
        ElseIf Cells.Find(what:=PART_TIME).Offset(2, 2) = "" Then
            Set positionOfNameItem = .Offset(2, 1)
            Set positionOfTransferAmountItem = .Offset(2, 2)
            'それ以外
        Else
            Set positionOfNameItem = .Offset(0, 1).End(xlDown).End(xlDown).Offset(1)
            Set positionOfTransferAmountItem = .Offset(0, 1).End(xlDown).End(xlDown).Offset(1, 1)
        End If
    End With

    '「■振込額一覧」シートの名前欄にWS関数追加
    bindingResult = MEIN_SHEET_NAME2 + nameItemeAddress
    bindingResult = "=IF(" & bindingResult & "="""",""""," & bindingResult & ")"
    positionOfNameItem = bindingResult

    '「■振込額一覧」シートの１１月振込額欄にWS関数追加
    bindingResult = MEIN_SHEET_NAME + positionOfTransferAmount
    positionOfTransferAmountItem = bindingResult

    'オートフィルで1年間分表示
    positionOfTransferAmountItem.AutoFill Destination:=positionOfTransferAmountItem.Resize(1, 12)

    '行追加
    positionOfNameItem.Offset(1).Select
    Rows(ActiveCell.Row).Select
    Selection.Insert (xlShiftDown)

    '年度支給額計欄にWS関数追加
    With Cells.Find(what:=PART_TIME)
        If .Offset(1, 14) = "" Then
            monthlyTotalFomula = .Offset(1, 2).Resize(1, 12).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            .Offset(0, 14).Offset(1) = "=SUM(" & monthlyTotalFomula & ")"
        ElseIf .Offset(2, 14) = "" Then
            monthlyTotalFomula = .Offset(2, 2).Resize(1, 12).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            .Offset(0, 14).End(xlDown).Offset(1) = "=SUM(" & monthlyTotalFomula & ")"
        Else
            monthlyTotalFomula = .Offset(1, 2).End(xlDown).Offset.Resize(1, 12).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            .Offset(0, 14).End(xlDown).End(xlDown).Offset(1) = "=SUM(" & monthlyTotalFomula & ")"
        End If

        'ｱﾙﾊﾞｲﾄ・ﾊﾟｰﾄ月次計の数式を変数に格納
        If .Offset(2, 2) = "" Then
            monthlyTotalFomula = .Offset(1, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        Else
            monthlyTotalFomula = Range(.Offset(1, 2), .Offset(1, 2).End(xlDown)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        End If

        'ｱﾙﾊﾞｲﾄ・ﾊﾟｰﾄ月別振込人数の数式設定
        If .Offset(2, 2) = "" Then
            AdjustmentNumberOfPeople = .Offset(1, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        Else
            AdjustmentNumberOfPeople = Range(.Offset(1, 2), .Offset(1, 2).End(xlDown)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        End If
    End With

    With Range("A:A").Find(what:="ｱﾙﾊﾞｲﾄ･ﾊﾟｰﾄ月次計").Offset(0, 2)
        '一月分のアルバイト月次計を表示設定
        .Formula = "=SUM(" & monthlyTotalFomula & ")"
        'オートフィルで1年間分のアルバイト月次計を表示設定
        .AutoFill Destination:=.Resize(1, 12)
    End With

    With Cells.Find(what:=PART_TIME).Offset(0, 2)
        '一月分の精算人数を表示設定
        .Formula = "=COUNTIF(" & AdjustmentNumberOfPeople & " ,""<>0"")"
        'オートフィルで1年間分のアルバイト精算人数を表示設定
        .AutoFill Destination:=.Resize(1, 12)
    End With

    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate

End Sub
'バイト削除ボタン押下時の処理
Public Sub partTimeDeleteButton()

'***一人分のバイト給与欄を削除する***

    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate

    'バイトの給与欄の最上行を設定
    topRow = Cells.Find(what:="月給与合計", SearchDirection:=xlPrevious).Row - 19

    'バイト給与欄の「時給」が存在しない場合、後続処理を停止する。
    If (Cells.Find(what:="時給") Is Nothing) = True Then
        MsgBox "削除するバイト欄が存在しません", vbCritical
        Exit Sub
    End If

    'バイトの給与欄を削除する最下行を設定
    buttomRow = topRow + 18

    'バイト給与欄範囲設定
    Set parttimeSalaryItem = Range(Cells(topRow, LEFTMOST_COLUMN), Cells(buttomRow, RIGHTMOST_COLUMN))

    'バイト給与欄削除
    parttimeSalaryItem.delete (xlShiftUp)

    Worksheets("■振込額一覧").Activate

    Dim positionError As Range
    Set positionError = Cells.Find(what:=PART_TIME).Offset(1, 1)

    'バイト給与欄をすべて削除済みの場合、処理変更
    With Cells.Find(what:=PART_TIME).Offset(0, 1)
        If IsError(positionError) Then
            .End(xlDown).Select
        Else
            .End(xlDown).End(xlDown).Select
        End If
    End With
    Rows(ActiveCell.Row).Select
    Selection.delete (xlUp)

    '***ｱﾙﾊﾞｲﾄ・ﾊﾟｰﾄ月次計の数式を変数に格納***

    Dim monthlyTotalJudgmentPoint As Range
    Dim monthlyTotalStatingPoint As Range

    '判定する条件の位置
    Set monthlyTotalJudgmentPoint = Cells.Find(what:=PART_TIME).Offset(2, 2)

    '月次計の最初のセル設定
    Set monthlyTotalStatingPoint = Cells.Find(what:=PART_TIME).Offset(1, 2)

    If monthlyTotalJudgmentPoint.Offset(-1) = "" Then
        monthlyTotalFomula = 0
        Cells.Find(what:=PART_TIME).Offset(0, 2) = 0
        Cells.Find(what:=PART_TIME).Offset(0, 2).AutoFill Destination:=Cells.Find(what:=PART_TIME).Offset(0, 2).Resize(1, 12)
    ElseIf monthlyTotalJudgmentPoint = "" Then
        monthlyTotalFomula = monthlyTotalStatingPoint.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Else
        monthlyTotalFomula = Range(monthlyTotalStatingPoint, monthlyTotalStatingPoint.End(xlDown)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End If

    '一月分のバイト月次計を設定
    Range("A:A").Find(what:="ｱﾙﾊﾞｲﾄ･ﾊﾟｰﾄ月次計").Offset(0, 2).Formula = "=SUM(" & monthlyTotalFomula & ")"

    'オートフィルで1年間分のバイト月次計を表示
    Range("A:A").Find(what:="ｱﾙﾊﾞｲﾄ･ﾊﾟｰﾄ月次計").Offset(0, 2).AutoFill Destination:=Range("A:A").Find(what:="ｱﾙﾊﾞｲﾄ･ﾊﾟｰﾄ月次計").Offset(0, 2).Resize(1, 12)

    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate

End Sub
