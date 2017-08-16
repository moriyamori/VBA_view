Attribute VB_Name = "ClickSideWorkButton"
Dim parttimeSalaryItem As Range
'�o�C�g�̒ǉ��{�^���������̏���
Public Sub partTimeAddButton()

'�u��2017�N�x�@�Ј����^�ڍׁv�V�[�g���A�N�e�B�u
    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate

'***�o�C�g�̋��^����ǉ�***

    '�o�C�g�̋��^����ǉ�����ŏ�s��ݒ�
    topRow = Cells.Find(what:="�����^���v", SearchDirection:=xlPrevious).Row

    '�o�C�g�̋��^����ǉ�����ŉ��s��ݒ�
    buttomRow = topRow + 18

    '�o�C�g���^���͈͐ݒ�
    Set parttimeSalaryItem = Range(Cells(topRow, LEFTMOST_COLUMN), Cells(buttomRow, RIGHTMOST_COLUMN))

    '��l���̋��^���g���C���T�[�g
    parttimeSalaryItem.Insert (xlShiftDown)

    '�uPartData�v���[�N�V�[�g���狋�^���e���v�����擾
    ThisWorkbook.Worksheets("PartData").Range(RANGE_OF_PART_TIME_SALARY_ITEMS).Copy

    '�o�C�g�̋��^�����w��̈ʒu�ɓ\��t��
    Cells(topRow, 1).PasteSpecial (xlPasteAll)

    '�R�s�[�p�_������
    Application.CutCopyMode = False

'***�u�U���z�ꗗ�v�V�[�g�ɁA�ǉ��������^���̐U���z�̑��v��ǉ�����B***

    Dim positionOfNameItem As Range
    Dim positionOfTransferAmountItem As Range

    '���^�ڍ׃V�[�g���疼�O���̃A�h���X���擾
    nameItemeAddress = Cells(topRow, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)

    '���^�ڍ׃V�[�g����Ђƌ����̐U���z�̂���A�h���X�擾
    positionOfTransferAmount = Cells(buttomRow, 4).Address(RowAbsolute:=False, ColumnAbsolute:=False)

    Worksheets("���U���z�ꗗ").Activate

    '�U���z�̘g�ǉ�---

    '�����A�A���o�C�g�E�p�[�g����l�����Ȃ��ꍇor��l�������Ȃ��ꍇ�͈͕̔ύX�ݒ�
    With Cells.Find(what:=PART_TIME)
        '0�l
        If .Offset(1, 2) = "" Then
            Set positionOfNameItem = .Offset(1, 1)
            Set positionOfTransferAmountItem = .Offset(1, 2)
            '1�l
        ElseIf Cells.Find(what:=PART_TIME).Offset(2, 2) = "" Then
            Set positionOfNameItem = .Offset(2, 1)
            Set positionOfTransferAmountItem = .Offset(2, 2)
            '����ȊO
        Else
            Set positionOfNameItem = .Offset(0, 1).End(xlDown).End(xlDown).Offset(1)
            Set positionOfTransferAmountItem = .Offset(0, 1).End(xlDown).End(xlDown).Offset(1, 1)
        End If
    End With

    '�u���U���z�ꗗ�v�V�[�g�̖��O����WS�֐��ǉ�
    bindingResult = MEIN_SHEET_NAME2 + nameItemeAddress
    bindingResult = "=IF(" & bindingResult & "="""",""""," & bindingResult & ")"
    positionOfNameItem = bindingResult

    '�u���U���z�ꗗ�v�V�[�g�̂P�P���U���z����WS�֐��ǉ�
    bindingResult = MEIN_SHEET_NAME + positionOfTransferAmount
    positionOfTransferAmountItem = bindingResult

    '�I�[�g�t�B����1�N�ԕ��\��
    positionOfTransferAmountItem.AutoFill Destination:=positionOfTransferAmountItem.Resize(1, 12)

    '�s�ǉ�
    positionOfNameItem.Offset(1).Select
    Rows(ActiveCell.Row).Select
    Selection.Insert (xlShiftDown)

    '�N�x�x���z�v����WS�֐��ǉ�
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

        '���޲āE�߰Č����v�̐�����ϐ��Ɋi�[
        If .Offset(2, 2) = "" Then
            monthlyTotalFomula = .Offset(1, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        Else
            monthlyTotalFomula = Range(.Offset(1, 2), .Offset(1, 2).End(xlDown)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        End If

        '���޲āE�߰Č��ʐU���l���̐����ݒ�
        If .Offset(2, 2) = "" Then
            AdjustmentNumberOfPeople = .Offset(1, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        Else
            AdjustmentNumberOfPeople = Range(.Offset(1, 2), .Offset(1, 2).End(xlDown)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        End If
    End With

    With Range("A:A").Find(what:="���޲ĥ�߰Č����v").Offset(0, 2)
        '�ꌎ���̃A���o�C�g�����v��\���ݒ�
        .Formula = "=SUM(" & monthlyTotalFomula & ")"
        '�I�[�g�t�B����1�N�ԕ��̃A���o�C�g�����v��\���ݒ�
        .AutoFill Destination:=.Resize(1, 12)
    End With

    With Cells.Find(what:=PART_TIME).Offset(0, 2)
        '�ꌎ���̐��Z�l����\���ݒ�
        .Formula = "=COUNTIF(" & AdjustmentNumberOfPeople & " ,""<>0"")"
        '�I�[�g�t�B����1�N�ԕ��̃A���o�C�g���Z�l����\���ݒ�
        .AutoFill Destination:=.Resize(1, 12)
    End With

    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate

End Sub
'�o�C�g�폜�{�^���������̏���
Public Sub partTimeDeleteButton()

'***��l���̃o�C�g���^�����폜����***

    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate

    '�o�C�g�̋��^���̍ŏ�s��ݒ�
    topRow = Cells.Find(what:="�����^���v", SearchDirection:=xlPrevious).Row - 19

    '�o�C�g���^���́u�����v�����݂��Ȃ��ꍇ�A�㑱�������~����B
    If (Cells.Find(what:="����") Is Nothing) = True Then
        MsgBox "�폜����o�C�g�������݂��܂���", vbCritical
        Exit Sub
    End If

    '�o�C�g�̋��^�����폜����ŉ��s��ݒ�
    buttomRow = topRow + 18

    '�o�C�g���^���͈͐ݒ�
    Set parttimeSalaryItem = Range(Cells(topRow, LEFTMOST_COLUMN), Cells(buttomRow, RIGHTMOST_COLUMN))

    '�o�C�g���^���폜
    parttimeSalaryItem.delete (xlShiftUp)

    Worksheets("���U���z�ꗗ").Activate

    Dim positionError As Range
    Set positionError = Cells.Find(what:=PART_TIME).Offset(1, 1)

    '�o�C�g���^�������ׂč폜�ς݂̏ꍇ�A�����ύX
    With Cells.Find(what:=PART_TIME).Offset(0, 1)
        If IsError(positionError) Then
            .End(xlDown).Select
        Else
            .End(xlDown).End(xlDown).Select
        End If
    End With
    Rows(ActiveCell.Row).Select
    Selection.delete (xlUp)

    '***���޲āE�߰Č����v�̐�����ϐ��Ɋi�[***

    Dim monthlyTotalJudgmentPoint As Range
    Dim monthlyTotalStatingPoint As Range

    '���肷������̈ʒu
    Set monthlyTotalJudgmentPoint = Cells.Find(what:=PART_TIME).Offset(2, 2)

    '�����v�̍ŏ��̃Z���ݒ�
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

    '�ꌎ���̃o�C�g�����v��ݒ�
    Range("A:A").Find(what:="���޲ĥ�߰Č����v").Offset(0, 2).Formula = "=SUM(" & monthlyTotalFomula & ")"

    '�I�[�g�t�B����1�N�ԕ��̃o�C�g�����v��\��
    Range("A:A").Find(what:="���޲ĥ�߰Č����v").Offset(0, 2).AutoFill Destination:=Range("A:A").Find(what:="���޲ĥ�߰Č����v").Offset(0, 2).Resize(1, 12)

    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate

End Sub
