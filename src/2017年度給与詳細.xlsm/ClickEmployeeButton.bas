Attribute VB_Name = "ClickEmployeeButton"
    Dim topRow As Integer
    Dim buttomRow As Integer
    Dim employeeSalaryItemRange As Range
    Dim monthlyTotalFomula As String
    Dim nameItemeAddress As String
    Dim positionOfTransferAmount As String
    Dim bindingResult As String
    Dim AdjustmentNumberOfPeople As String
       
'�Ј��ǉ��{�^���������̏���
Public Sub employeeAddButton()

    '�Ј����^�ڍ׃A�N�e�B�u
    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate

    '�Ј��̋��^���ǉ�����ŏ�s��ݒ�
    topRow = Range("C:C").Find(what:="��{��", SearchDirection:=xlPrevious).Row + ROW_TO_THE_TOP_ROW

    '�Ј��̋��^���ǉ�����ŉ��s��ݒ�
    buttomRow = topRow + 18

    '�Ј����^���͈�
    Set employeeSalaryItemRange = Range(Cells(topRow, LEFTMOST_COLUMN), Cells(buttomRow, RIGHTMOST_COLUMN))

    '��l���̋��^���Z�����C���T�[�g
    employeeSalaryItemRange.Insert (xlShiftDown)

    '�f�[�^���[�N�V�[�g���狋�^���e���v�����擾
    ThisWorkbook.Worksheets("EmployeeData").Range(EMPLOYEE_SALARY_ITEMS).Copy

    '���^���e���v�����w��̈ʒu�ɓ\��t��
    Cells(topRow, 1).PasteSpecial (xlPasteAll)

    '�R�s�[�_������
    Application.CutCopyMode = False

    '***�u�U���z�ꗗ�v�V�[�g�ɒǉ��������^���̐U���z���v��ǉ�����***

    '�u��2017�N�x�@�Ј����^�ڍׁv�V�[�g���疼�O���̈ʒu���擾
    nameItemeAddress = Cells(topRow, 1).Address(ReferenceStyle:=xlR1C1)

    '�u��2017�N�x�@�Ј����^�ڍׁv�V�[�g����U���z�̂���ʒu���擾
    positionOfTransferAmount = Cells(buttomRow, 4).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    Worksheets("���U���z�ꗗ").Activate

    '�U���z�̘g�ǉ�
    With Range("A:A").Find(what:="�Ј�")
        .Offset(0, 1).End(xlDown).End(xlDown).Offset(1).Select
        Rows(ActiveCell.Row).Select
        Selection.Insert (xlShiftDown)

        '�u���U���z�ꗗ�v�V�[�g�ɖ��O���ʒu�̐�����ǉ�
        bindingResult = MEIN_SHEET_NAME2 + nameItemeAddress
        bindingResult = "=IF(" & bindingResult & "="""",""""," & bindingResult & ")"
        .Offset(0, 1).End(xlDown).End(xlDown).Offset(1) = bindingResult

        '�u���U���z�ꗗ�v�V�[�g�̈ꌎ��
        bindingResult = MEIN_SHEET_NAME + positionOfTransferAmount
        .Offset(0, 1).End(xlDown).End(xlDown).Offset(0, 1) = bindingResult

        '�I�[�g�t�B����1�N�ԕ��\��
        .Offset(0, 1).End(xlDown).End(xlDown).Offset(0, 1).AutoFill Destination:= _
        .Offset(0, 1).End(xlDown).End(xlDown).Offset(0, 1).Resize(1, 12)
    End With

    '�N�x�x���z�v��ǉ�
    Cells.Find(what:="�N�x�x���z�v").End(xlDown).End(xlDown).AutoFill Destination:=Cells.Find(what:="�N�x�x���z�v").End(xlDown).End(xlDown).Resize(2)

    '�{�[�i�X�����ݒ�---
    
    Dim positionLncledingBonus As String

    positionLncledingBonus = ThisWorkbook.Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Cells(buttomRow, 17).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    bindingResult = MEIN_SHEET_NAME + positionLncledingBonus
    ThisWorkbook.Worksheets("���U���z�ꗗ").Cells.Find(what:="�{�[�i�X����").End(xlDown).End(xlDown).Offset(1) = bindingResult

    '�Ј������v�ݒ�---

    ' �Ј������v�̐�����ݒ�
    monthlyTotalFomula = Range(Range("C8"), Range("C8").End(xlDown)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Range("A:A").Find(what:="�Ј������v").Offset(0, 2).Formula = "=SUM(" & monthlyTotalFomula & ")"
    
    '�ꌎ�̐��Z�l�����͈̔͐ݒ�
    monthlyTotalFomula2 = Range(Range("C7"), Range("C7").End(xlDown)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    '�ꌎ���̐��Z�l��
    Range("A:A").Find(what:="���Ј�").Offset(0, 2).Formula = "=COUNTIF(" & monthlyTotalFomula2 & " ,""<>0"")"
           
    '�I�[�g�t�B����1�N�ԕ��̎Ј������v��\��
    Range("A:A").Find(what:="�Ј������v").Offset(0, 2).AutoFill Destination:=Range("A:A").Find(what:="�Ј������v").Offset(0, 2).Resize(1, 12)

    Range("A:A").Find(what:="���Ј�").Offset(0, 2).AutoFill Destination:=Range("A:A").Find(what:="���Ј�").Offset(0, 2).Resize(1, 12)
    
    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate


End Sub

'�Ј��폜�{�^���������̏���
Public Sub employeeDeleteButton()

    '��l���̋��^���̍ŏ�s��ݒ�---
    
    '�u��2017�N�x�@�Ј����^�ڍׁv�A�N�e�B�u
    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate

    '�Ј��̋��^�����폜����ŏ�s��ݒ�
    topRow = Range("C:C").Find(what:="��{��", SearchDirection:=xlPrevious).Row

    '�ŏ�s����l�������Ȃ��ꍇ�̍s�ʒu�������ꍇ�㑱�������~����B
    If topRow = 0 Or topRow = 14 Then
        MsgBox "����ȏ�폜�ł��܂���B", vbCritical
        Exit Sub
    End If

    '�Ј��̋��^�����폜����ŉ��s��ݒ�
    buttomRow = topRow + 18

    '�Ј����^���͈͐ݒ�
    Set employeeSalaryItemRange = Range(Cells(topRow, LEFTMOST_COLUMN), Cells(buttomRow, RIGHTMOST_COLUMN))

    '�Ј��̋��^���g��ݒ�
    employeeSalaryItemRange.delete (xlShiftUp)

    '�U���z�폜
    Worksheets("���U���z�ꗗ").Activate
    Cells.Find(what:="�Ј�").Offset(0, 1).End(xlDown).End(xlDown).Select
    Rows(ActiveCell.Row).Select
    Selection.delete (xlUp)

    '***�Ј������v�̐������i�[***
    
    '�Ј������v�̐l������l�̏ꍇ�A�����ύX
    If Range("C9") = "" Then
        monthlyTotalFomula = Range("C8").Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Else
        monthlyTotalFomula = Range(Range("C8"), Range("C8").End(xlDown)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End If
    With Range("A:A").Find(what:="�Ј������v").Offset(0, 2)
        .Formula = "=SUM(" & monthlyTotalFomula & ")"
        '�I�[�g�t�B����1�N�ԕ��̎Ј������v��\��
        .AutoFill Destination:=Range("A:A").Find(what:="�Ј������v").Offset(0, 2).Resize(1, 12)
    End With

    Worksheets(SHEET_EMPLOYEE_SALARY_DETAILS).Activate

End Sub
