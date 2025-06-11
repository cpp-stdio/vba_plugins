Attribute VB_Name = "Involved_Sheet"
Option Explicit
'##############################################################################################################################
'
'   �V�[�g�֘A
'
'   �V�K�쐬�� : 2017/08/30
'   �ŏI�X�V�� : 2024/07/05
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'   ���̖��O���V�[�g���ɓK�؂Ȗ��O�ł��邩��������
'
'   �߂�l : OK(True), NG(False)
'
'   sheetName : �V�[�g��
'==============================================================================================================================
Public Function LEGACY_checkSheetName(ByVal sheetname As String) As Boolean
    LEGACY_checkSheetName = False
    '��������1 : ��̖��O�ł͂Ȃ��B
    If StrComp(sheetname, "", vbBinaryCompare) = 0 Then Exit Function
    '��������2 : �܂�ł͂����Ȃ������񂪂Ȃ��B
    Dim textFor As Variant
    For Each textFor In Array(":", "\", "/", "?", "*", "[", "]")
        If InStr(sheetname, CStr(textFor)) > 0 Then Exit Function
    Next textFor
    '��������3 : ���O��31�����ȓ��ł���B
    If Len(sheetname) > 31 Then Exit Function
    '��������4 : �����̃V�[�g�͑��ݏo���Ȃ��B
    'aNewSheet�ɂĕs������������̂ŕ�������B
    LEGACY_checkSheetName = True
End Function
'==============================================================================================================================
'   ���������O�̃V�[�g��T��
'
'   �߂�l : ���������O�����V�[�g�B�Ȃ��ꍇ�́ANothing���ԋp�����
'
'   sheetName : �V�[�g��
'   book : �Ώۂ̃u�b�N�i�C�Ӂj
'==============================================================================================================================
Public Function LEGACY_sheetToEqualsName(ByVal sheetname As String, Optional ByRef book As Workbook = Nothing) As Worksheet

    Dim searchBook As Workbook
    Set searchBook = isBook(book)

    Dim sheet As Worksheet
    For Each sheet In searchBook.sheets
        If StrComp(sheet.name, sheetname, vbBinaryCompare) = 0 Then
            Set LEGACY_sheetToEqualsName = sheet
            Exit Function
        End If
    Next
    Set LEGACY_sheetToEqualsName = Nothing
End Function
'==============================================================================================================================
'   �V���ȃV�[�g���쐬
'
'   �߂�l :�V�K�쐬���ꂽWorksheet���ԋp����A�쐬�ς̏ꍇ�͂���Worksheet���ԋp�����B
'           �쐬�o���Ȃ������ꍇ��Nothing���ԋp�����
'
'   sheetName : �V�[�g��
'   book : �Ώۂ̃u�b�N�i�C�ӁA�����͂̏ꍇThisWorkbook�j
'==============================================================================================================================
Public Function LEGACY_aNewSheet(ByVal sheetname As String, Optional ByRef book As Workbook = Nothing) As Worksheet
    Set LEGACY_aNewSheet = Nothing
    '�K�؂Ȗ��O�łȂ��ꍇ
    If Not LEGACY_checkSheetName(sheetname) Then Exit Function
    '�Ώۂ̃u�b�N�����͂���Ă��Ȃ��ꍇ
    Dim addBook As Workbook
    Set addBook = isBook(book)
    '�쐬�ς݂�������
    Dim sheet As Worksheet
    Set sheet = LEGACY_sheetToEqualsName(sheetname, addBook)
    If Not sheet Is Nothing Then
        Set LEGACY_aNewSheet = sheet
        Exit Function
    End If
    '�V���ȃV�[�g�𖖔��֍쐬����
    addBook.sheets.add After:=Worksheets(Worksheets.count)
    Set sheet = addBook.sheets.Item(addBook.sheets.count)
    sheet.name = sheetname
    sheet.Activate '�A�N�e�B�u�����Ă��������������ڂ͗ǂ��B
    Set LEGACY_aNewSheet = sheet
End Function
'==============================================================================================================================
'   �u�b�N���ɂ���S�V�[�g�����擾
'
'   �߂�l :���������V�[�g����z��String�Ƃ��ĕԋp����
'           �쐬�o���Ȃ������ꍇ��Nothing���ԋp�����
'
'   book : �Ώۂ̃u�b�N�i�C�ӁA�����͂̏ꍇThisWorkbook�j
'==============================================================================================================================
Public Function LEGACY_getSheetNames(Optional ByRef book As Workbook = Nothing) As String()
    Dim r() As String
    Dim l As Long: l = 0
    '�Ώۂ̃u�b�N�����͂���Ă��Ȃ��ꍇ
    Dim getBook As Workbook
    Set getBook = isBook(book)
    Dim sheet As Worksheet
    For Each sheet In getBook.sheets
        ReDim Preserve r(l)
        r(l) = sheet.name
        l = l + 1
    Next
    LEGACY_getSheetNames = r
End Function
'==============================================================================================================================
'       long�^�̐��l�����ԍ�(AX��)���K�v�ɂȂ�ꍇ���Z���ɐ����𖄂ߍ���ő��x�㏸��_���ۂɂ��邪��s�ŏo�������������Ɣ��f��������
'
'   column : �ϊ�������Long�^
'==============================================================================================================================
Public Function LEGACY_isColumnNumber_toString(column As Long) As String
    LEGACY_isColumnNumber_toString = ""

    If column <= 0 Then Exit Function
    
    Dim tmp As Variant
    tmp = Split(Cells(1, column).Address(True, False), "$")
    LEGACY_isColumnNumber_toString = tmp(0)

End Function
'==============================================================================================================================
'   Long�^���̐��l����String�^(AX10)���̃A���t�@�x�b�g������^�̕ύX����
'   �s���Z�b�g�^�C�v
'
'   �߂�l : NG�^�� , OK�^�󔒈ȊO�̃A���t�@�x�b�g������
'
'       row    : �ϊ�������Long�^
'   column : �ϊ�������Long�^
'==============================================================================================================================
Public Function LEGACY_isColumnNumberAndRow_toString(row As Long, column As Long) As String
    LEGACY_isColumnNumberAndRow_toString = ""

    If row <= 0 Then Exit Function
    If column <= 0 Then Exit Function
    
    Dim tmp As Variant
    tmp = Split(Cells(row, column).Address(True, False), "$")
    LEGACY_isColumnNumberAndRow_toString = tmp(0) + tmp(1)

End Function
'==============================================================================================================================
'   �V�[�g���폜����
'
'   �߂�l : ����(True), ���s(False)
'
'   sheet : �폜����V�[�g�B���������ꍇ�A�A�N�Z�X�s�ɂȂ�̂Œ��ӂ��K�v
'   book  : �Ώۂ̃u�b�N�i�C�Ӂj
'==============================================================================================================================
Public Function LEGACY_aDeletedSheet(ByVal sheetname As String, Optional ByRef book As Workbook = Nothing) As Boolean
    Dim sheet As Worksheet
    Set sheet = LEGACY_sheetToEqualsName(sheetname, book)
    LEGACY_aDeletedSheet = LEGACY_aDeletedSheetEx(sheet, book)
    Set sheet = Nothing
End Function

Public Function LEGACY_aDeletedSheetEx(ByRef sheet As Worksheet, Optional ByRef book As Workbook = Nothing) As Boolean
    LEGACY_aDeletedSheetEx = False
    
    If sheet Is Nothing Then
        'Nothing�Ȃ̂ŁA���ɍ폜�ς݂Ɖ��肷��B
        LEGACY_aDeletedSheetEx = True
        Exit Function
    End If
    
    '�폜����^�C�~���O�Ń��b�Z�[�W���\������邪�@�\�I�ɕs�K�v�Ȃ̂Ŕ�\���ɂ��Ă���
    Application.DisplayAlerts = False
    Dim deleteBook As Workbook
    Set deleteBook = isBook(book)
    
    Dim deleteSheet As Worksheet
    For Each deleteSheet In deleteBook.sheets
        If StrComp(sheet.name, deleteSheet.name, vbBinaryCompare) = 0 Then
            Call deleteBook.sheets(sheet.name).delete
            Set sheet = Nothing  '�V�[�g���폜����
            LEGACY_aDeletedSheetEx = True '�߂�l��ύX
            Exit For
        End If
    Next
    
    '���b�Z�[�W��\����Ԃɖ߂�
    Application.DisplayAlerts = True
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   �V�[�g�̏���S�č폜����
'
'   sheet : �ΏۃV�[�g
'------------------------------------------------------------------------------------------------------------------------------
Public Function LEGACY_aInfoErasureSheet(ByRef sheet As Worksheet)
    Dim i As Long: i = 0
    '�Z����S�č폜
    sheet.Cells.clear
    sheet.Columns.clear
    sheet.Rows.clear
    '�e�[�u���̏����폜
    For i = sheet.ListObjects.count To 1 Step -1
        Call sheet.ListObjects.Item(i).delete
    Next i
    '���ߍ��݃O���t���폜
    For i = sheet.ChartObjects.count To 1 Step -1
        Call sheet.ChartObjects(i).delete
    Next i
    '������̃y�[�W��؂���폜
    'sheet.DisplayPageBreaks = False
    '�s�{�b�g�e�[�u�����폜
    For i = sheet.PivotTables.count To 1 Step -1
        Call sheet.PivotTables(i).ClearTable
    Next i
    '�}�A�N���b�v�A�[�g�A�}�`�ASmartArt�̍폜
    For i = sheet.Shapes.count To 1 Step -1
        Call sheet.Shapes.Item(i).delete
    Next i
    '�w�b�^�[�A�t�b�^�[�͊��S�ɍ폜���邱�Ƃ͕s�\�炵��
    With sheet.PageSetup
        For i = .Pages.count To 1 Step -1
            .Pages.Item(i).CenterFooter = ""
            .Pages.Item(i).CenterHeader = ""
            .Pages.Item(i).LeftFooter = ""
            .Pages.Item(i).LeftHeader = ""
            .Pages.Item(i).RightFooter = ""
            .Pages.Item(i).RightHeader = ""
        Next i
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .DifferentFirstPageHeaderFooter = True
    End With
    
End Function

'==============================================================================================================================
'   �V�[�g���œ��͂���Ă��鐔����l�ɕϊ�����
'   ��1�@���p����ꍇ�́uInvolved_Other.bas�v�̃C���|�[�g�����肢���܂�
'   ��2�@���������̎����Čv�Z���s���̂œ��삪�d���Ȃ�\��������܂�
'   ��3�@�uInvolved_Call�v�ɒP��Z���̐�����l�ɕϊ�����v���O����������܂��B
'
'   �߂�l : �ϊ�����(True), NG(False)
'
'   sheetName : �V�[�g��
'==============================================================================================================================
Public Function LEGACY_aSheetDeleteFormula(ByVal sheetname As String, Optional ByRef book As Workbook = Nothing) As Boolean
    LEGACY_aSheetDeleteFormula = False

    Dim sheet As Worksheet
    Set sheet = LEGACY_sheetToEqualsName(sheetname, book)
    If sheet Is Nothing Then Exit Function
    
    LEGACY_aSheetDeleteFormula = LEGACY_aSheetDeleteFormulaDx(sheet)
    Set sheet = Nothing
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   �V�[�g�p ver.
'
'   sheet : �V�[�g��}��(Nothing�̏ꍇ����)
'------------------------------------------------------------------------------------------------------------------------------
Public Function LEGACY_aSheetDeleteFormulaDx(ByRef sheet As Worksheet) As Boolean
    LEGACY_aSheetDeleteFormulaDx = False
    If sheet Is Nothing Then Exit Function
    
    Dim base As Range
    Dim cell As Range
    Dim row As Long
    Dim rowMax As Long
    Dim column As Long
    Dim columnMax As Long
    Dim text As String
    Dim value As Variant
    
    Set base = sheet.UsedRange.Range("A1")
    rowMax = sheet.UsedRange.Rows.count - 1
    columnMax = sheet.UsedRange.Columns.count - 1
    
    For row = rowMax To 0 Step -1
        For column = columnMax To 0 Step -1
            Set cell = base.Offset(row, column)
            
            If WorksheetFunction.IsFormula(cell) Then
                cell.Calculate '�Čv�Z
                text = cell.value
                '���l�̏ꍇ�͂��̂܂�"���l"�Ƃ��ĕ\��������i���t�A���z���͑ΏۊO�j
                If checkNumericalValue(text, value) Then
                    cell.value = value
                Else
                    cell.NumberFormatLocal = "@"
                    cell.value = text
                End If
            End If
        Next
    Next
    
    Set base = Nothing
    Set cell = Nothing
    LEGACY_aSheetDeleteFormulaDx = True
End Function
'==============================================================================================================================
'   ����̃V�[�g����ʃu�b�N�̂Ƃ��ē���̏ꏊ�ɕۑ�����
'
'   �߂�l : Workbook�iNG�̏ꍇ��Nothing���ԋp�����j
'
'   sheetname   : �V�[�g��
'   filename    : �ۑ����i�󔒂̏ꍇ�́u�V�[�g��.xlsx�v�ƂȂ�j
'   pathname    : �p�X���i�󔒂̏ꍇ�͖{�̂̃u�b�N�Ɠ��K�w�ɂȂ�j
'   book        : �u�b�N�iNothing�̏ꍇ��ThisWorkbook�Ƃ��Ă݂Ȃ��j
'==============================================================================================================================
Public Function LEGACY_saveSheet(ByVal sheetname As String, Optional ByVal fileName As String = "", _
                          Optional ByVal pathname As String = "", Optional ByRef book As Workbook = Nothing) As Workbook
    
    Set saveSheet = Nothing

    Dim sheet As Worksheet
    Set sheet = LEGACY_sheetToEqualsName(sheetname, book)
    
    Set LEGACY_saveSheet = LEGACY_saveSheetEx(sheet, fileName, pathname)
    Set sheet = Nothing
    
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   �V�[�g�p ver.
'
'   sheet       : �V�[�g�{��
'   filename    : �ۑ����i�󔒂̏ꍇ�́u�V�[�g��.xlsx�v�ƂȂ�j
'   pathname    : �p�X���i�󔒂̏ꍇ�͖{�̂̃u�b�N�Ɠ��K�w�ɂȂ�j
'------------------------------------------------------------------------------------------------------------------------------
Public Function LEGACY_saveSheetEx(ByRef sheet As Worksheet, Optional fileName As String = "", _
                                                      Optional pathname As String = "") As Workbook
                            
    Set LEGACY_saveSheetEx = Nothing

    If sheet Is Nothing Then Exit Function
    
    If StrComp(fileName, "", vbBinaryCompare) = 0 Then
        fileName = sheet.name + ".xlsx"
    End If
    
    If StrComp(pathname, "", vbBinaryCompare) = 0 Then
        pathname = ThisWorkbook.path
    End If
    
    sheet.copy                        '�ʂ̃u�b�N�փR�s�[
    
    Application.DisplayAlerts = False '���̊֐��𓮂����ƃ��b�Z�[�W���\������Ă��܂�����
    Call ActiveWorkbook.SaveAs(pathname + "\" + fileName)
    'Call ActiveWorkbook.Activate
    Application.DisplayAlerts = True  '���b�Z�[�W�\���h�~����

    Set LEGACY_saveSheetEx = ActiveWorkbook
    
End Function
    
'==============================================================================================================================
'   �u�b�N�̗L��
'==============================================================================================================================
Private Function isBook(ByRef book As Workbook) As Workbook
    If book Is Nothing Then
        Set isBook = ThisWorkbook
    Else
        Set isBook = book
    End If
End Function

