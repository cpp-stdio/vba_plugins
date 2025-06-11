Attribute VB_Name = "Involved_Cell"
Option Explicit
'##############################################################################################################################
'
'   �V�[�g�֘A
'
'   �V�K�쐬�� : 2022/11/10
'   �ŏI�X�V�� : 2024/01/30
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2016 , 16.0.5.56.1000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'   LEGACY_aDeleteStrikethrough�Ŏg�p�����\����
'   �Q�lURL�Fhttps://vbabeginner.net/remove-strikethroughs-preserving-fonts/
'==============================================================================================================================
Private Type ST_FONT
    Background      As Long
    Bold            As Boolean
    Color           As Double
    ColorIndex      As Long
    FontStyle       As String
    Italic          As Boolean
    name            As String
    OutlineFont     As Boolean
    Shadow          As Boolean
    Size            As Double
    Strikethrough   As Boolean
    Subscript       As Boolean
    Superscript     As Boolean
    ThemeColor      As Variant
    ThemeFont       As XlThemeFont
    TintAndShade    As Double
    Underline       As Long
End Type
'==============================================================================================================================
'   ���������̂��������̂ݍ폜����i�o�O�L���������C���ς݁j
'   �Q�lURL�Fhttps://vbabeginner.net/remove-strikethroughs-preserving-fonts/
'
'   �g�����F
'               Dim r   As range    '// �Z��
'               For Each r In Selection
'                   Call aCallDeleteStrikethrough(r)
'               Next
'
'   �߂�l : �폜���ꂽ�Z��, �G���[�̏ꍇ��Nothing
'
'   r : �ΏۃZ��
'==============================================================================================================================
Public Function LEGACY_aDeleteStrikethrough(ByRef r As range) As range

    ' �߂�l�����ݒ�
    Set LEGACY_aDeleteStrikethrough = r

    Dim strike  As Variant
    strike = r.Font.Strikethrough
    
    '�Z�����ݒ莞�͏����I��
    If StrComp(CStr(r), "", vbBinaryCompare) = 0 Then Exit Function
    
    '�����������ݒ肳��Ă���t���O��False�̏ꍇ���������Ȃ�
    If Not IsNull(strike) And strike = False Then Exit Function
    
On Error GoTo aDeleteStrikethrough_ErrorHandler '���L�œ�G���[���������邱�Ƃ�����

    Dim i       As Long         ' �����񒷃��[�v�J�E���^
    Dim iLen    As Long         ' �Z��������
    Dim c       As Characters   ' �������Characters�I�u�W�F�N�g
    Dim f       As Font         ' 1�������Ƃ�Font�I�u�W�F�N�g
    Dim fAr()   As ST_FONT      ' Font�I�u�W�F�N�g�ݒ�l�ێ��p�̍\���̔z��
    Dim s       As String       ' �������������ς݂̕�����
    Dim iFont   As Long         ' Font�I�u�W�F�N�g�ݒ�p�z��̃C���f�b�N�X

    iFont = 0
    iLen = Len(CStr(r.value))
    ReDim fAr(iLen)
    
    '// �Z����������P���������[�v
    For i = 1 To iLen
        '// 1��������Characters�I�u�W�F�N�g���擾
        Set c = r.Characters(i, 1)
        
        '// Font�I�u�W�F�N�g���擾
        Set f = c.Font
        
        '// �Ώۂ̂P�����Ɏ����������ݒ肳��Ă��Ȃ��ꍇ
        If f.Strikethrough = False And Not StrComp("", CStr(c.text), vbBinaryCompare) = 0 Then
            '// �����������ݒ�̕�������擾
            s = s & c.text
            
            '// Font�I�u�W�F�N�g�̊e�v���p�e�B��ێ�
            fAr(iFont).name = f.name
            fAr(iFont).FontStyle = f.FontStyle
            fAr(iFont).Size = f.Size
            fAr(iFont).Strikethrough = f.Strikethrough
            fAr(iFont).Superscript = f.Superscript
            fAr(iFont).Subscript = f.Subscript
            fAr(iFont).OutlineFont = f.OutlineFont
            fAr(iFont).Shadow = f.Shadow
            fAr(iFont).Underline = f.Underline
            'fAr(iFont).ThemeColor = f.ThemeColor
            fAr(iFont).Color = f.Color
            fAr(iFont).TintAndShade = f.TintAndShade
            fAr(iFont).ThemeFont = f.ThemeFont
 
            iFont = iFont + 1
        End If
    Next
    
    '// ������������������������Z���ɐݒ�
    r.FormulaR1C1 = s
    
    '// �ēx�Z���̕����񒷂��擾
    iLen = Len(s)
    
    '// ������������������������P���������[�v
    For i = 1 To iLen
        '// 1��������Font�I�u�W�F�N�g���Đݒ�̂��ߎ擾
        Set f = r.Characters(Start:=i, length:=1).Font
        
        '// �C���f�b�N�X�擾
        iFont = i - 1
        
        '// Font�I�u�W�F�N�g�̊e�v���p�e�B��ێ����Ă������l�ōĐݒ�
        f.name = fAr(iFont).name
        f.FontStyle = fAr(iFont).FontStyle
        f.Size = fAr(iFont).Size
        f.Strikethrough = fAr(iFont).Strikethrough
        f.Superscript = fAr(iFont).Superscript
        f.Subscript = fAr(iFont).Subscript
        f.OutlineFont = fAr(iFont).OutlineFont
        f.Shadow = fAr(iFont).Shadow
        f.Underline = fAr(iFont).Underline
        'f.ThemeColor = fAr(iFont).ThemeColor
        f.Color = fAr(iFont).Color
        f.TintAndShade = fAr(iFont).TintAndShade
        f.ThemeFont = fAr(iFont).ThemeFont
    Next
    
    Set LEGACY_aDeleteStrikethrough = r
    Exit Function
    
aDeleteStrikethrough_ErrorHandler:
    Set LEGACY_aDeleteStrikethrough = Nothing
    
End Function

'==============================================================================================================================
'   ���������̂��������̂ݍ폜����i�y�ʔŁj
'   �����ӁF�Z���̏����܂ŃR�s�[�ł��Ȃ�
'
'   �Q�lURL�Fhttps://stabucky.com/wp/archives/3209
'
'   �g�����F
'               Dim r   As range    '// �Z��
'               For Each r In Selection
'                   Call aCallDeleteStrikethrough(r)
'               Next
'
'   �߂�l : �폜���ꂽ�Z��, �G���[�̏ꍇ��Nothing
'
'   r : �ΏۃZ��
'==============================================================================================================================
Public Function LEGACY_aDeleteStrikethrough_verLight(ByRef r As range) As range

    ' �߂�l�����ݒ�
    Set LEGACY_aDeleteStrikethrough_verLight = r

    Dim strike  As Variant
    strike = r.Font.Strikethrough
    
    '�Z�����ݒ莞�͏����I��
    If StrComp(CStr(r), "", vbBinaryCompare) = 0 Then Exit Function
    
    '�����������ݒ肳��Ă���t���O��False�̏ꍇ���������Ȃ�
    If Not IsNull(strike) And strike = False Then Exit Function
    
    Dim i As Long
    ' �e�L�X�g���擾
    Dim textBefore As String: textBefore = CStr(r)
    Dim textAfter As String: textAfter = ""
    
    For i = 1 To Len(textBefore)
        ' Strikethrough�̒l��False�̏ꍇ�̂ݎ��o��
        If r.Characters(Start:=i, length:=1).Font.Strikethrough = False Then
            textAfter = textAfter + Mid(textBefore, i, 1)
        End If
    Next i
    
    '�߂�l�̒l�̕��ɂ݂̂ɃZ�b�g
    LEGACY_aDeleteStrikethrough_verLight = textAfter
End Function
'==============================================================================================================================
'   NumberFormat��NumberFormatLocal�ŕ\�������������Ȃ��Ă���Z�����C�����邽��
'   �����������̂őS�đΉ��͂�����Ȃ����߁A�����ǉ������肢���܂�
'
'   �g�����F
'
'           If aTypeErrorIsNumberFormat(r) = 7 Then
'               ...
'               ...
'           End If
'
'   �߂�l : VarType : https://www.sejuku.net/blog/68632
'           No  �����ɓ����l  ���s����
'           1   Integer         2
'           2   Double          5
'           3   String          8
'           4   Boolean         11
'           5   Date            7
'           6   Object          9
'           7   Variant         0
'           8   String()        8200
'           9   Integer()       8194
'
'   r : �ΏۃZ��
'==============================================================================================================================
Public Function LEGACY_aTypeErrorIsNumberFormat(ByRef r As range) As Long
    
    '--------------------------------------------------------------
    '   Date�^�̏ꍇ
    LEGACY_aTypeErrorIsNumberFormat = 7

    '���t������
    If r.NumberFormatLocal Like "m""��""d""��""" Then Exit Function
    If r.NumberFormatLocal = "yyyy/mm/dd" Then Exit Function
    If r.NumberFormatLocal = "yyyy�Nmm��dd��" Then Exit Function
    If r.NumberFormatLocal = "ggge�Nmm��dd��" Then Exit Function
    If r.NumberFormatLocal = "yyyy/mm/dd(aaa)" Then Exit Function
    If r.NumberFormatLocal = "yyyy/mm/dd hh:mm:ss" Then Exit Function
    '����
    If r.NumberFormatLocal = "hh��mm��dd�b" Then Exit Function
    If r.NumberFormatLocal = "hh:mm:dd" Then Exit Function

    '--------------------------------------------------------------
    '   Double�^�̏ꍇ
    LEGACY_aTypeErrorIsNumberFormat = 5

    '���l
    If r.NumberFormatLocal = "#0.000" Then Exit Function
    If r.NumberFormatLocal = "#,##0" Then Exit Function
    '�ʉ�
    If r.NumberFormatLocal = "\#,##0" Then Exit Function
    If r.NumberFormatLocal = "#,##0�~" Then Exit Function

    '--------------------------------------------------------------
    '   String�^�̏ꍇ
    LEGACY_aTypeErrorIsNumberFormat = 8

    '������
    If r.NumberFormatLocal = "G/�W��" Then Exit Function
    If r.NumberFormatLocal = "@" Then Exit Function
    
    '--------------------------------------------------------------
    '������Ȃ������ꍇ�͕�����Ɖ��肷��
    LEGACY_aTypeErrorIsNumberFormat = 8
End Function
