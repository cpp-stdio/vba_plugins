Attribute VB_Name = "Involved_Array"
Option Explicit
'##############################################################################################################################
'
'   �z��֘A�֐�
'   VBA�̔z��ɂ�2��ނ���BVariant�ŕύX�\�ȃ^�C�v�������łȂ��^�C�v�B����ɂ��֐���2��ޕK�v�ɂȂ�B
'
'   �V�K�쐬�� : 2019/11/18
'   �ŏI�X�V�� : 2024/01/30
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'
'   �z�񂪋�Ȃ̂��𔻒肷��
'   ���̊֐���VBA�̎d�l��A�ǂ����֐������邱�Ƃ��o���Ȃ��ׁA�قړ����R�[�h��2�񏑂��K�v������B
'   �Q�lURL : http://www.fingeneersblog.com/1612/
'
'   �߂�l : ��(true),��ł͂Ȃ�(false)
'
'   arrayVariant : ����p�̔z��
'
'==============================================================================================================================
Public Function LEGACY_arrayIsEmpty(ByRef arrayVariant As Variant) As Boolean
    LEGACY_arrayIsEmpty = True '�󂾂Ɖ���
On Error GoTo isEmptyArray_ErrorHandler

    'UBound�֐����g�p���ăG���[���������邩�ǂ������m�F
    If UBound(arrayVariant) > 0 Then
        LEGACY_arrayIsEmpty = False
    End If
    Exit Function
    
isEmptyArray_ErrorHandler:
    LEGACY_arrayIsEmpty = True
End Function

Public Function LEGACY_arrayIsEmptyEx(ByRef arrayVariant() As Variant) As Boolean
    LEGACY_arrayIsEmptyEx = True '�󂾂Ɖ���
On Error GoTo isEmptyArrayEx_ErrorHandler

    'UBound�֐����g�p���ăG���[���������邩�ǂ������m�F
    If UBound(arrayVariant) > 0 Then
        LEGACY_arrayIsEmptyEx = False
    End If
    Exit Function

isEmptyArrayEx_ErrorHandler:
    LEGACY_arrayIsEmptyEx = True
End Function

'==============================================================================================================================
'
'   �z��̈ꕔ��؂�o���A�V�����z��Ƃ��ĕԋp����B
'
'   �߂�l : ����(True), ���s(False)
'
'   oldArray : �؂�o���p�̔z��
'   newArray : �ԋp�p�z��
'   min      : �ǂ�����
'   max      : �ǂ��܂�
'==============================================================================================================================
Public Function LEGACY_arraySplit(ByRef oldArray As Variant, ByRef newArray As Variant, Optional ByVal min As Long = -&HFF, Optional ByVal max As Long = -&HFF) As Boolean
    LEGACY_arraySplit = False '���s�Ɖ���
    If LEGACY_arrayIsEmpty(oldArray) Then Exit Function
    If errorSplit(min, max, LBound(oldArray), UBound(oldArray)) Then Exit Function
    'VBA�̎d�l�ケ�������͌ʂŏ����Ȃ���΂Ȃ�Ȃ��B
    Dim i As Long
    Dim length As Long: length = -1
    
    If VarType(newArray) = vbEmpty Then
        newArray = Array()
    End If
    
    For i = min To max
        length = length + 1
        ReDim Preserve newArray(length)
        newArray(length) = oldArray(i)
    Next i
    
    LEGACY_arraySplit = True
End Function

Public Function LEGACY_arraySplitEx(ByRef oldArray() As Variant, ByRef newArray() As Variant, Optional ByVal min As Long = -&HFF, Optional ByVal max As Long = -&HFF) As Boolean
    LEGACY_arraySplitEx = False '���s�Ɖ���
    If LEGACY_arrayIsEmptyEx(oldArray) Then Exit Function
    If errorSplit(min, max, LBound(oldArray), UBound(oldArray)) Then Exit Function
    'VBA�̎d�l�ケ�������͌ʂŏ����Ȃ���΂Ȃ�Ȃ��B
    Dim i As Long
    Dim length As Long: length = -1
    For i = min To max
        length = length + 1
        ReDim Preserve newArray(length)
        newArray(length) = oldArray(i)
    Next i
    
    LEGACY_arraySplitEx = True
End Function

Private Function errorSplit(ByRef min As Long, ByRef max As Long, ByVal minArray As Long, ByVal maxArray As Long) As Boolean
    errorSplit = True

    If min < minArray Then
        min = minArray
    End If
    
    If max > maxArray Then
        max = maxArray
    End If
    
    'VBA�̎d�l�œ��������ł�OK�Ƃ���B
    If min < max Then Exit Function
    
    errorSplit = False
End Function

'==============================================================================================================================
'
'   �z��̔��]
'   ���̊֐���VBA�̎d�l��A�ǂ����֐������邱�Ƃ��o���Ȃ��ׁA�قړ����R�[�h��2�񏑂��K�v������B
'
'   �߂�l : ����(True), ���s(False)
'
'   reversed : ���]����z��
'
'==============================================================================================================================
Public Function LEGACY_arrayReversed(ByRef oldArray As Variant, ByRef newArray As Variant) As Boolean
    LEGACY_arrayReversed = False
    If LEGACY_arrayIsEmpty(oldArray) Then Exit Function
    
    'oldArray��newArray���������ƃ�������j�󂵂Ă��܂���
    Dim old As Variant
    old = LEGACY_arrayCopy(oldArray)
    
    ReDim newArray(UBound(old))
    
    Dim i As Long
    For i = LBound(old) To UBound(old)
        newArray(UBound(old) - i) = old(i)
    Next i
    LEGACY_arrayReversed = True
    
End Function

Public Function LEGACY_arrayReversedEx(ByRef oldArray() As Variant, ByRef newArray() As Variant) As Boolean
    LEGACY_arrayReversedEx = False
    If LEGACY_arrayIsEmptyEx(oldArray) Then Exit Function
    
    'oldArray��newArray���������ƃ�������j�󂵂Ă��܂���
    Dim old() As Variant
    old = LEGACY_arrayCopyEx(oldArray)
    
    ReDim newArray(UBound(old))
    
    Dim i As Long
    For i = LBound(old) To UBound(old)
        newArray(UBound(old) - i) = old(i)
    Next i
    LEGACY_arrayReversedEx = True
End Function

'==============================================================================================================================
'
'   �z��̃R�s�[
'   ���̊֐���VBA�̎d�l��A�ǂ����֐������邱�Ƃ��o���Ȃ��ׁA�قړ����R�[�h��2�񏑂��K�v������B
'
'   �߂�l : �R�s�[�����z��
'
'   copy : ���]����z��
'
'==============================================================================================================================
Public Function LEGACY_arrayCopy(ByRef copy As Variant) As Variant
    arrayCopy = Empty
    If LEGACY_arrayIsEmpty(copy) Then Exit Function

    Dim c As Variant
    ReDim c(UBound(copy))
    
    Dim i As Long
    For i = LBound(copy) To UBound(copy)
        c(i) = copy(i)
    Next i
    LEGACY_arrayCopy = c
End Function

Public Function LEGACY_arrayCopyEx(ByRef copy() As Variant) As Variant()
    Dim c() As Variant
    arrayCopyEx = c
    
    If LEGACY_arrayIsEmptyEx(copy) Then Exit Function

    ReDim c(UBound(copy))
    
    Dim i As Long
    For i = LBound(copy) To UBound(copy)
        c(i) = copy(i)
    Next i
    LEGACY_arrayCopyEx = c
End Function
