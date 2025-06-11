Attribute VB_Name = "Involved_Split"
Option Explicit
'##############################################################################################################################
'
'   �����񕪊��֘A
'
'   �V�K�쐬�� : 2017/08/30
'   �ŏI�X�V�� : 2024/01/30
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'   �����̒��ɂ���A����̕����񂩂����̕�����܂ł��擾����
'   �߂�l : ��������������ASplit�Ɩ������{�ƂƂ͈Ⴂ�uspecificA�v�ƁuspecificB�v�������Ƃ��ĕK�{�Ȃ̂Œ���
'
'   text      : �Ƃ��镶����
'   specificA : 1�ڂ̓���̕�����
'   specificB : 2�ڂ̓���̕�����
'
'   ���
'       Dim text As String: text = "<HTML><HEAD>HOGEHOGE</HEAD><HEAD>GEHOGEHO</HEAD></HTML>"
'       Dim textArray() As String
'       textArray = BetweenSplit(text, "<HEAD>", "</HEAD>")
'       > ("<HTML>","<HEAD>","HOGEHOGE","</HEAD>","<HEAD>","GEHOGEHO","</HEAD>","</HTML>")
'==============================================================================================================================
Public Function LEGACY_BetweenSplit(ByVal text As String, ByVal specificA As String, ByVal specificB As String) As String()
    Dim returnLength As Long: returnLength = 0
    Dim returnString() As String
    ReDim returnString(returnLength)
    '�G���[�Ή��̂��ߖ߂�l��������
    LEGACY_BetweenSplit = returnString
    '�󔒂̑}�����m�F
    If StrComp(text, "", vbBinaryCompare) = 0 Then Exit Function
    If StrComp(specificA, "", vbBinaryCompare) = 0 Then Exit Function
    If StrComp(specificB, "", vbBinaryCompare) = 0 Then Exit Function
    '����������Ȃ�p�r���Ⴄ��
    If StrComp(specificA, specificB, vbBinaryCompare) = 0 Then
        LEGACY_BetweenSplit = Split(text, specificA)
        Exit Function
    End If

    '------------------------------
    ' "specificA" ���̏���
    '------------------------------
    Dim textArray1() As String
    textArray1 = Split(text, specificA)
    '------------------------------
    ' "specificB" ���̏���
    '------------------------------
    Dim textArray2() As String
    Dim count1 As Long: count1 = 0
    Dim count2 As Long: count2 = 0

    Dim Body As String
    For count1 = 0 To UBound(textArray1)
        '�󔒂̏ꍇ�͏������s���K�v���Ȃ���
        If Not StrComp(textArray1(count1), "", vbBinaryCompare) = 0 Then
            textArray2 = Split(textArray1(count1), specificB)
            '�ŏ��ɃJ�b�g���ꂽ�AspecificA�̒l������
            If Not count1 = 0 Then
                ReDim Preserve returnString(returnLength)
                returnString(returnLength) = specificA
                returnLength = returnLength + 1
            End If
            '�z�񐔂�0�ȉ��̏ꍇ�A�������Ȃ������̂ŁA���̂܂ܑ}������B
            If UBound(textArray2) <= 0 Then
                ReDim Preserve returnString(returnLength)
                returnString(returnLength) = textArray1(count1)
                returnLength = returnLength + 1
            Else
                For count2 = 0 To UBound(textArray2)
                    '�J�b�g���ꂽspecificB��߂�
                    If Not count2 = 0 Then
                        ReDim Preserve returnString(returnLength)
                        returnString(returnLength) = specificB
                        returnLength = returnLength + 1
                    End If
                    '�J�b�g����Ȃ������������ɖ߂��B
                    If Not StrComp(textArray2(count2), "", vbBinaryCompare) = 0 Then
                        ReDim Preserve returnString(returnLength)
                        returnString(returnLength) = textArray2(count2)
                        returnLength = returnLength + 1
                    End If
                Next count2
            End If
        End If
    Next count1
    LEGACY_BetweenSplit = returnString
End Function

'==============================================================================================================================
'   Split�֐��̕�����
'
'   delimiters : �{�ƂƂ͈Ⴂ�AOptional�^�łȂ��A�z��łȂ��ƃG���[�o��̂Œ���
'   min        : delimiters�̂ǂ̈ʒu�����؂肷��̂� : ���̐��Adelimiters�ȏ�̓G���[
'   max        : delimiters�̂ǂ̈ʒu�܂ŋ�؂肷��̂� : ���̐��͑S�ċ�؂�Adelimiters�ȏ�ł��S�ċ�؂�
'
'   ���̑��A�����̐����͉��LURL���Q��
'   https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/split-function
'==============================================================================================================================
Public Function LEGACY_Splits(ByVal expression As String, delimiters() As String, Optional ByVal limit As Long = -1, Optional ByVal compare As VbCompareMethod = vbBinaryCompare, Optional ByVal min As Long = 0, Optional ByVal max As Long = -1) As String()
    Dim returnString() As String
    ReDim returnString(0)
    LEGACY_Splits = returnString
    If UBound(delimiters) < 0 Then Exit Function
    If min < 0 Then Exit Function
    If max < 0 Or max > UBound(delimiters) Then max = UBound(delimiters)

    Dim returnLength As Long: returnLength = 0
    Dim textCount As Long, textArray() As String
    Dim bodyCount As Long, bodyArray() As String
    Dim limitCount As Long, limitString As String
    '-1�̕�����limit�Ȃ̂ł���ŗǂ�
    textArray = Split(expression, delimiters(min), -1, compare)
    LEGACY_Splits = textArray
    If min = max Then Exit Function
    If min >= UBound(delimiters) Then Exit Function

    For textCount = 0 To UBound(textArray)
        If Not StrComp(textArray(textCount), "", vbBinaryCompare) = 0 Then
            '-1�̕�����limit�Ȃ̂ł���ŗǂ�
            bodyArray = LEGACY_Splits(textArray(textCount), delimiters, -1, compare, min + 1, max)
            For bodyCount = 0 To UBound(bodyArray)
                If Not StrComp(bodyArray(bodyCount), "", vbBinaryCompare) = 0 Then
                    ReDim Preserve returnString(returnLength)
                    returnString(returnLength) = bodyArray(bodyCount)
                    returnLength = returnLength + 1

                    If limit >= 0 And returnLength >= limit Then
                        limitString = ""
                        For limitCount = bodyCount To UBound(bodyArray)
                            limitString = limitString + bodyArray(limitCount)
                        Next limitCount

                        returnString(returnLength - 1) = limitString
                        LEGACY_Splits = returnString
                        Exit Function
                    End If
                End If
            Next bodyCount
        End If
    Next
    LEGACY_Splits = returnString
End Function
