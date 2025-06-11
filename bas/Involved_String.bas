Attribute VB_Name = "Involved_String"
Option Explicit
'##############################################################################################################################
'
'   ������(String)��VBA�̕W���@�\�����ł͑���Ȃ�������ǉ�����
'   �� 2024/01/30�FInvolved_Other����Ɨ�
'
'   �V�K�쐬�� : 2024/01/30
'   �ŏI�X�V�� : 2024/01/30
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Microsoft 365 Apps for enterprise
'   �ŏI�X�V�G�N�Z���o�[�W���� : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'
'   ���l�𔻒�
'   �߂�l : �͂�(true),������(false)
'
'   text  : ����p�̐��l
'   value : �����l�̓��������l�^(Long,Double)�̂ǂ��炩�A�G���[�̏ꍇ��Empty������
'           �ŏI�I�ɂ͌^�̔��肪�v��܂��B���Q�lURL�F�ၨ If VarType(value) = vbLong Then
'           http://officetanaka.net/excel/vba/function/VarType.htm
'
'==============================================================================================================================
Public Function LEGACY_checkNumericalValue(ByVal text As String, Optional ByRef value As Variant = Empty) As Boolean

    text = StrConv(text, vbNarrow)
    text = StrConv(text, vbLowerCase)
    text = LCase(text)
    If IsNumeric(text) Then
        value = Val(text)
        If StrComp(CStr(value), CStr(CLng(CStr(value))), vbBinaryCompare) = 0 Then
            value = CLng(CStr(value))
        End If
        LEGACY_checkNumericalValue = True
    Else
        value = Empty
        LEGACY_checkNumericalValue = False
    End If
End Function

'==============================================================================================================================
'
'   ������̒�����A�����݂̂𔲂��o���B�Q�lURL��
'   https://vbabeginner.net/vba%E3%81%A7%E6%96%87%E5%AD%97%E5%88%97%E3%81%8B%E3%82%89%E6%95%B0%E5%AD%97%E3%81%AE%E3%81%BF%E3%82%92%E6%8A%BD%E5%87%BA%E3%81%99%E3%82%8B/
'
'   �߂�l : �����o���������A�G���[�̏ꍇ�͋�̔z�񂪕ԋp����܂��B
'
'   text  : �������܂܂�镶����
'
'==============================================================================================================================
Public Function LEGACY_findNumber(ByVal text As String) As Variant()
    Dim reg As Object     '���K�\���N���X�I�u�W�F�N�g
    Dim matches As Object 'RegExp.Execute����
    Dim match As Object   '�������ʃI�u�W�F�N�g
    Dim i As Long         '���[�v�J�E���^
    
    Dim returnVariant() As Variant
    ReDim returnVariant(0)
    LEGACY_findNumber = returnVariant
    
    Set reg = CreateObject("VBScript.RegExp")
    
    '�����͈́�������̍Ō�܂Ō���
    reg.Global = True
    '��������������������
    reg.Pattern = "[0-9]"
    '�������s
    Set matches = reg.Execute(text)
    '������v�����������[�v
    For i = 0 To matches.count - 1
        '�R���N�V�����̌����[�v�I�u�W�F�N�g���擾
        Set match = matches.Item(i)
        '������v������
        ReDim Preserve returnVariant(i)
        returnVariant(i) = match.value
    Next
    LEGACY_findNumber = returnVariant
End Function

'==============================================================================================================================
'
'   ���s�R�[�h�݂̂����ւ���v���O����
'   �G�N�Z���ł͉��s�R�[�h�̎�ނ��ӊO�ɑ������ߊJ��
'
'   �߂�l :�@���s�R�[�h�������ꂽ������
'
'   text : ������
'   replaceText : ���s�R�[�h�Ǝ��ւ��镶����i�C�Ӂj
'
'==============================================================================================================================
Public Function LEGACY_ReplaceEnter(ByVal text As String, Optional ByVal replaceText As String = "") As String

    text = Replace(text, vbCr, replaceText)
    text = Replace(text, vbLf, replaceText)
    text = Replace(text, vbCrLf, replaceText)
    text = Replace(text, vbNewLine, replaceText)
    
    LEGACY_ReplaceEnter = text
    
End Function

'==============================================================================================================================
'
'   ������̒�����A����̕������O��𒊏o
'   �C���^�[�l�b�g�ɏ����ꂽ���ʂ肾�Ɠ���̕������Ȃ��ꍇ�G���[�ɂȂ�v���O�������~�܂��Ă��܂��̂Ŏ���
'
'   �߂�l : �����o���������A�G���[�̏ꍇ�͋�̔z�񂪕ԋp����܂��B
'
'   text : ������
'   deleteText : �폜���镶����
'
'
'==============================================================================================================================
Public Function LEGACY_LeftInStrString(ByVal text As String, ByVal deleteText As String) As String

    Dim r As String: r = text
    Dim i As Long: i = InStr(text, deleteText)
    If i >= 1 Then
        r = Left(text, i - 1)
    End If
    LEGACY_LeftInStrString = r
End Function

Public Function LEGACY_RigetInStrString(ByVal text As String, ByVal deleteText As String) As String

    Dim r As String: r = text
    Dim i As Long: i = InStr(text, deleteText)
    Dim l As Long: l = Len(deleteText)
    If i >= 1 And l >= 0 Then
        r = Mid(text, i + l)
    End If
    LEGACY_RigetInStrString = r
End Function
'==============================================================================================================================
'
'   ������̒��g���A���t�@�x�b�g�݂̂ō\������Ă��邩�H
'   https://vbabeginner.net/isalpha/
'
'   �߂�l : True(�A���t�@�x�b�g�̂�), False(����ȊO���܂�)
'
'==============================================================================================================================
Public Function LEGACY_IsAlphabets(text As String) As Boolean
    
    LEGACY_IsAlphabets = False
    If text = "" Then Exit Function
    
    LEGACY_IsAlphabets = Not text Like "*[!a-zA-Z��-���`-�y]*"
    
End Function
