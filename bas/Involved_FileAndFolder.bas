Attribute VB_Name = "Involved_FileAndFolder"
Option Explicit
'##############################################################################################################################
'
'   �t�H�C���֘A
'
'   �V�K�쐬�� : 2017/08/30
'   �ŏI�X�V�� : 2024/01/30
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'   �t�@�C�������m�F����B
'
'   �߂�l : OK(True), NG(False)
'
'   fileName : �t�@�C����
'==============================================================================================================================
Public Function LEGACY_checkFileName(ByVal fileName As String) As Boolean
    LEGACY_checkFileName = False
    '��������1 : ��̖��O�ł͂Ȃ��B
    If StrComp(fileName, "", vbBinaryCompare) = 0 Then Exit Function
    '��������2 : �܂�ł͂����Ȃ������񂪂Ȃ��B
    Dim textFor As Variant
    For Each textFor In Array("��", "/", ":", "*", "?", """", "<", ">", "|")
        If InStr(fileName, CStr(textFor)) > 0 Then Exit Function
    Next textFor
    LEGACY_checkFileName = True
End Function

'==============================================================================================================================
'   �t�@�C���ǂݍ��݁A������x�̕����R�[�h�ɑΉ����Ă���B
'   �߂�l : ���̓ǂݍ��񂾃t�@�C���̕�����: �G���[�̏ꍇ�͋�
'
'   fileName       : �t���p�X
'   characterCord  : �����R�[�h�w��(�C��) , �����l(Shift_JIS),(�񐄏��F_autodetect_all)
'==============================================================================================================================
Public Function LEGACY_readFile(ByVal fileName As String, Optional ByVal characterCord As String = "Shift_JIS") As String
    LEGACY_readFile = ""
    If Not Dir(fileName) <> "" Then Exit Function
    Dim Body As String

On Error GoTo readFile_ErrorHandler
    With CreateObject("ADODB.Stream")
        .type = 2   'adTypeText
        .Charset = characterCord
        .Open
        .LoadFromFile (fileName)
        Body = .ReadText(-1)
        .Close
    End With

    LEGACY_readFile = Body '�����ێ�
    Exit Function
readFile_ErrorHandler:
    LEGACY_readFile = ""
    Exit Function
End Function
