VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fileOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================================================================
'
'   �S�Ă͖����������̃e�L�X�g�t�@�C���͈�����N���X
'
'==============================================================================================================================
Public Text As String
'------------------------------------------------------------------------------------------------------------------------------
'   �R���X�g���N�^
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Text = ""
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'   �f�X�g���N�^
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Text = ""
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'   �t�@�C���ǂݍ���
'   �߂�l : ���̓ǂݍ��񂾃t�@�C���̕����� : �G���[�̏ꍇ�͋�
'   ����
'       fileName       : �t���p�X
'       characterCord  : �����R�[�h�w��(�C��) , �����l(Shift_JIS)
'------------------------------------------------------------------------------------------------------------------------------
Public Function readFile(ByVal fileName As String, Optional ByVal characterCord As String = "Shift_JIS") As Boolean
    readFile = False
    '�t�@�C�����݃`�F�b�N
    If Not Dir(fileName) <> "" Then Exit Function
    
    Dim Body As String
    '�G���[���
    On Error GoTo ErrorHandler
    With CreateObject("ADODB.Stream")
        .type = 2   'adTypeText
        .Charset = characterCord
        .Open
        .LoadFromFile (fileName)
        Body = .ReadText(-1)
        .Close
    End With
    
    Text = Body '�����ێ�
    readFile = True
    Exit Function
ErrorHandler:
    readFile = False
    Exit Function
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   �t�@�C����������
'   �߂�l : �������ݐ���(True),�������ݎ��s(False)
'   ����
'       fileName       : �t���p�X
'       characterCord  : �����R�[�h�w��(�C��) , �����l(Shift_JIS)
'       addFlag        : �t�@�C��������ꍇ�A�ǉ��ŏ������� , �����l(�������܂Ȃ�)
'------------------------------------------------------------------------------------------------------------------------------
Public Function writeFile(ByVal fileName As String, Optional ByVal characterCord As String = "Shift_JIS", Optional ByVal addFlag As Boolean = False) As Boolean
    writeFile = False
    '�������ރf�[�^�������ꍇ�B
    If StrComp(Text, "", vbBinaryCompare) = 0 Then Exit Function
    '�ǉ��ŏ������ނ��߂̊m�F����
    If addFlag Then
        If Not Dir(fileName) <> "" Then
            addFlag = False
        End If
    End If
    
    Dim Body As String: Body = ""
    '�G���[���
    On Error GoTo ErrorHandler
    With CreateObject("ADODB.Stream")
        .type = 2   'adTypeText
        .Charset = characterCord
        .Open
        If addFlag Then
            .LoadFromFile (fileName)
            Body = .ReadText(-1)
        End If
        .WriteText Body + Text
        .SaveToFile fileName, 2
        .Close
    End With
    
    writeFile = True
    Exit Function
ErrorHandler:
    writeFile = False
    Exit Function
End Function
