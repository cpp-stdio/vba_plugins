VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim FileLogUpdate As Boolean    '�X�V�̗v(True)��(False) : �����ݒ�(False)
Dim LogExtension  As String             '���O�̊g���q            : �����ݒ�(.txt)

Dim MyPath As String                    '���g�̃p�X���擾
Dim MyName As String                    '���g�̖��O���擾
Dim LogName As String                   '���O�̖��O(�t���p�X)
Dim Log As Object                       '���O�t�@�C���̃n���h��
'------------------------------------------------------------------------------------------------------------------------------
'�R���X�g���N�^
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    FileLogUpdate = False               '�X�V�̗v��
    LogExtension = ".txt"               '�g���q�̐ݒ�

    MyPath = ActiveWorkbook.Path
    MyName = ActiveWorkbook.Name
    LogName = MyPath & "\" & MyName
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    LogName = FSO.GetBaseName(LogName)
    LogName = MyPath + "\" + LogName + LogExtension
    
    '���O�t�@�C�����Ȃ���΍��܂�
    If FSO.FileExists(LogName) = False Then
        FSO.CreateTextFile LogName
    Else
        '�X�V�̂��ߍ폜���Ă����蒼���܂��B
        If FileLogUpdate Then
            FSO.DeleteFile LogName
            FSO.CreateTextFile LogName
        End If
    End If
    '���O�̃t�@�C�������܂��B
    Set Log = FSO.OpenTextFile(LogName, 8)
    Set FSO = Nothing
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'�f�X�g���N�^
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    '�Ō�ɕ���̂����Y��Ȃ��B
    Set Log = Nothing
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'���O�o���֐�
'------------------------------------------------------------------------------------------------------------------------------
Public Sub AddLog(ByVal Text As String)
    '�����{�^�u�{���b�Z�[�W����������
    Log.WriteLine Now & vbTab & Text
End Sub