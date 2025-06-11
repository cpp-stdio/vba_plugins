Attribute VB_Name = "Involved_Book"
Option Explicit
'##############################################################################################################################
'
'   �u�b�N�֘A�̃}�N��
'   ���p����ꍇ�͉��L�̃C���|�[�g�����肢���܂�
'   �EInvolved_Other.bas
'
'   �V�K�쐬�� : 2017/08/30
'   �ŏI�X�V�� : 2024/01/30
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'   ���������O�̃V�[�g��T���B
'
'   �߂�l : ���������O�����V�[�g�B�Ȃ��ꍇ�́ANothing���ԋp�����
'
'   sheetName : �V�[�g��
'   book : �Ώۂ̃u�b�N�i�C�Ӂj
'==============================================================================================================================
Public Function LEGACY_BookToEqualsName(ByVal bookName As String) As Workbook
    Set LEGACY_BookToEqualsName = Nothing

    Dim book As Workbook
    For Each book In Workbooks
        If StrComp(book.Name, bookName, vbBinaryCompare) = 0 Then
            Set LEGACY_BookToEqualsName = book
            Exit Function
        End If
    Next
End Function

'==============================================================================================================================
'   �u�b�N�{�̂̃R�s�[���쐬����B
'
'   �߂�l : ���������O�����V�[�g�B�Ȃ��ꍇ�́ANothing���ԋp�����
'
'   book          : �Ώۂ̃u�b�N�i�C�Ӂj
'   filename      : �ۑ����i�󔒂̏ꍇ��book��+���ݎ����ɂȂ�A�g���q�s�v�j
'   pathname      : �p�X���i�󔒂̏ꍇ�͖{�̂̃u�b�N�Ɠ��K�w�ɂȂ�j
'==============================================================================================================================
Public Function LEGACY_aCopyBook(ByRef book As Workbook, Optional ByVal filename As String = "", Optional ByVal pathname As String = "") As Boolean
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If StrComp(filename, "", vbBinaryCompare) = 0 Then
        filename = Replace(book.Name, "." + fso.GetExtensionName(book.FullName), "") + "_" + LEGACY_CurrentTime()
    End If
    
    If StrComp(pathname, "", vbBinaryCompare) = 0 Then
        pathname = book.path
    End If
    
    Dim copyFullpath As String
    Dim extensionname As String: extensionname = fso.GetExtensionName(book.FullName)
    
'fso.CopyFile�Ńt�H���_�����݂��Ă��Ȃ��ƃG���[�ɂȂ邽��
On Error GoTo ErrorHandler_aCopyBook
    
    copyFullpath = pathname + "\" + filename + "." + extensionname
    fso.CopyFile book.FullName, copyFullpath
    
    Set fso = Nothing
    LEGACY_aCopyBook = True
    Exit Function
    
ErrorHandler_aCopyBook:
    Set fso = Nothing
    LEGACY_aCopyBook = False
End Function

