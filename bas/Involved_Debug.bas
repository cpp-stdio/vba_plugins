Attribute VB_Name = "Involved_Debug"
Option Explicit
'##############################################################################################################################
'
'   �f�o�b�N���ɂ̂ݗL���Ȋ֐������݂���B
'   �f�o�b�N�p�Ȃ̂ŏ������Ԃ͍l�����Ă��Ȃ��A���������߂Ă���J���ł͕s�����ƌ����邩���m��Ȃ��B
'
'   �V�K�쐬�� : 2017/08/30
'   �ŏI�X�V�� : 2024/07/05
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

Private Enum atDevelopmentSwitching
    modeDebug   'Debug�������ƃG���[���\�����ꂽ����
    modeRelease '�����[�X���[�h�̏ꍇ�͂�����
End Enum

'�S�֐��ɗL���ȃt���O
Private Const atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug

'�e�֐������s�����邽�߂̃t���O�B�֐���ǉ������炱�������ǉ����邱�ƁB
Private Const atDevelopment_debugBox = atDevelopmentSwitching.modeDebug
Private Const atDevelopment_debugBookSave = atDevelopmentSwitching.modeDebug
Private Const atDevelopment_debugModuleImport = atDevelopmentSwitching.modeDebug
Private Const atDevelopment_debugModuleImportAll = atDevelopmentSwitching.modeDebug
Private Const atDevelopment_debugModuleExport = atDevelopmentSwitching.modeDebug
Private Const atDevelopment_debugModuleExportAll = atDevelopmentSwitching.modeDebug
'------------------------------------------------------------------------------------------------------------------------------
'   �f�o�b�N�p��MsgBox�B���񏑂��̂��ʓ|�Ȃ̂ō�����B
'   �����̐������߂�l�̐��������L���Q�ƁB�ꕔ�s�v�ȉӏ����������̂ŁA���������ȗ�
'
'   https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function
'------------------------------------------------------------------------------------------------------------------------------
Public Function LEGACY_debugBox(ByRef prompt As String, Optional ByVal button As VbMsgBoxStyle = vbOKOnly, Optional ByRef title As String = "Microsoft Excel") As VbMsgBoxResult
    LEGACY_debugBox = vbOK
    '�f�o�b�N���[�h�łȂ��Ƌ@�\���Ȃ��B
    If Not atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug Then Exit Function
    If Not atDevelopment_debugBox = atDevelopmentSwitching.modeDebug Then Exit Function
    
    LEGACY_debugBox = MsgBox(prompt, button, title)
End Function

'------------------------------------------------------------------------------------------------------------------------------
'   VBA��RAN�������u�Ԃɏ㏑���ۑ�����@�\���Ȃ��̂ŁA�Z�[�u���蓮�Ŏ�������B
'
'   book : �ۑ�������book���B�ݒ肵�Ȃ���ThisWorkbook���I������܂��B
'------------------------------------------------------------------------------------------------------------------------------
Public Function LEGACY_debugBookSave(Optional ByRef book As Workbook = Nothing)
    
    '�f�o�b�N���[�h�łȂ��Ƌ@�\���Ȃ��B
    If Not atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug Then Exit Function
    If Not atDevelopment_debugBookSave = atDevelopmentSwitching.modeDebug Then Exit Function
    
    Dim bookSave As Workbook
    If book Is Nothing Then
        Set bookSave = ThisWorkbook
    Else
        Set bookSave = book
    End If

    bookSave.Save
End Function

'==============================================================================================================================
'   �������W���[���C���|�[�g���G�N�X�|�[�g�Agit��svn���Ń\�[�X�Ǘ����������ꍇ�ɕ֗�
'
'   ���L�Q�lURL�� �Ƃ���Q�Ɛݒ�Ƀ`�F�b�N�����Ȃ���Γ��삵�Ȃ��������A
'   �`�F�b�N��t�����Ƃ��f�t�H���g�̏�Ԃœ����悤�ɂ���̂ɋ�J�����B
'
'   �Q�l�ɂ����C���|�[�g�v���O������
'   https://vbabeginner.net/%E6%A8%99%E6%BA%96%E3%83%A2%E3%82%B8%E3%83%A5%E3%83%BC%E3%83%AB%E7%AD%89%E3%81%AE%E4%B8%80%E6%8B%AC%E3%82%A4%E3%83%B3%E3%83%9D%E3%83%BC%E3%83%88/
'
'   �Q�l�ɂ����G�N�X�|�[�g�v���O������
'   https://vbabeginner.net/%E6%A8%99%E6%BA%96%E3%83%A2%E3%82%B8%E3%83%A5%E3%83%BC%E3%83%AB%E7%AD%89%E3%81%AE%E4%B8%80%E6%8B%AC%E3%82%A8%E3%82%AF%E3%82%B9%E3%83%9D%E3%83%BC%E3%83%88/
'
'   ��Excel�̐ݒ���ȉ��̒ʂ�ɕύX(�J���Ґ�p)
'     ���̐ݒ���s��Ȃ��ƁA�u���s���G���[ 1004 �v���O���~���O�ɂ�� visual basic �v���W�F�N�g�ւ̃A�N�Z�X�͐M�����Ɍ����܂��v
'     �ƃG���[���\������܂��B�K���s���悤�ɂ��ĉ������B
'     �t���O��modeRelease�ɕύX���邱�ƂŁA���̃G���[�͔������Ȃ��Ȃ�܂��B
'
'       �t�@�C�� -> �I�v�V���� -> �Z�L�����e�B�[�Z���^�[ -> [�Z�L�����e�B�[�Z���^�[�̐ݒ�]�{�^��������
'           �}�N���ݒ�i���y�C���j -> [VBA�v���W�F�N�g�I�u�W�F�N�g���f���ւ̃A�N�Z�X��M������]�@�`�F�b�NON
'
'==============================================================================================================================

'--------------------------------------------------------------
'   modulePaths : �C���|�[�g���郂�W���[���̃t���p�X�� : ��) C:\Users\Involved_Debug.bas
'   book        : �C���|�[�g����book���B�ݒ肵�Ȃ���ThisWorkbook���I������܂��B
'--------------------------------------------------------------
Public Function LEGACY_debugModuleImport(ByRef modulePaths() As String, Optional ByVal book As Workbook = Nothing)

    '�f�o�b�N���[�h�łȂ��Ƌ@�\���Ȃ��B
    If Not atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug Then Exit Function
    If Not atDevelopment_debugModuleImport = atDevelopmentSwitching.modeDebug Then Exit Function

    Dim extension  As String
    Dim Name       As String
    Dim textFor    As Variant
    Dim module     As Object '���W���[��
    Dim moduleList As Object 'VBA�v���W�F�N�g�̑S���W���[��
    
    Dim fso        As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim bookExport As Workbook
    If book Is Nothing Then
        Set bookExport = ThisWorkbook
    Else
        Set bookExport = book
    End If
    
    '�����Ώۃu�b�N�̃��W���[���ꗗ���擾
    Set moduleList = bookExport.VBProject.VBComponents
    
    'VBA�̎d�l�Ń��W���[�������t�@�C�����łȂ��ꍇ�����邪�Ή��o���Ȃ��ׁA�����ł͍l�����Ȃ��B
    For Each textFor In modulePaths

        '�g���q���������Ŏ擾
        extension = LCase(fso.GetExtensionName(textFor))

        '�p�X�����疼�O���擾
        Name = fso.GetBaseName(textFor)

        '�g���q�������ꂩ�̏ꍇ�A�C���|�[�g����B
        If StrComp(extension, "cls", vbBinaryCompare) = 0 Or _
            StrComp(extension, "frm", vbBinaryCompare) = 0 Or _
             StrComp(extension, "bas", vbBinaryCompare) = 0 Then
            
            For Each module In moduleList
                If StrComp(Name, module.Name, vbBinaryCompare) = 0 Then
                    '�����̃��W���[���폜
                    Call moduleList.Remove(module)
                    Exit For
                End If
            Next
            '���W���[����ǉ�
            Call moduleList.Import(textFor)
        End If
    Next

End Function

'--------------------------------------------------------------
'   modulePath  : �C���|�[�g���郂�W���[���̃t���p�X��
'   book        : �C���|�[�g����book���B�ݒ肵�Ȃ���ThisWorkbook���I������܂��B
'--------------------------------------------------------------
Public Function LEGACY_debugModuleImportAll(ByRef modulePath As String, Optional ByVal book As Workbook = Nothing)

    '�f�o�b�N���[�h�łȂ��Ƌ@�\���Ȃ��B
    If Not atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug Then Exit Function
    If Not atDevelopment_debugModuleImportAll = atDevelopmentSwitching.modeDebug Then Exit Function
    
    '�K�w�ǂݍ��݂��g�p���A�w�肳�ꂽ�p�X���̑S�Ẵv���O�������擾����
    Dim fhr As FolderHierarchyRead
    Set fhr = New FolderHierarchyRead
    fhr.Search (modulePath)

    Call LEGACY_debugModuleImport(fhr.getFileData(), book)
    
    Set fhr = Nothing

End Function

'--------------------------------------------------------------
'   modules  : �G�N�X�|�[�g���������W���[���� : ��) Array("Involved_Debug")
'   book     : �G�N�X�|�[�g������book���B�ݒ肵�Ȃ���ThisWorkbook���I������܂��B
'   filePath : �G�N�X�|�[�g�����t�H���_����w�肷��B�w�肪�Ȃ���book�̃o�X���I������܂��B
'--------------------------------------------------------------
Public Function LEGACY_debugModuleExport(ByRef modules() As String, Optional ByVal book As Workbook = Nothing, Optional ByVal folderPath As String = "")

    '�f�o�b�N���[�h�łȂ��Ƌ@�\���Ȃ��B
    If Not atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug Then Exit Function
    If Not atDevelopment_debugModuleExport = atDevelopmentSwitching.modeDebug Then Exit Function
    
    'module.Type�̓N���X���ɏ����ꂽEnum�ł���A�A�N�Z�X�s�ׁ̈A�ÓI�ϐ��ő�p����B
    Static vbext_ct_StdModule As Long: vbext_ct_StdModule = 1
    Static vbext_ct_MSForm As Long: vbext_ct_MSForm = 2
    Static vbext_ct_ClassModule As Long: vbext_ct_ClassModule = 3
    
    Dim module     As Object '���W���[��
    Dim moduleList As Object 'VBA�v���W�F�N�g�̑S���W���[��
    Dim extension  As String  '���W���[���̊g���q
    Dim textFor    As Variant
    
    Dim bookExport As Workbook
    If book Is Nothing Then
        Set bookExport = ThisWorkbook
    Else
        Set bookExport = book
    End If
    
    '�����Ώۃu�b�N�̃��W���[���ꗗ���擾
    Set moduleList = bookExport.VBProject.VBComponents
    
    '�ۑ���̎w�肪�Ȃ��̂�bookExport�Ɠ��K�w�ɃG�N�X�|�[�g����B
    If StrComp(folderPath, "", vbBinaryCompare) = 0 Then
        folderPath = bookExport.path
    End If
    
    For Each module In moduleList
        extension = ""
        '�g���q���w�肷��B
        If (module.type = vbext_ct_ClassModule) Then
            extension = ".cls" '�N���X
        ElseIf (module.type = vbext_ct_MSForm) Then
            extension = ".frm" '�t�H�[��(.frx���ꏏ�ɃG�N�X�|�[�g�����)
        ElseIf (module.type = vbext_ct_StdModule) Then
            extension = ".bas" '�W�����W���[��
        End If

        '�G�N�X�|�[�g
        If Not StrComp(extension, "", vbBinaryCompare) = 0 Then
            For Each textFor In modules
                '�z��̒��ɑ��݂��Ă���΁A�G�N�X�|�[�g����B
                If StrComp(textFor, module.Name, vbBinaryCompare) = 0 Then
                    Call module.Export(folderPath + "\" + module.Name + extension)
                End If
            Next
        End If
    Next
End Function

'--------------------------------------------------------------
'   book     : �G�N�X�|�[�g������book���B�ݒ肵�Ȃ���ThisWorkbook���I������܂��B
'   filePath : �G�N�X�|�[�g�����t�H���_����w�肷��B�w�肪�Ȃ���book�̃o�X���I������܂��B
'--------------------------------------------------------------
Public Function LEGACY_debugModuleExportAll(Optional ByVal folderPath As String = "", Optional ByVal book As Workbook = Nothing)

    '�f�o�b�N���[�h�łȂ��Ƌ@�\���Ȃ��B
    If Not atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug Then Exit Function
    If Not atDevelopment_debugModuleExportAll = atDevelopmentSwitching.modeDebug Then Exit Function
    
    Dim module     As Object '���W���[��
    Dim moduleList As Object 'VBA�v���W�F�N�g�̑S���W���[��
    Dim names() As String
    Dim length As Long: length = -1

    Dim bookExport As Workbook
    If book Is Nothing Then
        Set bookExport = ThisWorkbook
    Else
        Set bookExport = book
    End If
    
    '�����Ώۃu�b�N�̃��W���[���ꗗ���擾
    Set moduleList = bookExport.VBProject.VBComponents
    For Each module In moduleList
        length = length + 1
        ReDim Preserve names(length)
        names(length) = module.Name
    Next

    '�ۑ�����
    Call LEGACY_debugModuleExport(names, bookExport, folderPath)
End Function
