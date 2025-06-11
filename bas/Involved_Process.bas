Attribute VB_Name = "Involved_Process"
Option Explicit
'##############################################################################################################################
'
'   VBA�𓮂����ۂɕK�v�s���ȏ���
'
'   �V�K�쐬�� : 2017/08/30
'   �ŏI�X�V�� : 2024/07/05
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'�v���O�������s���Ԃ��J�n�v���p�̕ϐ�
Private beginTime As Date

'==============================================================================================================================
'
'   VBA�̃V�[�g�X�V���̏������y��������B
'   ���̊֐����ĂԂ����ŁA�������Ԃ�8.5�{�قǌ��シ��B
'
'   �Q�lURL
'   https://tonari-it.com/vba-processing-speed/
'
'==============================================================================================================================
Public Function LEGACY_reduceProcess_ToBegin()

    '���s���Ԍv���J�n
    beginTime = Time
    
    Application.Calculation = xlCalculationManual '�v�Z���[�h���}�j���A���ɂ���
    Application.EnableEvents = False              '�C�x���g���~������
    Application.ScreenUpdating = False            '��ʕ\���X�V���~������
    
        'Application.Cursor = xlWait                   '�}�E�X�|�C���^�̏�Ԃ������v�^�ɕύX

End Function

Public Function LEGACY_reduceProcess_ToEnd()
    
    Application.Calculation = xlCalculationAutomatic '�v�Z���[�h�������ɂ���
    Application.EnableEvents = True                  '�C�x���g���J�n������
    Application.ScreenUpdating = True                '��ʕ\���X�V���J�n������
    
        'Application.Cursor = xlDefault                   '�}�E�X�|�C���^�̏�Ԃ�W���^�ɕύX
    
    '���s���Ԍv���I��
    Application.StatusBar = "�������� / ���s���Ԃ� " + Format(Time - beginTime, "nn��ss�b") + " �ł���"
    
End Function

'==============================================================================================================================
'   100%�\���ŕ\���o�����[�^�[��ǉ����X�e�[�^�X�o�[�ɒǉ�����
'
'   �߂�l : OK(True), NG(False)
'
'   message : �X�e�[�^�X�o�[�ɏ����������b�Z�[�W
'   now     : ���݂̒l(for�������ɂ��g��������)
'   max     : �S�̐�
'==============================================================================================================================
Public Function LEGACY_StatusBar_100barometer(message As String, now As Long, max As Long)

    LEGACY_StatusBar_100barometer = False

    '0���Z�΍�
    If max <= 0 Then Exit Function
    
    Dim text As String
    text = message + " ( " + CStr(CLng(now / max * 100)) + "% : " + CStr(now) + "/" + CStr(max) + ") "
    
    '100%�����́��⁠�̕\�������������Ȃ邽��
    If now <= max Then
        text = text + String(CLng(now / max * 10), "��") + String(10 - CLng(now / max * 10), "��")
    End If
    
    Application.StatusBar = text
    LEGACY_StatusBar_100barometer = True

End Function
