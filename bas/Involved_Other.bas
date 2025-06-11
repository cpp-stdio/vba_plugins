Attribute VB_Name = "Involved_Other"
Option Explicit
'##############################################################################################################################
'
'   ���̑��A�悭�W�����������s�\�Ȋ֐��Q
'
'   �V�K�쐬�� : 2017/08/30
'   �ŏI�X�V�� : 2024/01/26
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'   ���ݎ�����Ԃ�
'
'   T_Flag :  0,�N����b�܂ł̂��ׂĂ̎���  (��)2018.04.23.03.54.02
'             1,�N������܂ł̓��t         (��)2018.04.23
'             2,������b�܂ł̎���         (��)03.54.02
'             3,�N�̂�                    (��)2018
'             4,���̂�                    (��)04
'             5,���̂�                    (��)23
'             6,���̂�                    (��)03
'             7,���̂�                    (��)54
'             8,�b�̂�                    (��)02
'           �@����ȊO                    �S��"0"�Ƃ��ď�������
'
'   ToBe   : �Ԃɓ���Ăق���������u2018/04/23�v�u2018.04.23�v�� T_Flag�̒l��0�`2�̎��̂ݗL��
'==============================================================================================================================
Public Function LEGACY_CurrentTime(Optional ByVal T_Flag As Long = 0, Optional ByVal ToBe As String = ".") As String

    Dim NowYear() As String
    Dim NowTime() As String
    NowYear = Split(Format(Date, "yyyy:mm:dd"), ":")
    NowTime = Split(Format(Time, "hh:mm:ss"), ":")

    Select Case T_Flag
        Case 1      '�N������܂ł̓��t
            LEGACY_CurrentTime = NowYear(0) + ToBe + NowYear(1) + ToBe + NowYear(2)
        Case 2      '������b�܂ł̎���
            LEGACY_CurrentTime = NowTime(0) + ToBe + NowTime(1) + ToBe + NowTime(2)
        Case 3      '�N�̂�
            LEGACY_CurrentTime = NowYear(0)
        Case 4      '���̂�
            LEGACY_CurrentTime = NowYear(1)
        Case 5      '���̂�
            LEGACY_CurrentTime = NowYear(2)
        Case 6      '���̂�
            LEGACY_CurrentTime = NowTime(0)
        Case 7      '���̂�
            LEGACY_CurrentTime = NowTime(1)
        Case 8      '�b�̂�
            LEGACY_CurrentTime = NowTime(2)
        Case Else   '0���܂߁A����ȊO
            LEGACY_CurrentTime = NowYear(0) + ToBe + NowYear(1) + ToBe + NowYear(2) + ToBe + NowTime(0) + ToBe + NowTime(1) + ToBe + NowTime(2)
    End Select
End Function
