Attribute VB_Name = "Involved_Clipboard"
Option Explicit
'##############################################################################################################################
'
'   �N���b�v�{�[�h�֘A�iRPA����p�֐��j
'   �g�p����ɂ́A�uMicrosoft Forms 2.0 Object Library�v���Q�Ɛݒ肵�܂��B
'
'   �V�K�쐬�� : 2017/08/30
'   �ŏI�X�V�� : 2024/01/30
'
'   �V�K�쐬�G�N�Z���o�[�W���� : Office Professional Plus 2010 , 14.0.7145.5000(32�r�b�g)
'   �ŏI�X�V�G�N�Z���o�[�W���� : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'   �N���b�v�{�[�h�ɕ������ݒ肷��B
'
'   text : �N���b�v�{�[�h�ɃA�b�v����e�L�X�g���
'==============================================================================================================================
Public Function LEGACY_SetClipboard_Text(ByVal text As String)
    If StrComp(text, "", vbBinaryCompare) = 0 Then Exit Function
    With New MSForms.DataObject
        .SetText text
        .PutInClipboard
    End With
End Function

'==============================================================================================================================
'   �N���b�v�{�[�h���當������擾����B
'==============================================================================================================================
Public Function LEGACY_GetClipboard_Text() As String
    Dim text As String: text = ""
    With New MSForms.DataObject
        .GetFromClipboard
        text = .GetText
    End With
    GetText = text
End Function
