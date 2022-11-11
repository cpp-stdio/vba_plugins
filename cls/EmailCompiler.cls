VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmailCompiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================================================
'   EmailSend�ɊO���f�[�^�𗬂����߂̃N���X(���M�Ƃ��\���Ƃ��͂�����ł͂��Ȃ�)
'==============================================================================================================================
Option Explicit
Dim DataCSV As New OpenCSV
Dim TemplateBody() As String
Dim TemplateBodyLength As Long
'------------------------------------------------------------------------------------------------------------------------------
'   �R���X�g���N�^�ƃf�X�g���N�^�[
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set DataCSV = New OpenCSV
    TemplateBodyLength = 0
    ReDim TemplateBody(TemplateBodyLength)
End Sub
Private Sub Class_Terminate()
    Set DataCSV = Nothing
    TemplateBodyLength = 0
    ReDim TemplateBody(TemplateBodyLength)
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'   �O���f�[�^��ǂݍ���
'------------------------------------------------------------------------------------------------------------------------------
Public Function SetEmail(Mail As EmailSend, FileNameCSV As String) As Boolean
    
    If Not DataCSV.OpenFile(FileNameCSV, ",", "UTF-8") Then 'UTF-8�`��
        SetEmail = False
        Exit Function
    End If
    
    Dim Analysis() As String
    Dim AnalysisLength As Long: AnalysisLength = 0
    ReDim Analysis(AnalysisLength)
    
    Dim X As Long: X = 0
    Dim Y As Long: Y = 0
    For Y = 0 To DataCSV.GetColumnLength()
        For X = 0 To DataCSV.GetRowLength(Y)
            Dim Result As String
            Result = DataCSV.GetRowDatum(Y, X)
            Result = UCase(Result)
            '--------------------------------------------------------------
            '��������f�[�^��ǂނ��߂̏����ɓ���B(�قړ����R�[�h�������܂�)
            '--------------------------------------------------------------
            Dim ResultData As String '�T�[�`�������ʂ��󂯎��
            '���o�l�}��
            If StrComp(Result, "<SENDER>", vbBinaryCompare) = 0 Then
                ResultData = SearchData(Y, X + 1, "</SENDER>")
                If ResultData = "" Then
                    MsgBox "���o�l���s���ł��B"
                Else
                    Mail.SetSender (ResultData)
                End If
            End If
            '����}��
            If StrComp(Result, "<DESTINATION>", vbBinaryCompare) = 0 Then
                ResultData = SearchData(Y, X + 1, "</DESTINATION>")
                If ResultData = "" Then
                    MsgBox "���悪�s���ł��B"
                Else
                    Mail.AddDestination (ResultData)
                End If
            End If
            '����}��(CC)
            If StrComp(Result, "<DESTINATIONCC>", vbBinaryCompare) = 0 Then
                ResultData = SearchData(Y, X + 1, "</DESTINATIONCC>")
                If ResultData = "" Then
                    MsgBox "����(CC)���s���ł��B"
                Else
                    Mail.AddDestinationCC (ResultData)
                End If
            End If
            
            '����}��(BCC)
            If StrComp(Result, "<DESTINATIONBCC>", vbBinaryCompare) = 0 Then
                ResultData = SearchData(Y, X + 1, "</DESTINATIONBCC>")
                If ResultData = "" Then
                    MsgBox "����(BCC)���s���ł��B"
                Else
                    Mail.AddDestinationBCC (ResultData)
                End If
            End If
            '�����}��
            If StrComp(Result, "<SUBJECT>", vbBinaryCompare) = 0 Then
                ResultData = SearchData(Y, X + 1, "</SUBJECT>")
                If ResultData = "" Then
                    MsgBox "�������s���ł��B"
                Else
                    Mail.SetSubjectText (ResultData)
                End If
            End If
            '�Y�t�t�@�C���}��
            If StrComp(Result, "<ATTACHMENTS>", vbBinaryCompare) = 0 Then
                ResultData = SearchData(Y, X + 1, "</ATTACHMENTS>")
                If ResultData = "" Then
                    MsgBox "�Y�t�t�@�C�����s���ł��B"
                Else
                    Mail.AddAttachments (ResultData)
                End If
            End If
            '�{���}��(��������������₱����)
            If StrComp(Result, "<BODY>", vbBinaryCompare) = 0 Then
                ResultData = SearchData(Y, X + 1, "</BODY>")
                TemplateBody = Split(ResultData, "<SendAutomatically>")
                If ResultData = "" Then
                    MsgBox "�{�����s���ł��B"
                End If
            End If
        Next X
    Next Y
    
    SetEmail = True
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   AW�̃��[�����M�̎d���ɏ��Ȃ���΂Ȃ�Ȃ��̂Ŗ{���𕪂���|�C���g�����܂����B
'------------------------------------------------------------------------------------------------------------------------------
Public Function GetTemplateBody() As String
    If TemplateBodyLength > UBound(TemplateBody) Then
        GetTemplateBody = ""
        Exit Function
    End If
    GetTemplateBody = TemplateBody(TemplateBodyLength)
    TemplateBodyLength = TemplateBodyLength + 1
End Function
'------------------------------------------------------------------------------------------------------------------------------
'����̈ʒu���瓯����������܂ł�T�����̒��g��Ԃ�
'------------------------------------------------------------------------------------------------------------------------------
Private Function SearchData(YLength As Long, XLength As Long, SearchText As String) As String
    Dim X As Long: X = 0
    Dim Y As Long: Y = 0
    Dim DataPool As String: DataPool = ""
    For Y = YLength To DataCSV.GetColumnLength()
        For X = 0 To DataCSV.GetRowLength(Y)
            '�ŏ��̈�񂾂��΍􂷂邱�Ƃ�����
            If Y = YLength Then
                If XLength > DataCSV.GetRowLength(Y) Then Exit For '���E�l�𒴂��Ă���΂���ȏ�͉������Ȃ�
                '�ŏ�����
                If X = 0 Then
                    X = XLength
                End If
            End If
            Dim Result As String
            Result = DataCSV.GetRowDatum(Y, X)
            Dim ResultData As String
            ResultData = UCase(Result)
            If StrComp(ResultData, SearchText, vbBinaryCompare) = 0 Then
                SearchData = DataPool
                Exit Function '����I��
            End If
            DataPool = DataPool + Result
        Next X
    Next Y
    SearchData = "" '�����������
End Function