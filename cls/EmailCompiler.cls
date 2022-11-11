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
'   EmailSendに外部データを流すためのクラス(送信とか表示とかはこちらではしない)
'==============================================================================================================================
Option Explicit
Dim DataCSV As New OpenCSV
Dim TemplateBody() As String
Dim TemplateBodyLength As Long
'------------------------------------------------------------------------------------------------------------------------------
'   コンストラクタとデストラクター
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
'   外部データを読み込む
'------------------------------------------------------------------------------------------------------------------------------
Public Function SetEmail(Mail As EmailSend, FileNameCSV As String) As Boolean
    
    If Not DataCSV.OpenFile(FileNameCSV, ",", "UTF-8") Then 'UTF-8形式
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
            'ここからデータを読むための処理に入る。(ほぼ同じコードが続きます)
            '--------------------------------------------------------------
            Dim ResultData As String 'サーチした結果を受け取る
            '差出人挿入
            If StrComp(Result, "<SENDER>", vbBinaryCompare) = 0 Then
                ResultData = SearchData(Y, X + 1, "</SENDER>")
                If ResultData = "" Then
                    MsgBox "差出人が不明です。"
                Else
                    Mail.SetSender (ResultData)
                End If
            End If
            '宛先挿入
            If StrComp(Result, "<DESTINATION>", vbBinaryCompare) = 0 Then
                ResultData = SearchData(Y, X + 1, "</DESTINATION>")
                If ResultData = "" Then
                    MsgBox "宛先が不明です。"
                Else
                    Mail.AddDestination (ResultData)
                End If
            End If
            '宛先挿入(CC)
            If StrComp(Result, "<DESTINATIONCC>", vbBinaryCompare) = 0 Then
                ResultData = SearchData(Y, X + 1, "</DESTINATIONCC>")
                If ResultData = "" Then
                    MsgBox "宛先(CC)が不明です。"
                Else
                    Mail.AddDestinationCC (ResultData)
                End If
            End If
            
            '宛先挿入(BCC)
            If StrComp(Result, "<DESTINATIONBCC>", vbBinaryCompare) = 0 Then
                ResultData = SearchData(Y, X + 1, "</DESTINATIONBCC>")
                If ResultData = "" Then
                    MsgBox "宛先(BCC)が不明です。"
                Else
                    Mail.AddDestinationBCC (ResultData)
                End If
            End If
            '件名挿入
            If StrComp(Result, "<SUBJECT>", vbBinaryCompare) = 0 Then
                ResultData = SearchData(Y, X + 1, "</SUBJECT>")
                If ResultData = "" Then
                    MsgBox "件名が不明です。"
                Else
                    Mail.SetSubjectText (ResultData)
                End If
            End If
            '添付ファイル挿入
            If StrComp(Result, "<ATTACHMENTS>", vbBinaryCompare) = 0 Then
                ResultData = SearchData(Y, X + 1, "</ATTACHMENTS>")
                If ResultData = "" Then
                    MsgBox "添付ファイルが不明です。"
                Else
                    Mail.AddAttachments (ResultData)
                End If
            End If
            '本文挿入(ここだけ少しややこしい)
            If StrComp(Result, "<BODY>", vbBinaryCompare) = 0 Then
                ResultData = SearchData(Y, X + 1, "</BODY>")
                TemplateBody = Split(ResultData, "<SendAutomatically>")
                If ResultData = "" Then
                    MsgBox "本文が不明です。"
                End If
            End If
        Next X
    Next Y
    
    SetEmail = True
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   AWのメール送信の仕方に乗らなければならないので本文を分けるポイントを作りました。
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
'特定の位置から同じ文字列をまでを探しその中身を返す
'------------------------------------------------------------------------------------------------------------------------------
Private Function SearchData(YLength As Long, XLength As Long, SearchText As String) As String
    Dim X As Long: X = 0
    Dim Y As Long: Y = 0
    Dim DataPool As String: DataPool = ""
    For Y = YLength To DataCSV.GetColumnLength()
        For X = 0 To DataCSV.GetRowLength(Y)
            '最初の一回だけ対策することがある
            If Y = YLength Then
                If XLength > DataCSV.GetRowLength(Y) Then Exit For '限界値を超えていればそれ以上は何もしない
                '最初だけ
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
                Exit Function '正常終了
            End If
            DataPool = DataPool + Result
        Next X
    Next Y
    SearchData = "" '無かったら空白
End Function
