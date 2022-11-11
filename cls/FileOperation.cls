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
'   全ては無理だが大抵のテキストファイルは扱えるクラス
'
'==============================================================================================================================
Public Text As String
'------------------------------------------------------------------------------------------------------------------------------
'   コンストラクタ
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Text = ""
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'   デストラクタ
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Text = ""
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'   ファイル読み込み
'   戻り値 : その読み込んだファイルの文字列 : エラーの場合は空白
'   引数
'       fileName       : フルパス
'       characterCord  : 文字コード指定(任意) , 初期値(Shift_JIS)
'------------------------------------------------------------------------------------------------------------------------------
Public Function readFile(ByVal fileName As String, Optional ByVal characterCord As String = "Shift_JIS") As Boolean
    readFile = False
    'ファイル存在チェック
    If Not Dir(fileName) <> "" Then Exit Function
    
    Dim Body As String
    'エラー回避
    On Error GoTo ErrorHandler
    With CreateObject("ADODB.Stream")
        .type = 2   'adTypeText
        .Charset = characterCord
        .Open
        .LoadFromFile (fileName)
        Body = .ReadText(-1)
        .Close
    End With
    
    Text = Body '原文保持
    readFile = True
    Exit Function
ErrorHandler:
    readFile = False
    Exit Function
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   ファイル書き込み
'   戻り値 : 書き込み成功(True),書き込み失敗(False)
'   引数
'       fileName       : フルパス
'       characterCord  : 文字コード指定(任意) , 初期値(Shift_JIS)
'       addFlag        : ファイルがある場合、追加で書き込む , 初期値(書き込まない)
'------------------------------------------------------------------------------------------------------------------------------
Public Function writeFile(ByVal fileName As String, Optional ByVal characterCord As String = "Shift_JIS", Optional ByVal addFlag As Boolean = False) As Boolean
    writeFile = False
    '書き込むデータが無い場合。
    If StrComp(Text, "", vbBinaryCompare) = 0 Then Exit Function
    '追加で書き込むための確認事項
    If addFlag Then
        If Not Dir(fileName) <> "" Then
            addFlag = False
        End If
    End If
    
    Dim Body As String: Body = ""
    'エラー回避
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

