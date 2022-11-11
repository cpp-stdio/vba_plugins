Attribute VB_Name = "Involved_File"
Option Explicit
'##############################################################################################################################
'
'   フォイル関連
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2019/11/04
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'
'##############################################################################################################################

'==============================================================================================================================
'   ファイル名を確認する。
'
'   戻り値 : OK(True), NG(False)
'
'   fileName : ファイル名
'==============================================================================================================================
Public Function checkFileName(ByVal fileName As String) As Boolean
    checkFileName = False
    '条件その1 : 空の名前ではない。
    If StrComp(fileName, "", vbBinaryCompare) = 0 Then Exit Function
    '条件その2 : 含んではいけない文字列がない。
    Dim textFor As Variant
    For Each textFor In Array("￥", "/", ":", "*", "?", """", "<", ">", "|")
        If InStr(fileName, CStr(textFor)) > 0 Then Exit Function
    Next textFor
    checkFileName = True
End Function

'==============================================================================================================================
'   ファイル読み込み、ある程度の文字コードに対応している。
'   戻り値 : その読み込んだファイルの文字列: エラーの場合は空白
'
'   fileName       : フルパス
'   characterCord  : 文字コード指定(任意) , 初期値(Shift_JIS),(非推奨：_autodetect_all)
'==============================================================================================================================
Public Function readFile(ByVal fileName As String, Optional ByVal characterCord As String = "Shift_JIS") As String
    readFile = ""
    If Not Dir(fileName) <> "" Then Exit Function
    Dim Body As String

On Error GoTo readFile_ErrorHandler
    With CreateObject("ADODB.Stream")
        .type = 2   'adTypeText
        .Charset = characterCord
        .Open
        .LoadFromFile (fileName)
        Body = .ReadText(-1)
        .Close
    End With

    readFile = Body '原文保持
    Exit Function
readFile_ErrorHandler:
    readFile = ""
    Exit Function
End Function

'==============================================================================================================================
'   ファイル書き込み、ある程度の文字コードに対応している。
'   戻り値 : 成功(True),失敗(False)
'
'   text           : 保存用の文字列
'   fileName       : フルパス
'   characterCord  : 文字コード指定(任意) , 初期値(Shift_JIS)
'   addFlag        : ファイルがある場合、追加で書き込む , 初期値(書き込まない)
'==============================================================================================================================
' Public Function writeFile(ByRef text As String, ByVal fileName As String, Optional ByVal characterCord As String = "Shift_JIS", Optional ByVal addFlag As Boolean = False) As Boolean
'     writeFile = False
'     '書き込むデータが無い場合。
'     If StrComp(text, "", vbBinaryCompare) = 0 Then Exit Function
'     '追加で書き込むための確認事項
'     If addFlag Then
'         If Not Dir(fileName) <> "" Then
'             addFlag = False
'         End If
'     End If
'
'     Dim Body As String: Body = ""
' On Error GoTo writeFile_ErrorHandler
'     With CreateObject("ADODB.Stream")
'         .Type = 2   'adTypeText
'         .Charset = characterCord
'         .Open
'         If addFlag Then
'             .LoadFromFile (fileName)
'             Body = .ReadText(-1)
'         End If
'         .WriteText Body + text
'         .SaveToFile fileName, 2
'         .Close
'     End With
'
'     writeFile = True
'     Exit Function
' writeFile_ErrorHandler:
'     writeFile = False
'     Exit Function
' End Function
