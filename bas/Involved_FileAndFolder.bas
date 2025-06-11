Attribute VB_Name = "Involved_FileAndFolder"
Option Explicit
'##############################################################################################################################
'
'   ファイル＆フォルダ関連
'   旧名 : Involved_File.bas
'   FolderHierarchyRead.clsはフォルダの階層読み込みに対応したプログラムなので併せてお使いください
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2025/06/12
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'   ファイル名を確認する。
'
'   戻り値 : OK(True), NG(False)
'
'   fileName : ファイル名
'==============================================================================================================================
Public Function LEGACY_checkFileName(ByVal fileName As String) As Boolean
    LEGACY_checkFileName = False
    '条件その1 : 空の名前ではない。
    If StrComp(fileName, "", vbBinaryCompare) = 0 Then Exit Function
    '条件その2 : 含んではいけない文字列がない。
    Dim textFor As Variant
    For Each textFor In Array("￥", "/", ":", "*", "?", """", "<", ">", "|")
        If InStr(fileName, CStr(textFor)) > 0 Then Exit Function
    Next textFor
    LEGACY_checkFileName = True
End Function

'==============================================================================================================================
'   ファイル読み込み、ある程度の文字コードに対応している。
'   戻り値 : その読み込んだファイルの文字列: エラーの場合は空白
'
'   fileName       : フルパス
'   characterCord  : 文字コード指定(任意) , 初期値(Shift_JIS),(非推奨：_autodetect_all)
'==============================================================================================================================
Public Function LEGACY_readFile(ByVal fileName As String, Optional ByVal characterCord As String = "Shift_JIS") As String
    LEGACY_readFile = ""
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

    LEGACY_readFile = Body '原文保持
    Exit Function
readFile_ErrorHandler:
    LEGACY_readFile = ""
    Exit Function
End Function
