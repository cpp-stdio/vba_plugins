VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim FileLogUpdate As Boolean    '更新の要(True)否(False) : 初期設定(False)
Dim LogExtension  As String             'ログの拡張子            : 初期設定(.txt)

Dim MyPath As String                    '自身のパスを取得
Dim MyName As String                    '自身の名前を取得
Dim LogName As String                   'ログの名前(フルパス)
Dim Log As Object                       'ログファイルのハンドル
'------------------------------------------------------------------------------------------------------------------------------
'コンストラクタ
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    FileLogUpdate = False               '更新の要否
    LogExtension = ".txt"               '拡張子の設定

    MyPath = ActiveWorkbook.Path
    MyName = ActiveWorkbook.Name
    LogName = MyPath & "\" & MyName
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    LogName = FSO.GetBaseName(LogName)
    LogName = MyPath + "\" + LogName + LogExtension
    
    'ログファイルがなければ作ります
    If FSO.FileExists(LogName) = False Then
        FSO.CreateTextFile LogName
    Else
        '更新のため削除してから作り直します。
        If FileLogUpdate Then
            FSO.DeleteFile LogName
            FSO.CreateTextFile LogName
        End If
    End If
    'ログのファイルを作ります。
    Set Log = FSO.OpenTextFile(LogName, 8)
    Set FSO = Nothing
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'デストラクタ
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    '最後に閉じるのをお忘れなく。
    Set Log = Nothing
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'ログ出し関数
'------------------------------------------------------------------------------------------------------------------------------
Public Sub AddLog(ByVal Text As String)
    '日時＋タブ＋メッセージを書き込み
    Log.WriteLine Now & vbTab & Text
End Sub
