VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FolderHierarchyRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================================================================
'
'   フォルダの階層を読み込むためのクラス
'   作成日 : 2017/08/21
'
'==============================================================================================================================
Dim HierarchyLength As Long          '階層ファイル個数(再起処理のためにいるもの)
Dim HierarchyFolderLength As Long    '階層フォルダ個数(再起処理のためにいるもの)
Dim Hierarchical() As String         '階層ファイル格納のための一次元配列
Dim HierarchicalFolder() As String   '階層フォルダ格納のための一次元配列
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Call CloseData '初期化の代用
End Sub
Private Sub Class_Terminate()
    Call CloseData
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'   フォルダの階層を読み込む
'   DeepHierarchyFlag : さらに深い階層を取得する？ : 要(True).否(False),初期値(False)
'   FullPath          : フルパスで取得する？       : 要(True).否(False),初期値(True)
'------------------------------------------------------------------------------------------------------------------------------
Public Sub Search(folderName As String, Optional DeepHierarchyFlag As Boolean = False, Optional FullPath As Boolean = True)
    '階層読み込み
    Dim Buf As String
    Buf = Dir(folderName & "\*.*")
    'ファイル名を取得
    Do While Buf <> ""
        ReDim Preserve Hierarchical(HierarchyLength)
        If FullPath Then
            Hierarchical(HierarchyLength) = folderName + "\" + Buf
        Else
            Hierarchical(HierarchyLength) = Buf
        End If
        HierarchyLength = HierarchyLength + 1
        Buf = Dir()
    Loop
    
    Dim f As Object
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .getFolder(folderName).SubFolders
            'フォルダの名前を取得の名前
            ReDim Preserve HierarchicalFolder(HierarchyFolderLength)
            HierarchicalFolder(HierarchyFolderLength) = f.path
            HierarchyFolderLength = HierarchyFolderLength + 1
            'さらに深い階層に潜る必要がある場合、再起処理開始
            If DeepHierarchyFlag Then
                Call Search(f.path, DeepHierarchyFlag, FullPath)
            End If
        Next f
    End With
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'   特定の拡張子に絞り込む ↑のSearchを先に実行してないと意味ないよ。(やる前にClone関数呼ぶのをおすすめする)
'   ExtensionName : 例.テキスト(.txt)の場合"txt"と入力して。最初のドットはいらない
'------------------------------------------------------------------------------------------------------------------------------
Public Function ExtensionSearch(ParamArray extensions() As Variant) As Boolean
    ExtensionSearch = False
    'データがない場合は処理できない
    If HierarchyLength = 0 Then
        Exit Function
    End If
    
    Dim i As Long:
    Dim e1 As Variant
    Dim e2 As Variant
    Dim count As Long: count = 0
    Dim Temporary() As String
    
    For i = 0 To HierarchyLength - 1
        '一番最後が拡張子なので
        e2 = Split(Hierarchical(i), ".")
        For Each e1 In extensions
            'OSの仕様上、大文字小文字は同じ拡張子
            If StrComp(UCase(CStr(e1)), UCase(e2(UBound(e2))), vbBinaryCompare) = 0 Then
                ReDim Preserve Temporary(count)
                Temporary(count) = Hierarchical(i)
                count = count + 1
            End If
        Next e1
    Next i
    'その拡張子がない場合は再構築はしない
    If count = 0 Then Exit Function
    
    'データの再構築
    HierarchyLength = count
    ReDim Hierarchical(HierarchyLength - 1)
    Hierarchical = Temporary
    ExtensionSearch = True
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   指定ファイルの下の階層のファイル、フォルダをすべて削除する
'------------------------------------------------------------------------------------------------------------------------------
Public Function DeleteAll(folderName As String)
    
    Dim Buf As String
    Buf = Dir(folderName & "\*.*")
    'ファイルを削除
    Do While Buf <> ""
        Kill folderName + "\" + Buf
        Buf = Dir()
    Loop
    
    Dim f As Object
    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    'フォルダを削除
    For Each f In fileSystem.getFolder(folderName).SubFolders
            fileSystem.DeleteFolder f.path, True
    Next f
    
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   ファイルゲッター
'------------------------------------------------------------------------------------------------------------------------------
Public Function getFileLength() As Long
    getFileLength = HierarchyLength - 1
End Function
Public Function getFileDatum(Index As Long) As String
    getFileDatum = Hierarchical(Index)
End Function
Public Function getFileData() As String()
    getFileData = Hierarchical
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   フォルダーゲッター
'------------------------------------------------------------------------------------------------------------------------------
Public Function getFolderLength() As Long
    getFolderLength = HierarchyFolderLength - 1
End Function
Public Function getFolderDatum(Index As Long) As String
    getFolderDatum = HierarchicalFolder(Index)
End Function
Public Function getFolderData() As String()
    getFolderData = HierarchicalFolder
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   クローン生成
'------------------------------------------------------------------------------------------------------------------------------
Public Function Clone(Original As FolderHierarchyRead)
    '必要な要素の受け渡し
    HierarchyLength = Original.GetHierarchyLength() + 1
    HierarchyFolderLength = Original.GetHierarchyFolderLength() + 1
    ReDim Hierarchical(HierarchyLength - 1)
    Hierarchical = Original.GetHierarchicalData()
    ReDim HierarchicalFolder(HierarchyFolderLength - 1)
    HierarchicalFolder = GetHierarchicalFolderData()
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   クローズ
'------------------------------------------------------------------------------------------------------------------------------
Public Function CloseData()
    HierarchyLength = 0
    HierarchyFolderLength = 0
    ReDim Hierarchical(HierarchyLength)
    ReDim HierarchicalFolder(HierarchyFolderLength)
End Function

