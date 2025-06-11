VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpenBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'##############################################################################################################################
'
'   エクセルのBookを開く。XLSM,XLSX,XLS等、基本的なエクセルファイルが開けます
'   Bookを開く際のエラー等を潰す目的で開発した（大体は潰せたと思う）
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2023/02/09
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Office Professional Plus 2016 , 16.0.5.56.1000(32ビット)
'
'##############################################################################################################################
Private Type sheet
    work As Worksheet   'シートの全体情報
    Name As String      'シートの名前
    cordName As String  'シートのコード名
End Type

Dim book As Workbook     'ブックデータ
Dim bookName As String   'ブックデータの名前
Dim shareFlag As Boolean '共有モードの要否(要(True)、否(False))
Dim sheets() As sheet    'シート
'------------------------------------------------------------------------------------------------------------------------------
'   コンストラクタ
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set book = Nothing
    shareFlag = False
    ReDim sheets(0)
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'   デストラクタ
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Call CloseFile
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'   ファイルオープン
'   戻り値 : 開けた(True)開けなかった(False)
'
'   引数 :
'       fileNamePath : 開く目的のエクセルファイルのパス
'
'   オプション引数 :
'       PasswordOrKey : パスワード（鍵ファイルのみ）
'       UpdateLinks : リンクの更新方法を指定
'           値      外部参照        リモート参照
'           ------------------------------------
'            0      更新しない      更新しない
'            1      更新する        更新しない
'            2      更新しない      更新する
'            3      更新する        更新する
'       ReadOnlyMode : OpenのIgnoreReadOnlyRecommendedと同じ、読み取り専用で開く True:開く(初期値),False:開かない
'------------------------------------------------------------------------------------------------------------------------------
Public Function OpenFile(fileNamePath As String, Optional PasswordOrKey As String = "", Optional UpdateLinks As Long = 0, Optional ReadOnlyMode As Boolean = True) As Boolean
    'ファイルの存在チェック
    bookName = Dir(fileNamePath) 'ブックの名前を取得
    If Not bookName <> "" Then
        bookName = ""
        OpenFile = False
        Exit Function
    End If
    
    
    '同じブックを開いている場合は、そのブックを使用する。
    Dim BookFlag As Boolean: BookFlag = False
    For Each book In Workbooks
        If StrComp(book.Name, bookName, vbBinaryCompare) = 0 Then
            BookFlag = True
            Exit For
        End If
    Next
     
On Error GoTo ErrorHandler '下記でエラーが発生するため

    '更新警告等が利用趣旨と合わないため
    Application.DisplayAlerts = False

    If Not BookFlag Then
        'エクセル情報取得
        Set book = Nothing
        'パスワードが掛かっていた場合の処理
        If StrComp(PasswordOrKey, "", vbBinaryCompare) = 0 Then
            Set book = Workbooks.Open(fileNamePath, UpdateLinks, IgnoreReadOnlyRecommended:=ReadOnlyMode)
        Else
            Workbooks.Open filename:=fileNamePath, ReadOnly:=False, password:=PasswordOrKey, _
                           WriteResPassword:=PasswordOrKey, IgnoreReadOnlyRecommended:=ReadOnlyMode
            
            Application.Wait Now + TimeValue("0:00:03") '念のため開き終わるのを待つ
            Set book = Application.ActiveWorkbook
        End If
    End If
    '共有モードなのかを確認
    shareFlag = book.MultiUserEditing
    If shareFlag Then
        'ポップアップが表示されない場合もあるのでコメントで残す
        'Application.DisplayAlerts = False
        book.ExclusiveAccess
        'Application.DisplayAlerts = True
    End If
    
    'シート情報の取得
    Dim sheetLength As Long: sheetLength = book.Worksheets.Count
    ReDim sheets(sheetLength - 1)
    
    Dim i As Long
    For i = 0 To sheetLength - 1
        Set sheets(i).work = book.sheets(i + 1)
        sheets(i).Name = sheets(i).work.Name
        sheets(i).cordName = sheets(i).work.CodeName
    Next i
    
    Application.DisplayAlerts = True
    OpenFile = True
    Exit Function
ErrorHandler:
    Application.DisplayAlerts = True
    OpenFile = False
End Function

'------------------------------------------------------------------------------------------------------------------------------
'   ファイルクローズ
'
'   2019/01/16 追記
'       共有の場合は誰かが開いていると保存できずにエラーが発生した後、プログラムが停止します。
'       今は納品を優先するため放置します
'------------------------------------------------------------------------------------------------------------------------------
Public Function CloseFile(Optional overwrite As Boolean = False)
    '既にクローズ済み
    If book Is Nothing Then Exit Function
    
    'シート情報を解放
    If UBound(sheets) > 0 Then
        Dim i As Long
        For i = 0 To UBound(sheets)
            Set sheets(i).work = Nothing
        Next
    End If

    '共有が掛かっていた場合、かけ直す
    If shareFlag Then
        'Application.DisplayAlerts = False
        book.SaveAs filename:=book.FullName, accessMode:=xlShared
        book.Close (False)
        'Application.DisplayAlerts = True
    Else
        If overwrite Then
            Application.DisplayAlerts = False
            book.Save
            Application.DisplayAlerts = True
        End If
        book.Close (False)
    End If
        
    Set book = Nothing

End Function
'------------------------------------------------------------------------------------------------------------------------------
'シート名でシート索引を取得
'------------------------------------------------------------------------------------------------------------------------------
Private Function GetSheetIndex(Name As String) As Long
    Dim i As Long: i = 0
    GetSheetIndex = -&HFF
    For i = 0 To UBound(sheets)
        If StrComp(sheets(i).Name, Name, vbBinaryCompare) = 0 Then
            GetSheetIndex = i
            Exit Function
        End If
    Next i
End Function
'------------------------------------------------------------------------------------------------------------------------------
'
'   ゲッター
'
'------------------------------------------------------------------------------------------------------------------------------
Public Function GetBook() As Workbook
    Set GetBook = book
End Function
Public Function GetBookName() As String
    GetBookName = bookName
End Function

Public Function GetSheetLength() As Long
    GetSheetLength = UBound(sheets)
End Function

Public Function GetSheet(index As Long) As Worksheet
    Set GetSheet = sheets(index).work
End Function

Public Function GetSheet2(Name As String) As Worksheet
    Dim index As Long: index = GetSheetIndex(Name)
    If index = -&HFF Then
        Set GetSheetDatum2 = Nothing
        Exit Function
    End If
    Set GetSheetDatum2 = Me.GetSheet(index)
End Function

Public Function GetSheetNameDatum(index As Long) As String
    GetSheetNameDatum = sheets(index).Name
End Function


'そのブックは共有モードだったのか
Public Function GetShare() As Boolean
    GetShare = shareFlag
End Function
'ブックを閉じるとき共有モードで保存するか
Public Function SetShare(share As Boolean)
    shareFlag = share
End Function
