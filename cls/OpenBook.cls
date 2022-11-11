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
'------------------------------------------------------------------------------------------------------------------------------
'
'   XLSM,XLSX,XLS等、基本的なエクセルファイルが開けます。
'   最終確認 : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'
'------------------------------------------------------------------------------------------------------------------------------
Private Type sheet
    work As Worksheet   'シートの全体情報
    name As String      'シートの名前
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
'   ファイルオープン : 開けた(True)開けなかった(False)
'------------------------------------------------------------------------------------------------------------------------------
Public Function OpenFile(fileNamePath As String, Optional PasswordOrKey As String = "") As Boolean
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
        If StrComp(book.name, bookName, vbBinaryCompare) = 0 Then
            BookFlag = True
            Exit For
        End If
    Next
     
On Error GoTo ErrorHandler '下記でエラーが発生するため

    If Not BookFlag Then
        'エクセル情報取得
        Set book = Nothing
        'パスワードが掛かっていた場合の処理
        If StrComp(PasswordOrKey, "", vbBinaryCompare) = 0 Then
            Set book = Workbooks.Open(fileNamePath, 3)
        Else
            Workbooks.Open fileName:=fileNamePath, ReadOnly:=False, password:=PasswordOrKey, _
                           WriteResPassword:=PasswordOrKey, IgnoreReadOnlyRecommended:=True
            
            Application.Wait Now + TimeValue("0:00:03") '念のため開き終わるのを待つ
            Set book = Application.ActiveWorkbook
        End If
    End If
    '共有モードなのかを確認
    shareFlag = book.MultiUserEditing
    If shareFlag Then
        'Application.DisplayAlerts = False
        book.ExclusiveAccess
        'Application.DisplayAlerts = True
    End If
    
    'シート情報の取得
    Dim sheetLength As Long: sheetLength = book.Worksheets.count
    ReDim sheets(sheetLength - 1)
    
    Dim i As Long
    For i = 0 To sheetLength - 1
        Set sheets(i).work = book.sheets(i + 1)
        sheets(i).name = sheets(i).work.name
        sheets(i).cordName = sheets(i).work.CodeName
    Next i
    OpenFile = True
    Exit Function
ErrorHandler:
    OpenFile = False
End Function

'------------------------------------------------------------------------------------------------------------------------------
'   ファイルクローズ
'
'   2019/01/16 追記
'   共有の場合は誰かが開いていると保存できずにエラーが発生した後、プログラムが停止します。
'   今は納品が先なので、このまま放置します。
'------------------------------------------------------------------------------------------------------------------------------
Public Function CloseFile()
    'シート情報を解放
    If UBound(sheets) > 0 Then
        Dim i As Long
        For i = 0 To UBound(sheets)
            Set sheets(i).work = Nothing
        Next
    End If
    'ブック情報を解放
    If Not book Is Nothing Then
        '共有が掛かっていた場合、かけ直す
        If shareFlag Then
            'Application.DisplayAlerts = False
            book.SaveAs fileName:=book.FullName, accessMode:=xlShared
            book.Close (False)
            'Application.DisplayAlerts = True
        Else
            book.Close (False)
        End If
        
        Set book = Nothing
    End If
End Function
'------------------------------------------------------------------------------------------------------------------------------
'シート名でシート索引を取得
'------------------------------------------------------------------------------------------------------------------------------
Private Function GetSheetIndex(name As String) As Long
    Dim i As Long: i = 0
    GetSheetIndex = -&HFF
    For i = 0 To UBound(sheets)
        If StrComp(sheets(i).name, name, vbBinaryCompare) = 0 Then
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

Public Function GetSheet(Index As Long) As Worksheet
    Set GetSheetDatum = sheets(Index).work
End Function

Public Function GetSheet2(name As String) As Worksheet
    Dim Index As Long: Index = GetSheetIndex(name)
    If Index = -&HFF Then
        Set GetSheetDatum2 = Nothing
        Exit Function
    End If
    Set GetSheetDatum2 = Me.GetSheet(Index)
End Function

Public Function GetSheetNameDatum(Index As Long) As String
    GetSheetNameDatum = sheets(Index).name
End Function


'そのブックは共有モードだったのか
Public Function GetShare() As Boolean
    GetShare = shareFlag
End Function
'ブックを閉じるとき共有モードで保存するか
Public Function SetShare(share As Boolean)
    shareFlag = share
End Function
