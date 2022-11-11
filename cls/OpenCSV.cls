VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpenCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type DataCSV
    Index As Long
    LineData() As String
End Type

Dim Data() As DataCSV
Dim Row As Long, Column As Long 'キーの位置からデータを取得するタイプ用
'------------------------------------------------------------------------------------------------------------------------------
'   コンストラクタ
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    ReDim Data(0)
    Row = 0
    Column = 0
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'   デストラクタ
'------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    ReDim Data(0)
    Row = 0
    Column = 0
End Sub
'------------------------------------------------------------------------------------------------------------------------------
'   CSVのデータを開く : 開けた(True)開けなかった(False)
'
'   CSV_Name            : ファイル名(フルパス)
'   Separation          : CVSには、","や"vbTab"など色々な種類があるためそれに合わせた区切りを使う
'   IndentionKey        : 改行コード、Excel上のは"vbCrLf","vbLf","vbCr "の3種類が存在する
'   CharacterCodeName   : 文字コード、"Unicode","UTF-8"など開きたい文字コードを指定。
'                             ↑  余談だがExcel2010で"CSV(カンマ区切り)"で保存した場合の基本文字コードは"Shift_JIS"
'------------------------------------------------------------------------------------------------------------------------------
Public Function OpenFile(CSV_Name As String, Optional Separation As String = ",", Optional IndentionKey As String = vbCrLf, Optional CharacterCodeName As String = "Shift_JIS") As Boolean
    'ファイルの存在チェック
    If Not Dir(CSV_Name) <> "" Then
        OpenFile = False
        Exit Function
    End If
    
    'データが本当にCSVか拡張子を調べる。
    Dim PassName() As String
    Dim Extension() As String
    PassName = Split(CSV_Name, "\")
    Extension = Split(PassName(UBound(PassName)), ".")
    If Not StrComp(Extension(UBound(Extension)), "csv", vbTextCompare) = 0 Then 'OSの関係で大文字、小文字の区別が付けれない。
        OpenFile = False
        Exit Function
    End If
    
    'どうやら正しいデータだったらしいのでようやくファイルを開く
    Dim Index       As Long
    Dim LineDataLength  As Long: LineDataLength = 0
    Dim EnterRec()  As String
    Dim StrRec()    As String
    Dim Buf         As String
    'CSVの情報を取得する
    With CreateObject("ADODB.Stream")
        .Type = 2   'adTypeText
        .Charset = CharacterCodeName
        .Open
        .LoadFromFile (CSV_Name)
        Buf = .ReadText(-1)
        .Close
        EnterRec = Split(Buf, IndentionKey)
        For Index = 0 To UBound(EnterRec)
            StrRec = Split(EnterRec(Index), Separation) '基本、カンマ区切りで配列へ
            If UBound(StrRec) >= 0 Then
                ReDim Preserve Data(LineDataLength)
                Data(LineDataLength).LineData = StrRec
                Data(LineDataLength).Index = UBound(StrRec)
                LineDataLength = LineDataLength + 1
            End If
        Next Index
    End With
    OpenFile = True
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   特定のいらない文字列を削除 : OpenFileを呼んでいないと意味はない
'------------------------------------------------------------------------------------------------------------------------------
Public Function DeleteString(Delete As String)
    If StrComp(Delete, "", vbBinaryCompare) = 0 Then Exit Function '何のために呼んだのか分からない
    Dim ColumnCount As Long: ColumnCount = 0
    Dim RowCount As Long: RowCount = 0
    '削除する
    For ColumnCount = 0 To GetColumnLength()
        For RowCount = 0 To GetRowLength(ColumnCount)
            Data(ColumnCount).LineData(RowCount) = Replace(Data(ColumnCount).LineData(RowCount), Delete, "")
        Next RowCount
    Next ColumnCount
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   ゲッター
'------------------------------------------------------------------------------------------------------------------------------
Public Function GetColumnLength() As Long '列数の取得
    GetColumnLength = UBound(Data)
End Function
Public Function GetRowLength(ColumnNum As Long) As Long '行数の取得
    If ColumnNum < 0 Or ColumnNum > GetColumnLength() Then
        GetRowLength = -&HFF
    Else
        GetRowLength = Data(ColumnNum).Index
    End If
End Function
Public Function GetRowData(ColumnNum As Long) As String() '列のデータ取得
    If ColumnNum < 0 Or ColumnNum > GetColumnLength() Then
        GetRowData = ""
    Else
        GetRowData = Data(ColumnNum).LineData
    End If
End Function
Public Function GetRowDatum(ColumnNum As Long, RowNum As Long) As String '列のデータ取得
    If ColumnNum < 0 Or ColumnNum > GetColumnLength() Then
        GetRowDatum = ""
    ElseIf RowNum < 0 Or GetRowLength(ColumnNum) < 0 Or RowNum > GetRowLength(ColumnNum) Then
        GetRowDatum = ""
    Else
        GetRowDatum = Data(ColumnNum).LineData(RowNum)
    End If
End Function

'------------------------------------------------------------------------------------------------------------------------------
'
'   キーの位置からデータを取得するタイプ用
'
'------------------------------------------------------------------------------------------------------------------------------
Public Function NextLine() As Boolean '最後まで達したか？達しました(True),まだだよ(false)
    Column = Column + 1   '行数を更新
    Row = 0      '列数を初期化
    If Column > GetColumnLength(Column) Then
        Column = 0
        NextLine = True
        Exit Function
    End If
    NextLine = False
End Function

'列の最初から読みたい人用
Public Function StateSearch1(State As String) As Boolean
    StateSearch1 = StateSearch2(State, 0)
End Function

'データを途中から読み込みたい人用
Public Function StateSearch2(State As String, RowNum As Long) As Boolean
    Row = RowNum 'まずは列数を初期化
    Dim i As Long
    For i = 0 To GetRowLength(Row)
        If State = GetRowDatum(Column, Row) Then
            Row = Row + 1
            StateSearch2 = True
            Exit Function
        End If
        Row = Row + 1
    Next i
    StateSearch2 = False
End Function

Public Function ToString() As String
    ToString = GetRowDatum(Column, Row)
    Row = Row + 1
    If Column > GetRowLength(Column) Then
        Column = 0
    End If
End Function
Public Function ToBoolean() As Boolean   'ブール型(Boolean)へのデータ変換
    ToBoolean = CBool(ToString())
End Function
Public Function ToByte() As Byte         'バイト型(Byte)へのデータ変換
    ToByte = CByte(ToString())
End Function
Public Function ToCurrency() As Currency '通貨型(Currency)へのデータ変換
    ToCurrency = CCur(ToString())
End Function
Public Function ToDate() As Date         '日付型(Date)へのデータ変換
    ToDate = CDate(ToString())
End Function
Public Function ToDouble() As Date       '倍精度浮動小数点実数型(Double)へのデータ変換
    ToDouble = CDbl(ToString())
End Function
Public Function ToInteger() As Date      '整数型(Integer)へのデータ変換
    ToInteger = CInt(ToString())
End Function
Public Function ToLong() As Long         '長整数型(Long)へのデータ変換
    ToLong = CLng(ToString())
End Function
Public Function ToSingle() As Single     '単精度浮動小数点実数型(Single)へのデータ変換
    ToSingle = CSng(ToString())
End Function
Public Function ToVariant() As Single     'バリアント型(Variant)へのデータ変換
    ToVariant = CVar(ToString())
End Function
