Attribute VB_Name = "Involved_Book"
Option Explicit
'##############################################################################################################################
'
'   ブック関連のマクロ
'   利用する場合は下記のインポートもお願いします
'   ・Involved_Other.bas
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2024/01/30
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'   等しい名前のシートを探す。
'
'   戻り値 : 等しい名前を持つシート。ない場合は、Nothingが返却される
'
'   sheetName : シート名
'   book : 対象のブック（任意）
'==============================================================================================================================
Public Function LEGACY_BookToEqualsName(ByVal bookName As String) As Workbook
    Set LEGACY_BookToEqualsName = Nothing

    Dim book As Workbook
    For Each book In Workbooks
        If StrComp(book.Name, bookName, vbBinaryCompare) = 0 Then
            Set LEGACY_BookToEqualsName = book
            Exit Function
        End If
    Next
End Function

'==============================================================================================================================
'   ブック本体のコピーを作成する。
'
'   戻り値 : 等しい名前を持つシート。ない場合は、Nothingが返却される
'
'   book          : 対象のブック（任意）
'   filename      : 保存名（空白の場合はbook名+現在時刻になる、拡張子不要）
'   pathname      : パス名（空白の場合は本体のブックと同階層になる）
'==============================================================================================================================
Public Function LEGACY_aCopyBook(ByRef book As Workbook, Optional ByVal filename As String = "", Optional ByVal pathname As String = "") As Boolean
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If StrComp(filename, "", vbBinaryCompare) = 0 Then
        filename = Replace(book.Name, "." + fso.GetExtensionName(book.FullName), "") + "_" + LEGACY_CurrentTime()
    End If
    
    If StrComp(pathname, "", vbBinaryCompare) = 0 Then
        pathname = book.path
    End If
    
    Dim copyFullpath As String
    Dim extensionname As String: extensionname = fso.GetExtensionName(book.FullName)
    
'fso.CopyFileでフォルダが存在していないとエラーになるため
On Error GoTo ErrorHandler_aCopyBook
    
    copyFullpath = pathname + "\" + filename + "." + extensionname
    fso.CopyFile book.FullName, copyFullpath
    
    Set fso = Nothing
    LEGACY_aCopyBook = True
    Exit Function
    
ErrorHandler_aCopyBook:
    Set fso = Nothing
    LEGACY_aCopyBook = False
End Function

