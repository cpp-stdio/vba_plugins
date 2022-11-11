Attribute VB_Name = "Involved_Book"
Option Explicit
'##############################################################################################################################
'
'   ブック関連のマクロ
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2019/10/28
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
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
Public Function BookToEqualsName(ByVal bookName As String) As Workbook
    Set BookToEqualsName = Nothing

    Dim book As Workbook
    For Each book In Workbooks
        If StrComp(book.name, bookName, vbBinaryCompare) = 0 Then
            Set BookToEqualsName = book
            Exit Function
        End If
    Next
End Function
