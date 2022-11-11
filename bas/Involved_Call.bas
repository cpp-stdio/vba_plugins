Attribute VB_Name = "Involved_Call"
Option Explicit
'##############################################################################################################################
'
'   セル関連
'
'   新規作成日 : 2022/11/10
'   最終更新日 : 2022/11/11
'
'   新規作成エクセルバージョン : Office Professional Plus 2016 , 16.0.5.56.1000(32ビット)
'   最終更新エクセルバージョン : Office Professional Plus 2016 , 16.0.5.56.1000(32ビット)
'
'##############################################################################################################################

'==============================================================================================================================
'   セル内で入力されている数式を値に変換する
'   ※1　利用する場合は「Involved_Other.bas」のインポートをお願いします
'   ※2　数式部分の自動再計算を行うので動作が重くなる可能性があります
'   ※3　「Involved_Sheet」にシート内全ての数式を値に変換するプログラムがあります
'
'   戻り値 : 変換完了(True), NG(False)
'
'   Range : セル情報を入力、Worksheet内のCellsとRangeは実は同じ型
'   rowMax : 変更範囲（任意）
'   columnMax : 変更範囲（任意）
'   sheetName : シート名
'==============================================================================================================================
Public Function cellsDeleteFormula(ByRef cells As Range, Optional ByVal rowMax As Long = 0, Optional ByVal columnMax As Long = 0) As Boolean
    cellsDeleteFormula = False
    If cells Is Nothing Then Exit Function
    
    Dim cell As Range
    Dim row As Long
    Dim column As Long
    Dim text As String
    Dim value As Variant
    
    '未入力の場合はMax値と判断
    If rowMax <= 0 Then rowMax = cells.Rows.Count - 1
    If columnMax <= 0 Then columnMax = cells.column.Count - 1
    
    For row = rowMax To 0 Step -1
        For column = columnMax To 0 Step -1
            Set cell = cells.Offset(row, column)
            
            If WorksheetFunction.IsFormula(cell) Then
                cell.Calculate '再計算
                text = cell.value
                '数値の場合はそのまま"数値"として表示させる（日付、金額等は対象外）
                If checkNumericalValue(text, value) Then
                    cell.value = value
                Else
                    cell.NumberFormatLocal = "@"
                    cell.value = text
                End If
            End If
        Next
    Next
    
    Set cell = Nothing
    cellsDeleteFormula = True
End Function
