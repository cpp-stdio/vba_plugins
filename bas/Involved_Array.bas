Attribute VB_Name = "Involved_Array"
Option Explicit
'##############################################################################################################################
'
'   配列関連関数
'   VBAの配列には2種類ある。Variantで変更可能なタイプかそうでないタイプ。これにより関数も2種類必要になる。
'
'   新規作成日 : 2019/11/18
'   最終更新日 : 2019/11/20
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'
'##############################################################################################################################

'==============================================================================================================================
'
'   配列が空なのかを判定する
'   この関数はVBAの仕様上、どこも関数化することが出来ない為、ほぼ同じコードを2回書く必要がある。
'   参考URL : http://www.fingeneersblog.com/1612/
'
'   戻り値 : 空(true),空ではない(false)
'
'   arrayVariant : 判定用の配列
'
'==============================================================================================================================
Public Function arrayIsEmpty(ByRef arrayVariant As Variant) As Boolean
    arrayIsEmpty = True '空だと仮定
On Error GoTo isEmptyArray_ErrorHandler

    'UBound関数を使用してエラーが発生するかどうかを確認
    If UBound(arrayVariant) > 0 Then
        arrayIsEmpty = False
    End If
    Exit Function
    
isEmptyArray_ErrorHandler:
    arrayIsEmpty = True
End Function

Public Function arrayIsEmptyEx(ByRef arrayVariant() As Variant) As Boolean
    arrayIsEmptyEx = True '空だと仮定
On Error GoTo isEmptyArrayEx_ErrorHandler

    'UBound関数を使用してエラーが発生するかどうかを確認
    If UBound(arrayVariant) > 0 Then
        arrayIsEmptyEx = False
    End If
    Exit Function

isEmptyArrayEx_ErrorHandler:
    arrayIsEmptyEx = True
End Function

'==============================================================================================================================
'
'   配列の一部を切り出し、新しい配列として返却する。
'
'   戻り値 : 成功(True), 失敗(False)
'
'   oldArray : 切り出し用の配列
'   newArray : 返却用配列
'   min      : どこから
'   max      : どこまで
'==============================================================================================================================
Public Function arraySplit(ByRef oldArray As Variant, ByRef newArray As Variant, Optional ByVal min As Long = -&HFF, Optional ByVal max As Long = -&HFF) As Boolean
    arraySplit = False '失敗と仮定
    If arrayIsEmpty(oldArray) Then Exit Function
    If errorSplit(min, max, LBound(oldArray), UBound(oldArray)) Then Exit Function
    'VBAの仕様上ここだけは個別で書かなければならない。
    Dim i As Long
    Dim length As Long: length = -1
    
    If VarType(newArray) = vbEmpty Then
        newArray = Array()
    End If
    
    For i = min To max
        length = length + 1
        ReDim Preserve newArray(length)
        newArray(length) = oldArray(i)
    Next i
    
    arraySplit = True
End Function

Public Function arraySplitEx(ByRef oldArray() As Variant, ByRef newArray() As Variant, Optional ByVal min As Long = -&HFF, Optional ByVal max As Long = -&HFF) As Boolean
    arraySplitEx = False '失敗と仮定
    If arrayIsEmptyEx(oldArray) Then Exit Function
    If errorSplit(min, max, LBound(oldArray), UBound(oldArray)) Then Exit Function
    'VBAの仕様上ここだけは個別で書かなければならない。
    Dim i As Long
    Dim length As Long: length = -1
    For i = min To max
        length = length + 1
        ReDim Preserve newArray(length)
        newArray(length) = oldArray(i)
    Next i
    
    arraySplitEx = True
End Function

Private Function errorSplit(ByRef min As Long, ByRef max As Long, ByVal minArray As Long, ByVal maxArray As Long) As Boolean
    errorSplit = True

    If min < minArray Then
        min = minArray
    End If
    
    If max > maxArray Then
        max = maxArray
    End If
    
    'VBAの仕様で同じ数字でもOKとする。
    If min < max Then Exit Function
    
    errorSplit = False
End Function

'==============================================================================================================================
'
'   配列の反転
'   この関数はVBAの仕様上、どこも関数化することが出来ない為、ほぼ同じコードを2回書く必要がある。
'
'   戻り値 : 成功(True), 失敗(False)
'
'   reversed : 反転する配列
'
'==============================================================================================================================
Public Function arrayReversed(ByRef oldArray As Variant, ByRef newArray As Variant) As Boolean
    arrayReversed = False
    If arrayIsEmpty(oldArray) Then Exit Function
    
    'oldArrayとnewArrayが同じだとメモリを破壊してしまう為
    Dim old As Variant
    old = arrayCopy(oldArray)
    
    ReDim newArray(UBound(old))
    
    Dim i As Long
    For i = LBound(old) To UBound(old)
        newArray(UBound(old) - i) = old(i)
    Next i
    arrayReversed = True
    
End Function

Public Function arrayReversedEx(ByRef oldArray() As Variant, ByRef newArray() As Variant) As Boolean
    arrayReversedEx = False
    If arrayIsEmptyEx(oldArray) Then Exit Function
    
    'oldArrayとnewArrayが同じだとメモリを破壊してしまう為
    Dim old() As Variant
    old = arrayCopyEx(oldArray)
    
    ReDim newArray(UBound(old))
    
    Dim i As Long
    For i = LBound(old) To UBound(old)
        newArray(UBound(old) - i) = old(i)
    Next i
    arrayReversedEx = True
End Function

'==============================================================================================================================
'
'   配列のコピー
'   この関数はVBAの仕様上、どこも関数化することが出来ない為、ほぼ同じコードを2回書く必要がある。
'
'   戻り値 : コピーした配列
'
'   copy : 反転する配列
'
'==============================================================================================================================
Public Function arrayCopy(ByRef copy As Variant) As Variant
    arrayCopy = Empty
    If arrayIsEmpty(copy) Then Exit Function

    Dim c As Variant
    ReDim c(UBound(copy))
    
    Dim i As Long
    For i = LBound(copy) To UBound(copy)
        c(i) = copy(i)
    Next i
    arrayCopy = c
End Function

Public Function arrayCopyEx(ByRef copy() As Variant) As Variant()
    Dim c() As Variant
    arrayCopyEx = c
    
    If arrayIsEmptyEx(copy) Then Exit Function

    ReDim c(UBound(copy))
    
    Dim i As Long
    For i = LBound(copy) To UBound(copy)
        c(i) = copy(i)
    Next i
    arrayCopyEx = c
End Function
