Attribute VB_Name = "Involved_Split"
Option Explicit
'##############################################################################################################################
'
'   文字列分割関連
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2019/11/05
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'
'##############################################################################################################################

'==============================================================================================================================
'   文字の中にある、特定の文字列から特定の文字列までを取得する
'   戻り値 : 分割した文字列、Splitと名だが、specificAとspecificBも挿入されるので注意。
'
'   text      : とある文字列
'   specificA : 1つ目の特定の文字列
'   specificB : 2つ目の特定の文字列
'
'   解説
'       Dim text As String: text = "<HTML><HEAD>HOGEHOGE</HEAD><HEAD>GEHOGEHO</HEAD></HTML>"
'       Dim textArray() As String
'       textArray = BetweenSplit(text, "<HEAD>", "</HEAD>")
'       > ("<HTML>","<HEAD>","HOGEHOGE","</HEAD>","<HEAD>","GEHOGEHO","</HEAD>","</HTML>")
'==============================================================================================================================
Public Function BetweenSplit(ByVal text As String, ByVal specificA As String, ByVal specificB As String) As String()
    Dim returnLength As Long: returnLength = 0
    Dim returnString() As String
    ReDim returnString(returnLength)
    'エラー対応のため戻り値を初期化
    BetweenSplit = returnString
    '空白の挿入を確認
    If StrComp(text, "", vbBinaryCompare) = 0 Then Exit Function
    If StrComp(specificA, "", vbBinaryCompare) = 0 Then Exit Function
    If StrComp(specificB, "", vbBinaryCompare) = 0 Then Exit Function
    '同じ文字列なら用途が違う為
    If StrComp(specificA, specificB, vbBinaryCompare) = 0 Then
        BetweenSplit = Split(text, specificA)
        Exit Function
    End If

    '------------------------------
    ' "specificA" 側の処理
    '------------------------------
    Dim textArray1() As String
    textArray1 = Split(text, specificA)
    '------------------------------
    ' "specificB" 側の処理
    '------------------------------
    Dim textArray2() As String
    Dim count1 As Long: count1 = 0
    Dim count2 As Long: count2 = 0

    Dim Body As String
    For count1 = 0 To UBound(textArray1)
        '空白の場合は処理を行う必要がない為
        If Not StrComp(textArray1(count1), "", vbBinaryCompare) = 0 Then
            textArray2 = Split(textArray1(count1), specificB)
            '最初にカットされた、specificAの値を入れる
            If Not count1 = 0 Then
                ReDim Preserve returnString(returnLength)
                returnString(returnLength) = specificA
                returnLength = returnLength + 1
            End If
            '配列数が0以下の場合、文言がなかったので、そのまま挿入する。
            If UBound(textArray2) <= 0 Then
                ReDim Preserve returnString(returnLength)
                returnString(returnLength) = textArray1(count1)
                returnLength = returnLength + 1
            Else
                For count2 = 0 To UBound(textArray2)
                    'カットされたspecificBを戻す
                    If Not count2 = 0 Then
                        ReDim Preserve returnString(returnLength)
                        returnString(returnLength) = specificB
                        returnLength = returnLength + 1
                    End If
                    'カットされなかった分を元に戻す。
                    If Not StrComp(textArray2(count2), "", vbBinaryCompare) = 0 Then
                        ReDim Preserve returnString(returnLength)
                        returnString(returnLength) = textArray2(count2)
                        returnLength = returnLength + 1
                    End If
                Next count2
            End If
        End If
    Next count1
    BetweenSplit = returnString
End Function

'==============================================================================================================================
'   Split関数の複数版
'
'   delimiters : 本家とは違い、Optional型でない、配列でないとエラー出るので注意
'   min        : delimitersのどの位置から区切りするのか : 負の数、delimiters以上はエラー
'   max        : delimitersのどの位置まで区切りするのか : 負の数は全て区切る、delimiters以上でも全て区切る
'
'   その他、引数の説明は下記URLを参照
'   https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/split-function
'==============================================================================================================================
Public Function Splits(ByVal expression As String, delimiters() As String, Optional ByVal limit As Long = -1, Optional ByVal compare As VbCompareMethod = vbBinaryCompare, Optional ByVal min As Long = 0, Optional ByVal max As Long = -1) As String()
    Dim returnString() As String
    ReDim returnString(0)
    Splits = returnString
    If UBound(delimiters) < 0 Then Exit Function
    If min < 0 Then Exit Function
    If max < 0 Or max > UBound(delimiters) Then max = UBound(delimiters)

    Dim returnLength As Long: returnLength = 0
    Dim textCount As Long, textArray() As String
    Dim bodyCount As Long, bodyArray() As String
    Dim limitCount As Long, limitString As String
    '-1の部分はlimitなのでこれで良い
    textArray = Split(expression, delimiters(min), -1, compare)
    Splits = textArray
    If min = max Then Exit Function
    If min >= UBound(delimiters) Then Exit Function

    For textCount = 0 To UBound(textArray)
        If Not StrComp(textArray(textCount), "", vbBinaryCompare) = 0 Then
            '-1の部分はlimitなのでこれで良い
            bodyArray = Splits(textArray(textCount), delimiters, -1, compare, min + 1, max)
            For bodyCount = 0 To UBound(bodyArray)
                If Not StrComp(bodyArray(bodyCount), "", vbBinaryCompare) = 0 Then
                    ReDim Preserve returnString(returnLength)
                    returnString(returnLength) = bodyArray(bodyCount)
                    returnLength = returnLength + 1

                    If limit >= 0 And returnLength >= limit Then
                        limitString = ""
                        For limitCount = bodyCount To UBound(bodyArray)
                            limitString = limitString + bodyArray(limitCount)
                        Next limitCount

                        returnString(returnLength - 1) = limitString
                        Splits = returnString
                        Exit Function
                    End If
                End If
            Next bodyCount
        End If
    Next
    Splits = returnString
End Function
