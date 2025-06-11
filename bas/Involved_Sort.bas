Attribute VB_Name = "Sort"
Option Explicit
'##############################################################################################################################
'
'   シート関連
'
'   新規作成日 : 2017/06/21
'   最終更新日 : 2024/01/30
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################
'------------------------------------------------------------------------------------------------------------------------------
'   数が少ない場合に使うソート
'   SortData : ソートしたい一次元配列
'------------------------------------------------------------------------------------------------------------------------------
Function LEGACY_BubbleSort(SortData As Variant, Min As Long, Max As Long)
    Dim Tmp As Variant '配列移動用
    Dim X As Long, Y As Long
    For X = Min To Max
        For Y = Min To Max
            If SortData(X) < SortData(Y) Then
                Tmp = SortData(X)
                SortData(X) = SortData(Y)
                SortData(Y) = Tmp
            End If
        Next Y
    Next X
End Function
'------------------------------------------------------------------------------------------------------------------------------
'   数が多い場合に使うソート
'   SortData : ソートしたい一次元配列
'------------------------------------------------------------------------------------------------------------------------------
Function LEGACY_QuickSort(SortData As Variant, Min As Long, Max As Long)
    Dim Left As Long: Left = Min    '左ループカウンタ
    Dim Right As Long: Right = Max  '右ループカウンタ
    Dim Median As Variant           '中央値
    Dim Tmp As Variant              '配列移動用
    'ソート終了位置省略時は配列要素数を設定
    If (Right <= -1) Then
        Right = UBound(SortData)
    End If
    Median = SortData((Min + Max) / 2)
    Do
        '配列の左側から中央値より大きい値を探す
        Do
            If (SortData(Left) >= Median) Then
                Exit Do
            End If
            Left = Left + 1
        Loop
        '配列の右側から中央値より大きい値を探す
        Do
            If (Median >= SortData(Right)) Then
                Exit Do
            End If
            Right = Right - 1
        Loop
        '左側の方が大きければここで処理終了する
        If Left >= Right Then
            Exit Do
        End If
        '右側の方が大きい場合は、左右を入れ替える
        Tmp = SortData(Left)
        SortData(Left) = SortData(Right)
        SortData(Right) = Tmp
        '// 左側を１つ右にずらす
        Left = Left + 1
        '// 右側を１つ左にずらす
        Right = Right - 1
    Loop
    '中央値の左側を再帰で恐ろしいクイックソートの開始
    If (Min < Left - 1) Then
        Call LEGACY_QuickSort(SortData, Min, Left - 1)
    End If
    '中央値の右側を再帰で恐ろしいクイックソートの開始
    If (Right + 1 < Max) Then
        Call LEGACY_QuickSort(SortData, Right + 1, Max)
    End If
End Function
