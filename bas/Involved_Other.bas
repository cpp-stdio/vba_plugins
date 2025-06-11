Attribute VB_Name = "Involved_Other"
Option Explicit
'##############################################################################################################################
'
'   その他、よくジャンル分け不能な関数群
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2024/01/26
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'   現在時刻を返す
'
'   T_Flag :  0,年から秒までのすべての時刻  (例)2018.04.23.03.54.02
'             1,年から日までの日付         (例)2018.04.23
'             2,時から秒までの時間         (例)03.54.02
'             3,年のみ                    (例)2018
'             4,月のみ                    (例)04
'             5,日のみ                    (例)23
'             6,時のみ                    (例)03
'             7,分のみ                    (例)54
'             8,秒のみ                    (例)02
'           　それ以外                    全て"0"として処理する
'
'   ToBe   : 間に入れてほしい文字列「2018/04/23」「2018.04.23」等 T_Flagの値が0～2の時のみ有効
'==============================================================================================================================
Public Function LEGACY_CurrentTime(Optional ByVal T_Flag As Long = 0, Optional ByVal ToBe As String = ".") As String

    Dim NowYear() As String
    Dim NowTime() As String
    NowYear = Split(Format(Date, "yyyy:mm:dd"), ":")
    NowTime = Split(Format(Time, "hh:mm:ss"), ":")

    Select Case T_Flag
        Case 1      '年から日までの日付
            LEGACY_CurrentTime = NowYear(0) + ToBe + NowYear(1) + ToBe + NowYear(2)
        Case 2      '時から秒までの時間
            LEGACY_CurrentTime = NowTime(0) + ToBe + NowTime(1) + ToBe + NowTime(2)
        Case 3      '年のみ
            LEGACY_CurrentTime = NowYear(0)
        Case 4      '月のみ
            LEGACY_CurrentTime = NowYear(1)
        Case 5      '日のみ
            LEGACY_CurrentTime = NowYear(2)
        Case 6      '時のみ
            LEGACY_CurrentTime = NowTime(0)
        Case 7      '分のみ
            LEGACY_CurrentTime = NowTime(1)
        Case 8      '秒のみ
            LEGACY_CurrentTime = NowTime(2)
        Case Else   '0を含め、それ以外
            LEGACY_CurrentTime = NowYear(0) + ToBe + NowYear(1) + ToBe + NowYear(2) + ToBe + NowTime(0) + ToBe + NowTime(1) + ToBe + NowTime(2)
    End Select
End Function
