Attribute VB_Name = "Involved_Process"
Option Explicit
'##############################################################################################################################
'
'   処理関連
'   「Involved_Debug」の関数を利用しているので、同時に読み込んでおくこと
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2019/11/04
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'
'##############################################################################################################################

'プログラム実行時間を開始計測用の変数
Private beginTime As Date

'==============================================================
'
'   プログラム処理時間を計測する。
'
'==============================================================

'実行時間計測開始
Public Function processMeasure_ToBegin()
    beginTime = Time
End Function

'実行時間計測終了
Public Function performanceMeasure_ToEnd(Optional ByRef message As String = "")
    Call debugBox(message + "実行時間は " + Format(Time - beginTime, "nn分ss秒") + " でした", vbInformation + vbOKOnly)
End Function

'==============================================================
'
'   VBAのシート更新時の処理を軽減させる。
'   この関数を呼ぶだけで、処理時間が8.5倍ほど向上する。
'
'   参考URL
'   https://tonari-it.com/vba-processing-speed/
'
'==============================================================
Public Function reduceProcess_ToBegin()
    Application.Calculation = xlCalculationManual '計算モードをマニュアルにする
    Application.EnableEvents = False              'イベントを停止させる
    Application.ScreenUpdating = False            '画面表示更新を停止させる
End Function

Public Function reduceProcess_ToEnd()
    Application.Calculation = xlCalculationAutomatic '計算モードを自動にする
    Application.EnableEvents = True                  'イベントを開始させる
    Application.ScreenUpdating = True                '画面表示更新を開始させる
End Function
