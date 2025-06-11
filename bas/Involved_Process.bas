Attribute VB_Name = "Involved_Process"
Option Explicit
'##############################################################################################################################
'
'   VBAを動かす際に必要不可欠な処理
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2024/07/05
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'プログラム実行時間を開始計測用の変数
Private beginTime As Date

'==============================================================================================================================
'
'   VBAのシート更新時の処理を軽減させる。
'   この関数を呼ぶだけで、処理時間が8.5倍ほど向上する。
'
'   参考URL
'   https://tonari-it.com/vba-processing-speed/
'
'==============================================================================================================================
Public Function LEGACY_reduceProcess_ToBegin()

    '実行時間計測開始
    beginTime = Time
    
    Application.Calculation = xlCalculationManual '計算モードをマニュアルにする
    Application.EnableEvents = False              'イベントを停止させる
    Application.ScreenUpdating = False            '画面表示更新を停止させる
    
        'Application.Cursor = xlWait                   'マウスポインタの状態を砂時計型に変更

End Function

Public Function LEGACY_reduceProcess_ToEnd()
    
    Application.Calculation = xlCalculationAutomatic '計算モードを自動にする
    Application.EnableEvents = True                  'イベントを開始させる
    Application.ScreenUpdating = True                '画面表示更新を開始させる
    
        'Application.Cursor = xlDefault                   'マウスポインタの状態を標準型に変更
    
    '実行時間計測終了
    Application.StatusBar = "処理完了 / 実行時間は " + Format(Time - beginTime, "nn分ss秒") + " でした"
    
End Function

'==============================================================================================================================
'   100%表示で表すバロメーターを追加をステータスバーに追加する
'
'   戻り値 : OK(True), NG(False)
'
'   message : ステータスバーに書きたいメッセージ
'   now     : 現在の値(for文中等にお使い下さい)
'   max     : 全体数
'==============================================================================================================================
Public Function LEGACY_StatusBar_100barometer(message As String, now As Long, max As Long)

    LEGACY_StatusBar_100barometer = False

    '0除算対策
    If max <= 0 Then Exit Function
    
    Dim text As String
    text = message + " ( " + CStr(CLng(now / max * 100)) + "% : " + CStr(now) + "/" + CStr(max) + ") "
    
    '100%超えは■や□の表示がおかしくなるため
    If now <= max Then
        text = text + String(CLng(now / max * 10), "■") + String(10 - CLng(now / max * 10), "□")
    End If
    
    Application.StatusBar = text
    LEGACY_StatusBar_100barometer = True

End Function
