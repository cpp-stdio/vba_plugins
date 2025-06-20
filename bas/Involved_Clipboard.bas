Attribute VB_Name = "Involved_Clipboard"
Option Explicit
'##############################################################################################################################
'
'   クリップボード関連（RPA時専用関数）
'   使用するには、「Microsoft Forms 2.0 Object Library」を参照設定します。
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2024/01/30
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Microsoft 365 Apps for enterprise
'
'##############################################################################################################################

'==============================================================================================================================
'   クリップボードに文字列を設定する。
'
'   text : クリップボードにアップするテキスト情報
'==============================================================================================================================
Public Function LEGACY_SetClipboard_Text(ByVal text As String)
    If StrComp(text, "", vbBinaryCompare) = 0 Then Exit Function
    With New MSForms.DataObject
        .SetText text
        .PutInClipboard
    End With
End Function

'==============================================================================================================================
'   クリップボードから文字列を取得する。
'==============================================================================================================================
Public Function LEGACY_GetClipboard_Text() As String
    Dim text As String: text = ""
    With New MSForms.DataObject
        .GetFromClipboard
        text = .GetText
    End With
    GetText = text
End Function
