Attribute VB_Name = "Involved_Debug"
Option Explicit
'##############################################################################################################################
'
'   デバック時にのみ有効な関数が存在する。
'   デバック用なので処理時間は考えられていない。速さを求めている開発では不向きと言えるかも知れない。
'
'   新規作成日 : 2017/08/30
'   最終更新日 : 2019/11/18
'
'   新規作成エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'   最終更新エクセルバージョン : Office Professional Plus 2010 , 14.0.7145.5000(32ビット)
'
'##############################################################################################################################

Private Enum atDevelopmentSwitching
    modeDebug   'Debugだけだとエラーが表示されたため
    modeRelease 'リリースモードの場合はこっち
End Enum

'全関数に有効なフラグ
Private Const atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug

'各関数を実行させるためのフラグ。関数を追加したらこっちも追加すること。
Private Const atDevelopment_debugBox = atDevelopmentSwitching.modeDebug
Private Const atDevelopment_debugBookSave = atDevelopmentSwitching.modeDebug
Private Const atDevelopment_debugModuleExport = atDevelopmentSwitching.modeDebug
Private Const atDevelopment_debugModuleExportAll = atDevelopmentSwitching.modeDebug
Private Const atDevelopment_debugModuleImport = atDevelopmentSwitching.modeDebug
'------------------------------------------------------------------------------------------------------------------------------
'   デバック用のMsgBox。毎回書くのが面倒なので作った。
'   引数の説明も戻り値の説明も下記を参照。一部不要な箇所があったので、そこだけ省略
'
'   https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function
'------------------------------------------------------------------------------------------------------------------------------
Public Function debugBox(ByRef prompt As String, Optional ByVal button As VbMsgBoxStyle = vbOKOnly, Optional ByRef title As String = "Microsoft Excel") As VbMsgBoxResult
    debugBox = vbOK
    'デバックモードでないと機能しない。
    If Not atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug Then Exit Function
    If Not atDevelopment_debugBox = atDevelopmentSwitching.modeDebug Then Exit Function
    
    debugBox = MsgBox(prompt, button, title)
End Function

'------------------------------------------------------------------------------------------------------------------------------
'   VBAをRANをした瞬間に上書き保存する機能がないので、セーブを手動で実装する。
'
'   book : 保存したいbook情報。設定しないとThisWorkbookが選択されます。
'------------------------------------------------------------------------------------------------------------------------------
Public Function debugBookSave(Optional ByRef book As Workbook = Nothing)
    
    'デバックモードでないと機能しない。
    If Not atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug Then Exit Function
    If Not atDevelopment_debugBookSave = atDevelopmentSwitching.modeDebug Then Exit Function
    
    Dim bookSave As Workbook
    If book Is Nothing Then
        Set bookSave = ThisWorkbook
    Else
        Set bookSave = book
    End If

    bookSave.Save
End Function

'==============================================================================================================================
'   自動モジュールインポート＆エクスポート、gitやsvn等でソース管理をしたい場合に便利
'
'   下記参考URL↓ とある参照設定にチェックをつけなければ動作しなかったが、
'   チェックを付けずともデフォルトの状態で動くようにするのに苦労した。
'
'   参考にしたインポートプログラム↓
'   https://vbabeginner.net/%E6%A8%99%E6%BA%96%E3%83%A2%E3%82%B8%E3%83%A5%E3%83%BC%E3%83%AB%E7%AD%89%E3%81%AE%E4%B8%80%E6%8B%AC%E3%82%A4%E3%83%B3%E3%83%9D%E3%83%BC%E3%83%88/
'   参考にしたエクスポートプログラム↓
'   https://vbabeginner.net/%E6%A8%99%E6%BA%96%E3%83%A2%E3%82%B8%E3%83%A5%E3%83%BC%E3%83%AB%E7%AD%89%E3%81%AE%E4%B8%80%E6%8B%AC%E3%82%A8%E3%82%AF%E3%82%B9%E3%83%9D%E3%83%BC%E3%83%88/
'
'   ※Excelの設定を以下の通りに変更(開発者専用)
'     この設定を行わないと、「実行時エラー 1004 プログラミングによる visual basic プロジェクトへのアクセスは信頼性に欠けます」
'     とエラーが表示されます。必ず行うようにして下さい。
'     フラグをmodeReleaseに変更することで、このエラーは発生しなくなります。
'
'       ファイル -> オプション -> セキュリティーセンター -> [セキュリティーセンターの設定]ボタンを押下
'       マクロ設定（左ペイン） -> [VBAプロジェクトオブジェクトモデルへのアクセスを信頼する]　チェックON
'
'==============================================================================================================================

'--------------------------------------------------------------
'   modulePaths : インポートするモジュールのパス名 : 例) Array("Involved_Debug")
'   book        : インポートするbook情報。設定しないとThisWorkbookが選択されます。
'--------------------------------------------------------------
Public Function debugModuleImport(ByRef modulePaths() As String, Optional ByVal book As Workbook = Nothing)

    'デバックモードでないと機能しない。
    If Not atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug Then Exit Function
    If Not atDevelopment_debugModuleImport = atDevelopmentSwitching.modeDebug Then Exit Function

    Dim extension  As String
    Dim name       As String
    Dim textFor    As Variant
    Dim module     As Object 'モジュール
    Dim moduleList As Object 'VBAプロジェクトの全モジュール
    Dim fso        As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim bookExport As Workbook
    If book Is Nothing Then
        Set bookExport = ThisWorkbook
    Else
        Set bookExport = book
    End If
    
    '処理対象ブックのモジュール一覧を取得
    Set moduleList = bookExport.VBProject.VBComponents
    
    'VBAの仕様でモジュール名がファイル名でない場合があるが対応出来ない為、ここでは考慮しない。
    For Each textFor In modulePaths
        '拡張子を小文字で取得
        extension = LCase(fso.GetExtensionName(textFor))
        'パス名から名前を取得
        name = fso.GetBaseName(textFor)
        '拡張子がいずれかの場合、インポートする。
        If StrComp(extension, "cls", vbBinaryCompare) = 0 Or _
            StrComp(extension, "frm", vbBinaryCompare) = 0 Or _
             StrComp(extension, "bas", vbBinaryCompare) = 0 Then
            
            For Each module In moduleList
                If StrComp(name, module.name, vbBinaryCompare) = 0 Then
                    '同名のモジュール削除
                    Call moduleList.Remove(module)
                    Exit For
                End If
            Next
            'モジュールを追加
            Call moduleList.Import(textFor)
        End If
    Next

End Function

'--------------------------------------------------------------
'   modules  : エクスポートしたいモジュール名 : 例) Array("Involved_Debug")
'   book     : エクスポートしたいbook情報。設定しないとThisWorkbookが選択されます。
'   filePath : エクスポートされるフォルダ先を指定する。指定がないとbookのバスが選択されます。
'--------------------------------------------------------------
Public Function debugModuleExport(ByRef modules() As String, Optional ByVal book As Workbook = Nothing, Optional ByVal folderPath As String = "")

    'デバックモードでないと機能しない。
    If Not atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug Then Exit Function
    If Not atDevelopment_debugModuleExport = atDevelopmentSwitching.modeDebug Then Exit Function
    
    'module.Typeはクラス内に書かれたEnumであり、アクセス不可の為、静的変数で代用する。
    Static vbext_ct_StdModule As Long: vbext_ct_StdModule = 1
    Static vbext_ct_MSForm As Long: vbext_ct_MSForm = 2
    Static vbext_ct_ClassModule As Long: vbext_ct_ClassModule = 3
    
    Dim module     As Object 'モジュール
    Dim moduleList As Object 'VBAプロジェクトの全モジュール
    Dim extension  As String  'モジュールの拡張子
    Dim textFor    As Variant
    
    Dim bookExport As Workbook
    If book Is Nothing Then
        Set bookExport = ThisWorkbook
    Else
        Set bookExport = book
    End If
    
    '処理対象ブックのモジュール一覧を取得
    Set moduleList = bookExport.VBProject.VBComponents
    
    '保存先の指定がないのでbookExportと同階層にエクスポートする。
    If StrComp(folderPath, "", vbBinaryCompare) = 0 Then
        folderPath = bookExport.path
    End If
    
    For Each module In moduleList
        extension = ""
        '拡張子を指定する。
        If (module.type = vbext_ct_ClassModule) Then
            extension = ".cls" 'クラス
        ElseIf (module.type = vbext_ct_MSForm) Then
            extension = ".frm" 'フォーム(.frxも一緒にエクスポートされる)
        ElseIf (module.type = vbext_ct_StdModule) Then
            extension = ".bas" '標準モジュール
        End If

        'エクスポート
        If Not StrComp(extension, "", vbBinaryCompare) = 0 Then
            For Each textFor In modules
                '配列の中に存在していれば、エクスポートする。
                If StrComp(textFor, module.name, vbBinaryCompare) = 0 Then
                    Call module.Export(folderPath + "\" + module.name + extension)
                End If
            Next
        End If
    Next
End Function

'--------------------------------------------------------------
'   book     : エクスポートしたいbook情報。設定しないとThisWorkbookが選択されます。
'   filePath : エクスポートされるフォルダ先を指定する。指定がないとbookのバスが選択されます。
'--------------------------------------------------------------
Public Function debugModuleExportAll(Optional ByVal book As Workbook = Nothing, Optional ByVal folderPath As String = "")

    'デバックモードでないと機能しない。
    If Not atDevelopmentSwitchingMode = atDevelopmentSwitching.modeDebug Then Exit Function
    If Not atDevelopment_debugModuleExportAll = atDevelopmentSwitching.modeDebug Then Exit Function
    
    Dim bookExport As Workbook
    If book Is Nothing Then
        Set bookExport = ThisWorkbook
    Else
        Set bookExport = book
    End If
    
    Dim module     As Object 'モジュール
    Dim moduleList As Object 'VBAプロジェクトの全モジュール
    Dim names() As String
    Dim length As Long: length = -1
    
    '処理対象ブックのモジュール一覧を取得
    Set moduleList = bookExport.VBProject.VBComponents
    For Each module In moduleList
        length = length + 1
        ReDim Preserve names(length)
        names(length) = module.name
    Next
    '保存する
    Call debugModuleExport(names, bookExport, folderPath)
End Function
