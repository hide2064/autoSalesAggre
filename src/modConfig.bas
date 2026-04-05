Attribute VB_Name = "modConfig"
Option Explicit

' ============================================================
' modConfig — 定数・マスタ読み込みモジュール
'
' 役割:
'   ・ワークブック全体で使う定数（列番号・シート名・セルアドレス・
'     HDR_* ヘッダー文字列・UI 色）を一元管理する。
'   ・Scripting.Dictionary の生成ファクトリ関数 NewDict() を提供する。
'   ・Config シートから製品マスタ・口銭マスタ・ヘッダー名寄せ・
'     Power Automate URL を読み込む関数群を提供する。
'   ・RunAll 後に集計シートの部署ドロップダウンを更新する
'     RefreshDeptList を提供する。
'   ・Config シートの整合性を確認する ValidateConfig を提供する。
'
' 設計方針:
'   ・HDR_* 定数が allシートのヘッダー書き込み・ProcessSourceSheet の
'     Select Case マッピング・modSetup のサンプルデータの三箇所で
'     共通して参照される唯一の真実の源(single source of truth)。
'   ・NewDict() を経由することで、全モジュールが同一の
'     大文字小文字を無視する(vbTextCompare)辞書を使用できる。
'   ・CLR_* 色定数を定義することで、配色変更を1箇所で完結させる。
'     値は RGB(r,g,b) = r + g*256 + b*65536 で事前計算した Long 値。
'   ・DICT_KEY_SEP をキーセパレータの唯一の定義とし、
'     modAggregation での辞書キー生成・分解を一元化する。
' ============================================================

' ============================================================
' Config シート — テーブル開始位置定数
'
' Config シートのレイアウト（列）:
'   A–B  : 製品マスタ   (製品コード → 製品名)         行3〜
'   D–E  : 口銭マスタ   (売上種別   → 口銭比率%)      行3〜
'   G–H  : ヘッダー名寄せ (正規名   → カンマ区切りエイリアス) 行3〜
'   J    : 集計用部署リスト (J2="全部署"固定, J3〜 自動更新)
'   L–M  : SharePoint連携  (L2=集計送信URL, L3=全データ送信URL)
'   O    : 保存済みフィルター条件 (O2=部署, O3=開始日, O4=終了日)
' ============================================================

' --- 製品マスタ (A列・B列) ---
Public Const CFG_PRODUCT_HDR_ROW    As Integer = 2   ' ヘッダー行番号 (A2="製品コード")
Public Const CFG_PRODUCT_COL        As Integer = 1   ' 製品コード列 (A列=1)

' --- 口銭マスタ (D列・E列) ---
Public Const CFG_COMMISSION_HDR_ROW As Integer = 2   ' ヘッダー行番号 (D2="売上種別")
Public Const CFG_COMMISSION_COL     As Integer = 4   ' 売上種別列 (D列=4)

' --- ヘッダー名寄せ設定 (G列・H列) ---
Public Const CFG_HEADER_HDR_ROW     As Integer = 2   ' ヘッダー行番号 (G2="正規名")
Public Const CFG_HEADER_COL         As Integer = 7   ' 正規名列 (G列=7)

' --- 集計用部署リスト (J列) ---
Public Const CFG_DEPT_HDR_ROW       As Integer = 2   ' J2="全部署" の固定行
Public Const CFG_DEPT_COL           As Integer = 10  ' 部署リスト列 (J列=10)

' --- SharePoint / Power Automate 設定 (L列・M列) ---
Public Const CFG_PA_LABEL_COL    As Integer = 12  ' L列: URL ラベル
Public Const CFG_PA_URL_COL      As Integer = 13  ' M列: URL 値を入力するセル
Public Const CFG_PA_URL_ROW      As Integer = 2   ' M2: 集計シート送信用 URL
Public Const CFG_PA_URL_ALL_ROW  As Integer = 3   ' M3: allシート全データ送信用 URL

' --- 保存済みフィルター条件 (O列) ---
' 集計シートの B1/B2/B3 の値を "保存" する際の格納先
Public Const CFG_SAVED_FILTER_COL  As Integer = 15  ' O列: 保存済みフィルター条件
Public Const CFG_SAVED_DEPT_ROW    As Integer = 2   ' O2: 保存済み部署
Public Const CFG_SAVED_FROM_ROW    As Integer = 3   ' O3: 保存済み開始日
Public Const CFG_SAVED_TO_ROW      As Integer = 4   ' O4: 保存済み終了日

' ============================================================
' all シート — 列インデックス定数 (1始まり)
'
' all シートは全ソースファイルのデータを正規化して集約するシート。
' 列1〜8 はソースから写し取り、列9〜10 はマスタ参照で計算、
' 列11 は読み込み元ファイル名を記録する。
' ============================================================
Public Const ALL_COL_CLIENT     As Integer = 1   ' A: 客先名
Public Const ALL_COL_PROD_CODE  As Integer = 2   ' B: 製品コード
Public Const ALL_COL_AMOUNT     As Integer = 3   ' C: 売上金額
Public Const ALL_COL_UNIT_PRICE As Integer = 4   ' D: 製品単価
Public Const ALL_COL_QTY        As Integer = 5   ' E: 売上数量
Public Const ALL_COL_DATE       As Integer = 6   ' F: 売上発生日
Public Const ALL_COL_SALE_TYPE  As Integer = 7   ' G: 売上種別
Public Const ALL_COL_DEPT       As Integer = 8   ' H: 部署           ← TSV由来の最終列
Public Const ALL_COL_PROD_NAME  As Integer = 9   ' I: 製品名 (製品マスタから計算)
Public Const ALL_COL_MARGIN     As Integer = 10  ' J: 部署取り分 (売上金額×口銭率で計算)
Public Const ALL_COL_SOURCE     As Integer = 11  ' K: ソースファイル名 (拡張子なし)
Public Const ALL_TOTAL_COLS     As Integer = 11  ' all シートの総列数

' TSV列からallシート列へのマッピング配列サイズ
' ProcessSourceSheet の colMap 配列はインデックス 0〜(COL_MAP_COUNT-1) を使う
' = ALL_COL_CLIENT-1 〜 ALL_COL_DEPT-1 の範囲 (現在は 0〜7 = 8要素)
Public Const COL_MAP_COUNT As Integer = 8  ' ALL_COL_DEPT - ALL_COL_CLIENT + 1

' ============================================================
' シート名定数
' ============================================================
Public Const SH_MAIN    As String = "main"        ' 実行ログ・操作ボタン
Public Const SH_CONFIG  As String = "Config"      ' マスタ・設定値
Public Const SH_ALL     As String = "all"         ' 全ソースデータ集約
Public Const SH_AGGR    As String = "集計"        ' 部署・日付フィルタ付き集計表示
Public Const SH_PIVOT   As String = "ピボット"    ' Excel ネイティブ PivotTable
Public Const SH_ERROR   As String = "エラー"      ' データ処理エラーレポート
Public Const SH_MONTHLY As String = "月次サマリー" ' 月別売上サマリー

' ============================================================
' UI 色定数
' 算出式: CLR = R + G*256 + B*65536
' 変更時はここだけ修正すれば全シート・全グラフに反映される。
' ============================================================
Public Const CLR_HEADER_BG    As Long = 15785160  ' RGB(200,220,240) ヘッダー行の青系背景
Public Const CLR_GROUP_ROW    As Long = 14474460  ' RGB(220,220,220) 集計グループ行のグレー背景
Public Const CLR_CHART_AMT    As Long = 11829830  ' RGB( 70,130,180) グラフ 売上金額棒（鋼鉄青）
Public Const CLR_CHART_MARGIN As Long = 42495     ' RGB(255,165,  0) グラフ 口銭総額棒（オレンジ）
Public Const CLR_PLOT_AREA    As Long = 16316664  ' RGB(248,248,248) グラフ プロットエリア薄グレー
Public Const CLR_LABEL_TEXT   As Long = 6579300   ' RGB(100,100,100) 説明文などの薄いテキスト色
Public Const CLR_ERROR_ROW    As Long = 14474495  ' RGB(255,220,220) エラー行の薄赤背景
Public Const CLR_MONTHLY_HDR  As Long = 14479560  ' RGB(200,240,220) 月次サマリーヘッダーの薄緑

' ============================================================
' 集計キーのセパレータ定数
' dictSummary のキー形式: 製品名 & DICT_KEY_SEP & 客先名
' "||" は通常テキストに含まれにくい文字列を選択。
' modAggregation でのキー生成・Split 両方に使用する。
' ============================================================
Public Const DICT_KEY_SEP As String = "||"

' ============================================================
' 集計シート — セルアドレス・行番号定数
'
' 集計シートのレイアウト:
'   A1:B1  部署選択 (B1 にドロップダウン)
'   A2:B2  開始日   (B2 に日付入力)
'   A3:B3  終了日   (B3 に日付入力)
'   5行目  集計テーブルのヘッダー行
'   6行目〜 集計データ行 (製品グループ行 + 客先明細行 + 総合計行)
' ============================================================
Public Const AGGR_DEPT_CELL As String = "B1"  ' 部署選択ドロップダウンセル
Public Const AGGR_FROM_CELL As String = "B2"  ' 集計開始日入力セル
Public Const AGGR_TO_CELL   As String = "B3"  ' 集計終了日入力セル
Public Const AGGR_HDR_ROW   As Integer = 5    ' 集計テーブルヘッダー行番号
Public Const AGGR_DATA_ROW  As Integer = 6    ' 集計テーブルデータ開始行番号

' ============================================================
' main シート定数
' ============================================================
Public Const MAIN_LOG_START_ROW As Integer = 3  ' ログ書き込み開始行 (1行目=タイトル, 2行目=列ヘッダー)

' ============================================================
' 正規ヘッダー名定数 (HDR_*)
'
' 全モジュールで共通して参照するヘッダー文字列の唯一の定義。
' 使用箇所:
'   ・modSetup.SetupAllSheet — all シートのヘッダー書き込み
'   ・modDataProcess.BuildAllSheet — all シートのヘッダー再書き込み
'   ・modDataProcess.ProcessSourceSheet — Select Case によるマッピング
'   ・modSetup.SetupConfigSheet — ヘッダー名寄せサンプルデータの正規名
' ============================================================
Public Const HDR_CLIENT     As String = "客先名"
Public Const HDR_PROD_CODE  As String = "製品コード"
Public Const HDR_AMOUNT     As String = "売上金額"
Public Const HDR_UNIT_PRICE As String = "製品単価"
Public Const HDR_QTY        As String = "売上数量"
Public Const HDR_DATE       As String = "売上発生日"
Public Const HDR_SALE_TYPE  As String = "売上種別"
Public Const HDR_DEPT       As String = "部署"
Public Const HDR_PROD_NAME  As String = "製品名"
Public Const HDR_MARGIN     As String = "部署取り分"
Public Const HDR_SOURCE     As String = "ソースファイル"

' ============================================================
' NewDict — Scripting.Dictionary 生成ファクトリ
'
' 戻り値: 大文字小文字を無視する(vbTextCompare) Dictionary オブジェクト
'
' 全モジュールでこの関数を経由することにより、辞書の比較モードを
' 一箇所で統一し、呼び出し元でのモード指定ミスを防ぐ。
' ============================================================
Public Function NewDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare  ' キーの大文字小文字を区別しない
    Set NewDict = d
End Function

' ============================================================
' LoadProductDict — 製品マスタ読み込み
'
' Config シートの A列(製品コード)・B列(製品名) を読み込み、
' 「製品コード → 製品名」の辞書を返す。
' 同一コードが複数行ある場合は先着優先（2件目以降は無視）。
'
' 戻り値: Dictionary(製品コード As String → 製品名 As String)
' ============================================================
Public Function LoadProductDict() As Object
    Dim dict As Object
    Dim ws As Worksheet
    Dim r As Long
    Dim code As String

    Set dict = NewDict()
    Set ws = ThisWorkbook.Sheets(SH_CONFIG)

    ' ヘッダー行の次行(3行目)から空欄になるまで読み込む
    r = CFG_PRODUCT_HDR_ROW + 1
    Do While Trim(CStr(ws.Cells(r, CFG_PRODUCT_COL).Value)) <> ""
        code = Trim(CStr(ws.Cells(r, CFG_PRODUCT_COL).Value))
        If dict.Exists(code) Then
            ' 同一コードの2件目以降は先着優先で無視するが警告を記録する
            Debug.Print "modConfig: 製品コード重複 [" & code & "] 行" & r & " (先着優先で無視)"
        Else
            Dim prodName As String
            prodName = Trim(CStr(ws.Cells(r, CFG_PRODUCT_COL + 1).Value))
            If prodName = "" Then
                Debug.Print "modConfig: 製品名が空です [コード=" & code & "] 行" & r
            End If
            dict(code) = prodName
        End If
        r = r + 1
    Loop

    Set LoadProductDict = dict
End Function

' ============================================================
' LoadCommissionDict — 口銭マスタ読み込み
'
' Config シートの D列(売上種別)・E列(口銭比率%) を読み込み、
' 「売上種別 → 口銭比率(Double)」の辞書を返す。
' 口銭比率が数値でない場合は 0 を登録し、デバッグ出力に警告を残す。
'
' 戻り値: Dictionary(売上種別 As String → 口銭比率% As Double)
' ============================================================
Public Function LoadCommissionDict() As Object
    Dim dict As Object
    Dim ws As Worksheet
    Dim r As Long
    Dim stype As String
    Dim rateVal As Variant

    Set dict = NewDict()
    Set ws = ThisWorkbook.Sheets(SH_CONFIG)

    ' ヘッダー行の次行(3行目)から空欄になるまで読み込む
    r = CFG_COMMISSION_HDR_ROW + 1
    Do While Trim(CStr(ws.Cells(r, CFG_COMMISSION_COL).Value)) <> ""
        stype = Trim(CStr(ws.Cells(r, CFG_COMMISSION_COL).Value))
        If Not dict.Exists(stype) Then
            rateVal = ws.Cells(r, CFG_COMMISSION_COL + 1).Value
            If IsNumeric(rateVal) Then
                dict(stype) = CDbl(rateVal)
            Else
                ' 数値でない場合は 0 を登録してログに警告を残す
                dict(stype) = 0
                Debug.Print "modConfig: 口銭比率が数値でありません [" & stype & "] = " & CStr(rateVal)
            End If
        End If
        r = r + 1
    Loop

    Set LoadCommissionDict = dict
End Function

' ============================================================
' LoadHeaderMap — ヘッダー名寄せ辞書の読み込み
'
' Config シートの G列(正規名)・H列(カンマ区切りエイリアス) を読み込み、
' 「エイリアス(小文字) → 正規名」の辞書を返す。
' 正規名自身もエイリアスとして登録することで、既に正規名が使われて
' いるファイルも同じロジックで処理できる。
'
' 戻り値: Dictionary(エイリアス(小文字) As String → 正規名 As String)
' ============================================================
Public Function LoadHeaderMap() As Object
    Dim dict As Object
    Dim ws As Worksheet
    Dim r As Long
    Dim canonical As String  ' G列の正規名
    Dim aliases As String    ' H列のカンマ区切りエイリアス文字列
    Dim parts() As String    ' エイリアスを Split した配列
    Dim i As Integer
    Dim a As String          ' 個々のエイリアス(小文字化済み)

    Set dict = NewDict()
    Set ws = ThisWorkbook.Sheets(SH_CONFIG)

    r = CFG_HEADER_HDR_ROW + 1
    Do While Trim(CStr(ws.Cells(r, CFG_HEADER_COL).Value)) <> ""
        canonical = Trim(CStr(ws.Cells(r, CFG_HEADER_COL).Value))
        aliases   = Trim(CStr(ws.Cells(r, CFG_HEADER_COL + 1).Value))

        ' 正規名自身を小文字キーで登録（ソースが既に正規名を使っている場合に対応）
        If Not dict.Exists(LCase(canonical)) Then dict(LCase(canonical)) = canonical

        ' カンマ区切りのエイリアスを分解して個別に登録
        parts = Split(aliases, ",")
        For i = 0 To UBound(parts)
            a = LCase(Trim(parts(i)))
            If a <> "" And Not dict.Exists(a) Then dict(a) = canonical
        Next i
        r = r + 1
    Loop

    Set LoadHeaderMap = dict
End Function

' ============================================================
' LoadPowerAutomateUrl — 集計シート送信用 Power Automate URL の読み込み
'
' Config シートの M2 に設定された Power Automate URL を返す。
' 未設定（空欄）の場合は空文字列を返す。
'
' 戻り値: URL 文字列（空欄時は ""）
' ============================================================
Public Function LoadPowerAutomateUrl() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_CONFIG)
    LoadPowerAutomateUrl = Trim(CStr(ws.Cells(CFG_PA_URL_ROW, CFG_PA_URL_COL).Value))
End Function

' ============================================================
' LoadPowerAutomateUrlAll — allシート全データ送信用 Power Automate URL の読み込み
'
' Config シートの M3 に設定された Power Automate URL を返す。
' 未設定（空欄）の場合は M2 の集計用 URL にフォールバックする。
' これにより M3 未設定時でも既存の M2 URL で動作を継続できる。
'
' 戻り値: URL 文字列（M3 も M2 も空欄の場合は ""）
' ============================================================
Public Function LoadPowerAutomateUrlAll() As String
    Dim ws As Worksheet
    Dim urlAll As String

    Set ws = ThisWorkbook.Sheets(SH_CONFIG)
    urlAll = Trim(CStr(ws.Cells(CFG_PA_URL_ALL_ROW, CFG_PA_URL_COL).Value))

    ' M3 が未設定の場合は M2（集計用）にフォールバック
    If urlAll = "" Then
        urlAll = Trim(CStr(ws.Cells(CFG_PA_URL_ROW, CFG_PA_URL_COL).Value))
    End If

    LoadPowerAutomateUrlAll = urlAll
End Function

' ============================================================
' ValidateConfig — Config シートの整合性チェック
'
' 戻り値: 問題件数 (0 = 問題なし)
'
' チェック内容:
'   1. 製品マスタの登録件数が 0 → 警告
'   2. 口銭マスタの登録件数が 0 → 警告
'   3. ヘッダー名寄せの登録件数が 0 → 警告
'   4. 口銭比率が 0〜100 の範囲外 → 警告
'
' 各問題は LogMessage に出力される。
' RunAll 冒頭で呼び出すことで設定漏れを早期発見できる。
' ============================================================
Public Function ValidateConfig() As Integer
    Dim ws As Worksheet
    Dim issueCount As Integer
    Dim r As Long
    Dim rateVal As Variant
    Dim rate As Double

    Set ws = ThisWorkbook.Sheets(SH_CONFIG)
    issueCount = 0

    ' --- 製品マスタの件数チェック ---
    If Trim(CStr(ws.Cells(CFG_PRODUCT_HDR_ROW + 1, CFG_PRODUCT_COL).Value)) = "" Then
        LogMessage "[Config警告] 製品マスタが空です。製品名がすべて[未登録]になります。"
        issueCount = issueCount + 1
    End If

    ' --- 口銭マスタの件数チェック ---
    If Trim(CStr(ws.Cells(CFG_COMMISSION_HDR_ROW + 1, CFG_COMMISSION_COL).Value)) = "" Then
        LogMessage "[Config警告] 口銭マスタが空です。部署取り分がすべて0になります。"
        issueCount = issueCount + 1
    End If

    ' --- ヘッダー名寄せの件数チェック ---
    If Trim(CStr(ws.Cells(CFG_HEADER_HDR_ROW + 1, CFG_HEADER_COL).Value)) = "" Then
        LogMessage "[Config警告] ヘッダー名寄せが未設定です。列マッピングが機能しません。"
        issueCount = issueCount + 1
    End If

    ' --- 口銭比率の範囲チェック (0〜100%) ---
    r = CFG_COMMISSION_HDR_ROW + 1
    Do While Trim(CStr(ws.Cells(r, CFG_COMMISSION_COL).Value)) <> ""
        rateVal = ws.Cells(r, CFG_COMMISSION_COL + 1).Value
        If IsNumeric(rateVal) Then
            rate = CDbl(rateVal)
            If rate < 0 Or rate > 100 Then
                LogMessage "[Config警告] 口銭比率が範囲外です [" & _
                           CStr(ws.Cells(r, CFG_COMMISSION_COL).Value) & "] = " & rate & "%"
                issueCount = issueCount + 1
            End If
        End If
        r = r + 1
    Loop

    ValidateConfig = issueCount
End Function

' ============================================================
' RefreshDeptList — 部署リストの更新と集計シートドロップダウンの再設定
'
' 引数:
'   dictDept — CollectUniqueDepts() が返す「部署名 → 1」の辞書
'
' 処理内容:
'   1. Config!J3 以降をクリアして最新の部署名を書き込む
'   2. J2 に "全部署" を固定値として設定する
'   3. 集計!B1 のリスト検証(Validation)を新しい部署範囲に更新する
'   4. 集計!B1 が空の場合は "全部署" を初期値としてセットする
' ============================================================
Public Sub RefreshDeptList(dictDept As Object)
    Dim ws As Worksheet
    Dim clearRow As Long
    Dim r As Long
    Dim key As Variant
    Dim lastDeptRow As Long
    Dim wsAggr As Worksheet

    Set ws = ThisWorkbook.Sheets(SH_CONFIG)

    ' --- J3 以降の既存部署データをクリア ---
    clearRow = CFG_DEPT_HDR_ROW + 1
    Do While Trim(CStr(ws.Cells(clearRow, CFG_DEPT_COL).Value)) <> ""
        ws.Cells(clearRow, CFG_DEPT_COL).ClearContents
        clearRow = clearRow + 1
    Loop

    ' --- J2 に "全部署" を固定値として設定 ---
    ws.Cells(CFG_DEPT_HDR_ROW, CFG_DEPT_COL).Value = "全部署"

    ' --- J3 以降に新しい部署名を書き込む ---
    r = CFG_DEPT_HDR_ROW + 1
    For Each key In dictDept.Keys
        ws.Cells(r, CFG_DEPT_COL).Value = key
        r = r + 1
    Next key

    lastDeptRow = r - 1  ' 部署リストの最終行（ドロップダウン範囲の終端）

    ' --- 集計!B1 のドロップダウンリストを新しい部署範囲に更新 ---
    Set wsAggr = ThisWorkbook.Sheets(SH_AGGR)
    With wsAggr.Range(AGGR_DEPT_CELL).Validation
        .Delete  ' 既存の Validation を削除してから再設定
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=" & SH_CONFIG & "!$J$" & CFG_DEPT_HDR_ROW & ":$J$" & lastDeptRow
    End With

    ' B1 が空になった場合は "全部署" を初期値としてセット
    If Trim(CStr(wsAggr.Range(AGGR_DEPT_CELL).Value)) = "" Then
        wsAggr.Range(AGGR_DEPT_CELL).Value = "全部署"
    End If
End Sub
