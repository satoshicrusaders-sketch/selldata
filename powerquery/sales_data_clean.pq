// ============================================================================
// 販売実績（月次一覧）スクリーニング用 Power Query M スクリプト
//
// 対象: OneDrive for Business / SharePoint 上の
//   07. 売上関連 / 月次一覧 - コピー / [8期|9期|10期|11期|12期|13期] / YYYY.MM.xlsx (or YYYYMM.xlsx)
//
// 前提:
//   同じPower BIファイル内に、以下2つのクエリがすでに作成済みであること
//   (このスクリプトはクエリ名で直接参照するため、先にこれらを作成・実行してから
//    このクエリを作成すること)。
//     - 「病院マスタ」  (hospital_master.pq)
//     - 「取引先マスタ」(distributor_master.pq)
//
// 使い方:
//   1. Power BI Desktop > Power Query エディター > 新しいクエリ > 空のクエリ
//   2. 詳細エディター(Advanced Editor)を開き、このファイルの中身を丸ごと貼り付け
//   3. 下の RootFolder の値を、実際にOneDriveが同期されているローカルパスに書き換える
//      (例: "C:\Users\sfuru\OneDrive - M'z company\07. 売上関連\月次一覧 - コピー")
//   4. クエリ名を「売上実績_統合」などに変更して保存・更新
//
// このスクリプトが自動で行うこと:
//   - フォルダ配下(期ごとのサブフォルダを含む)の全 .xlsx を再帰的に取得
//   - ファイル名から年月を抽出(「2021.01.xlsx」「202409.xlsx」の両方に対応)
//   - シート内でヘッダー行(「取引先」を含む行)を自動検出し、その手前のタイトル行を除去
//   - 列名の表記ゆれを統一(例: 「メーカー」→「メーカー①」、「ロットNo」→「ロット№」)
//   - シート内に複数回登場するページ見出し行・合計行を除外
//   - 期によって列構成が違っても Table.Combine が自動で欠落列をnull補完し、1つの表に統合
//   - 病院コード(伝票番号1の下3桁 → 取れなければSource.Nameの●と-の間の下3桁)を抽出し、
//     病院マスタとコード一致→(だめなら)病院名一致の優先順で住所・電話番号・URLを直接マージ
//   - 伝票番号1(=貸出先コード)→(だめなら)取引先名/略称一致の優先順で、
//     取引先マスタから正式な取引先名称を「取引先_統一」列として解決
//   - メーカー①②③の半角カナを全角カナに正規化し、前後・連続スペースを統一
//     (メーカー専用マスタは無いため機械的な正規化のみ)
//
// 注意:
//   - 期によって列構成が大きく異なる(例: 11-12期は数量/単価/担当者列が存在しない)ため、
//     結合後のテーブルではこれらの列が期によって空欄になるのは仕様です。
//   - 実際のシート名やヘッダー位置がここで想定した構造と異なる月がある場合、
//     fnExtractSalesTable 内でエラーが出るので、そのファイルだけ個別に確認してください。
// ============================================================================

let
    RootFolder = "C:\Users\sfuru\OneDrive - M'z company\07. 売上関連\月次一覧 - コピー",

    // ヘッダー/列名から半角・全角スペースと改行を除去して比較しやすくする(null安全)
    fnCleanName = (n as any) as text =>
        let
            t = if n = null then "" else Text.From(n),
            noHalfSpace = Text.Replace(t, " ", ""),
            noFullSpace = Text.Replace(noHalfSpace, "　", ""),
            noLineBreak = Text.Replace(Text.Replace(noFullSpace, "#(lf)", ""), "#(cr)", "")
        in
            Text.Trim(noLineBreak),

    // 列名の表記ゆれ統一マップ(実データ変更のみ差分登録。ほとんどの列は元々統一名)
    ColumnRenameMap = {
        {"メーカー", "メーカー①"},
        {"ロットNo", "ロット№"},
        {"ロットNo.", "ロット№"},
        {"ロットNO", "ロット№"},
        {"ロットNo№", "ロット№"}
    },

    // 1枚のシート/テーブル(生データ)を処理する。「取引先」ヘッダーが
    // 見つからない場合は null を返す(そのシートは対象外とみなす)。
    //
    // 2パターンに対応:
    //   (a) 名前付きテーブル(Excel の ListObject)は Excel.Workbook が最初から
    //       正式な列名(=「取引先」等)を持つ状態で返してくるので、そのまま使う
    //   (b) 通常のシートはタイトル行等が混ざった生データなので、
    //       「取引先」という値を含む行を自前で探してその行をヘッダーに昇格する
    fnProcessSheetData = (sheet as table, fileName as text, yearMonth as text) as nullable table =>
        let
            existingColNamesCleaned = List.Transform(Table.ColumnNames(sheet), fnCleanName),
            alreadyHasHeader = List.Contains(existingColNamesCleaned, "取引先"),

            promotedTable =
                if alreadyHasHeader then
                    sheet
                else
                    let
                        indexed = Table.AddIndexColumn(sheet, "__RowIndex", 0, 1),
                        otherColNames = List.RemoveItems(Table.ColumnNames(indexed), {"__RowIndex"}),
                        // 「取引先」という文字列(スペース除去後)を含む行をヘッダー行とみなす
                        headerMatches = Table.SelectRows(
                            indexed,
                            each List.AnyTrue(
                                List.Transform(
                                    otherColNames,
                                    (colName) => fnCleanName(Record.Field(_, colName)) = "取引先"
                                )
                            )
                        )
                    in
                        if Table.RowCount(headerMatches) = 0 then
                            null
                        else
                            Table.PromoteHeaders(
                                Table.Skip(sheet, headerMatches{0}[__RowIndex]),
                                [PromoteAllScalars = true]
                            ),

            result =
                if promotedTable = null then
                    null
                else
                    let
                        cleanedHeaders = Table.TransformColumnNames(promotedTable, fnCleanName),
                        renamed = Table.RenameColumns(cleanedHeaders, ColumnRenameMap, MissingField.Ignore),

                        keyCol =
                            if List.Contains(Table.ColumnNames(renamed), "取引先") then
                                "取引先"
                            else
                                Table.ColumnNames(renamed){0},

                        // シート内に複数回出てくるページ見出し行(値が"取引先"そのもの)と空行を除外
                        noRepeatedHeaders = Table.SelectRows(
                            renamed,
                            each
                                Record.Field(_, keyCol) <> null
                                and fnCleanName(Record.Field(_, keyCol)) <> "取引先"
                                and fnCleanName(Record.Field(_, keyCol)) <> ""
                        ),

                        // 末尾の合計行(取引先列に会社名でなく件数などの数値が入る)を除外
                        noFooterTotal = Table.SelectRows(
                            noRepeatedHeaders,
                            each
                                not (
                                    try Number.From(Record.Field(_, keyCol)) is number
                                    otherwise false
                                )
                        ),

                        withMeta = Table.AddColumn(
                            Table.AddColumn(noFooterTotal, "年月", each yearMonth, type text),
                            "ソースファイル",
                            each fileName,
                            type text
                        )
                    in
                        withMeta
        in
            result,

    // 1ファイル(1か月分)を読み込み、正規化済みテーブルを返す。
    // ブック内の全シート/全テーブルを検索し、「取引先」ヘッダーを含むものだけを採用して結合する。
    // (期によってデータが1枚目のシートにあるとは限らないため、位置決め打ちにしない)
    fnExtractSalesTable = (content as binary, fileName as text) as table =>
        let
            digitsOnly = Text.Select(fileName, {"0".."9"}),
            yearMonth = Text.Start(digitsOnly, 6),

            // useHeaders を null にして、通常シートは生データ(未昇格)のまま受け取る。
            // (true にすると先頭行=タイトル行が誤ってヘッダー化され、本来の
            //  ヘッダー行(取引先...)を永久に見失うため)
            workbook = Excel.Workbook(content, null),

            // 名前付きテーブル(ListObject)は、その裏にある通常シートの一部の範囲でもあるため、
            // 両方処理すると同じ行が二重に読み込まれる。名前付きテーブルがあればそれを優先し、
            // 見つかった場合は通常シートの走査は行わない。
            //
            // さらに、ファイルによっては同じデータを読み込んだ名前付きテーブルが
            // 複数存在する(例: 「_2024_12売上」「_2024_12売上__2」「_2024_12売上__3」のように
            // Excelが同名テーブル作成時につける連番の重複コピー)。これを全部結合すると
            // 同じ売上が2重・3重にカウントされてしまうため、複数見つかった場合は
            // 行数が最も多い(＝最も完全と思われる)1つだけを採用する。
            tableKindSheets = Table.SelectRows(workbook, each [Kind] = "Table"),
            tableKindProcessedAll = List.Select(
                List.Transform(tableKindSheets[Data], each fnProcessSheetData(_, fileName, yearMonth)),
                each _ <> null
            ),
            tableKindProcessed =
                if List.IsEmpty(tableKindProcessedAll) then
                    {}
                else
                    {List.Sort(tableKindProcessedAll, (a, b) => Number.From(Table.RowCount(b)) - Number.From(Table.RowCount(a))){0}},

            sheetKindSheets = Table.SelectRows(workbook, each [Kind] = "Sheet"),
            sheetKindProcessedAll = List.Select(
                List.Transform(sheetKindSheets[Data], each fnProcessSheetData(_, fileName, yearMonth)),
                each _ <> null
            ),
            sheetKindProcessed =
                if List.IsEmpty(sheetKindProcessedAll) then
                    {}
                else
                    {List.Sort(sheetKindProcessedAll, (a, b) => Number.From(Table.RowCount(b)) - Number.From(Table.RowCount(a))){0}},

            validTables = if not List.IsEmpty(tableKindProcessed) then tableKindProcessed else sheetKindProcessed
        in
            if List.IsEmpty(validTables) then
                error Error.Record("HeaderNotFound", "ヘッダー行(取引先)がどのシート/テーブルにも見つかりません: " & fileName)
            else
                Table.Combine(validTables),

    // フォルダ配下(サブフォルダ含む)の .xlsx を取得。一時ファイル(~$始まり)は除外
    SourceFiles = Folder.Files(RootFolder),
    XlsxFiles = Table.SelectRows(
        SourceFiles,
        each
            Text.EndsWith([Extension], ".xlsx")
            and not Text.StartsWith([Name], "~$")
    ),

    WithParsedTable = Table.AddColumn(
        XlsxFiles,
        "ParsedTable",
        each fnExtractSalesTable([Content], [Name]),
        type table
    ),

    // 期によって列構成が違っても Table.Combine が自動で和集合を取り、欠落列をnull補完する
    Combined = Table.Combine(WithParsedTable[ParsedTable]),

    // 金額・数量列にも、注記などが紛れ込んで数値変換できないセルがまれにある。
    // 日付と同様、変換に失敗した値は問答無用でnullにする(型変換エラーで行ごと壊れるのを防ぐ)。
    fnCleanNumericValue = (v as any) as nullable number =>
        if v = null then
            null
        else
            try Number.From(v, "ja-JP") otherwise null,

    // 主要な金額・数量列の型を、存在するものだけ数値化(期によって無い列があるため存在チェック)
    NumericColumnCandidates = {
        "売上数量", "価格1", "価格2", "掛率", "売上金額",
        "税抜売上計", "消費税", "税込売上計", "税抜仕入計"
    },
    ExistingNumericCols = List.Intersect({Table.ColumnNames(Combined), NumericColumnCandidates}),
    CleanedNumericValues = Table.TransformColumns(
        Combined,
        List.Transform(ExistingNumericCols, each {_, fnCleanNumericValue})
    ),
    TypedNumeric = Table.TransformColumnTypes(
        CleanedNumericValues,
        List.Transform(ExistingNumericCols, each {_, type number})
    ),

    // 一部の月(特に2021年9月前後)は「税抜売上計」「税込売上計」が空欄で、
    // 代わりに「売上金額」(数量×単価×掛率)に実際の売上金額が入っている。
    // これに気づかず放置すると、その月だけ売上が計上されず仕入だけ残るため、
    // 粗利が大きくマイナスになる重大な集計誤りになっていた。
    // 「税抜売上計」が空欄の行だけ「売上金額」で補完し、「税込売上計」も
    // 空欄なら補完後の税抜売上計から10%消費税を足して計算する。
    WithRevenueFallback =
        if List.Contains(Table.ColumnNames(TypedNumeric), "税抜売上計") and List.Contains(Table.ColumnNames(TypedNumeric), "売上金額") then
            let
                fixedExTax = Table.AddColumn(
                    TypedNumeric,
                    "税抜売上計_補完",
                    each
                        let
                            exTax = Record.FieldOrDefault(_, "税抜売上計", null),
                            fallback = Record.FieldOrDefault(_, "売上金額", null)
                        in
                            if exTax <> null then exTax else fallback,
                    type nullable number
                ),
                fixedIncTax =
                    if List.Contains(Table.ColumnNames(TypedNumeric), "税込売上計") then
                        Table.AddColumn(
                            fixedExTax,
                            "税込売上計_補完",
                            each
                                let
                                    incTax = Record.FieldOrDefault(_, "税込売上計", null),
                                    exTaxFixed = Record.FieldOrDefault(_, "税抜売上計_補完", null)
                                in
                                    if incTax <> null then incTax else if exTaxFixed <> null then exTaxFixed * 1.1 else null,
                            type nullable number
                        )
                    else
                        fixedExTax,
                colsToDrop = List.Intersect({Table.ColumnNames(fixedIncTax), {"税抜売上計", "税込売上計"}}),
                dropped = Table.RemoveColumns(fixedIncTax, colsToDrop),
                renamed = Table.RenameColumns(
                    dropped,
                    List.Select(
                        {{"税抜売上計_補完", "税抜売上計"}, {"税込売上計_補完", "税込売上計"}},
                        each List.Contains(Table.ColumnNames(dropped), _{0})
                    )
                )
            in
                renamed
        else
            TypedNumeric,

    // OP日・伝票日付・売上日は、手術がない明細行(運送料・器械貸出料など)では
    // 「1/0」のようなダミー値や、日付として解釈できない文字列が入っていることがある。
    // 個別のパターンを想定して除外するのではなく、実際に日付へ変換を試みて、
    // 失敗した値は問答無用でnullにする(型変換エラーで行ごと壊れるのを防ぐ)。
    //
    // 加えて、変換自体は成功しても「0202/03/08」のような桁抜け入力ミスや、
    // Excelの日付起点由来の「1899/12/30」等、業務上あり得ない日付も混在しているため、
    // 事業開始(2020年)以降・当面の未来(2030年)まで、という妥当な範囲外の値もnullにする。
    // (売上金額自体は残るので合計には影響せず、時系列集計でのみ「不明日付」として扱われる)
    fnCleanDateValue = (v as any) as nullable date =>
        if v = null then
            null
        else
            let
                parsed = try Date.From(v, "ja-JP") otherwise null
            in
                if parsed = null then
                    null
                else if parsed < #date(2020, 1, 1) or parsed > #date(2030, 12, 31) then
                    null
                else
                    parsed,

    // 日付列も、存在するものだけ日付化
    DateColumnCandidates = {"売上日", "OP日", "伝票日付"},
    ExistingDateCols = List.Intersect({Table.ColumnNames(WithRevenueFallback), DateColumnCandidates}),
    CleanedDateValues = Table.TransformColumns(
        WithRevenueFallback,
        List.Transform(ExistingDateCols, each {_, fnCleanDateValue})
    ),
    TypedDates = Table.TransformColumnTypes(
        CleanedDateValues,
        List.Transform(ExistingDateCols, each {_, type date}),
        "ja-JP"
    ),

    // OP日(製品を使用した日=手術日)と売上日(納品書発行日)は意味が異なる別の日付。
    // 本来はOP日を代用すべきではないが、実際のデータでは売上日・伝票日付が
    // 丸ごと欠落している月がかなりの割合であり(2024年10月・2025年5月・6月など)、
    // OP日を完全に除外すると集計そのものが3割近く欠落してしまう。
    // そのため優先順位(売上日 > 伝票日付 > OP日)でOP日を最後の代用として残しつつ、
    // どの日付源を使ったかを別列「取引日_出所」に記録し、後から検証・除外できるようにする。
    WithUnifiedDate = Table.AddColumn(
        Table.AddColumn(
            TypedDates,
            "取引日",
            each
                let
                    candidates = {
                        Record.FieldOrDefault(_, "売上日", null),
                        Record.FieldOrDefault(_, "伝票日付", null),
                        Record.FieldOrDefault(_, "OP日", null)
                    },
                    nonNulls = List.RemoveNulls(candidates)
                in
                    if List.IsEmpty(nonNulls) then null else nonNulls{0},
            type nullable date
        ),
        "取引日_出所",
        each
            if Record.FieldOrDefault(_, "売上日", null) <> null then "売上日"
            else if Record.FieldOrDefault(_, "伝票日付", null) <> null then "伝票日付"
            else if Record.FieldOrDefault(_, "OP日", null) <> null then "OP日(代用)"
            else null,
        type nullable text
    ),

    // 直近の月には、OP(手術)は実施済みだが請求・製品明細の入力がまだ終わっていない
    // 「処理中」の行が残っている(製品名が空欄かつ税抜売上計も0/null)。
    // 売上には計上されていないので分析上はノイズになるため除外する。
    fnIsBlankOrNull = (v as any) as logical => v = null or v = "",
    IsIncompletePendingRow = (r as record) as logical =>
        fnIsBlankOrNull(Record.FieldOrDefault(r, "製品名1", null))
        and fnIsBlankOrNull(Record.FieldOrDefault(r, "製品名2", null))
        and (
            Record.FieldOrDefault(r, "税抜売上計", null) = null
            or Record.FieldOrDefault(r, "税抜売上計", null) = 0
        ),
    ExcludedIncompleteRows = Table.SelectRows(WithUnifiedDate, each not IsIncompletePendingRow(_)),

    // 病院マスタの「コード」は、伝票番号1の下3桁と対応している。
    // 伝票番号1列が存在しない期間の行は病院コードもnullになる。
    fnHospitalCodeFromSlipNo = (v as any) as nullable text =>
        if v = null then
            null
        else
            let
                trimmed = Text.Trim(Text.From(v)),
                padded = Text.PadStart(trimmed, 3, "0")
            in
                if trimmed = "" then null else Text.End(padded, 3),

    // 伝票番号1が無い行向けの代替: Source.Name列の「●xxxx-xxxxx」パターンから
    // 「●」と「-」の間の数字の下3桁を取り出す
    fnHospitalCodeFromSourceName = (v as any) as nullable text =>
        if v = null then
            null
        else
            let
                t = Text.From(v),
                afterMark = if Text.Contains(t, "●") then Text.AfterDelimiter(t, "●") else null,
                beforeDash =
                    if afterMark <> null and Text.Contains(afterMark, "-") then
                        Text.BeforeDelimiter(afterMark, "-")
                    else
                        null,
                digitsOnly = if beforeDash <> null then Text.Select(beforeDash, {"0".."9"}) else null
            in
                if digitsOnly = null or digitsOnly = "" then null else Text.End(Text.PadStart(digitsOnly, 3, "0"), 3),

    WithHospitalCode = Table.AddColumn(
        ExcludedIncompleteRows,
        "病院コード",
        each
            let
                fromSlipNo = fnHospitalCodeFromSlipNo(Record.FieldOrDefault(_, "伝票番号1", null))
            in
                if fromSlipNo <> null then
                    fromSlipNo
                else
                    fnHospitalCodeFromSourceName(Record.FieldOrDefault(_, "Source.Name", null)),
        type nullable text
    ),

    // 病院名の表記ゆれ(前後の空白等)を吸収するための照合用トリム列
    WithTrimmedHospitalName = Table.AddColumn(
        WithHospitalCode,
        "病院名_マッチ用",
        each Text.Trim(Text.From(Record.FieldOrDefault(_, "病院名", ""))),
        type text
    ),

    // 病院マスタ(同じPower BIファイル内の別クエリ)を参照し、
    // ①病院コード一致 → ②(コードで見つからない場合のみ)病院名一致 の優先順で
    // 住所・電話番号・URLを解決して直接マージする
    HospitalMaster = 病院マスタ,

    MergedByCode = Table.NestedJoin(
        WithTrimmedHospitalName, {"病院コード"},
        HospitalMaster, {"コード"},
        "マスタ_コード一致", JoinKind.LeftOuter
    ),
    ExpandedByCode = Table.ExpandTableColumn(
        MergedByCode, "マスタ_コード一致",
        {"住所", "電話番号", "URL"},
        {"住所_コード", "電話番号_コード", "URL_コード"}
    ),

    MergedByName = Table.NestedJoin(
        ExpandedByCode, {"病院名_マッチ用"},
        HospitalMaster, {"名称"},
        "マスタ_名称一致", JoinKind.LeftOuter
    ),
    ExpandedByName = Table.ExpandTableColumn(
        MergedByName, "マスタ_名称一致",
        {"住所", "電話番号", "URL"},
        {"住所_名称", "電話番号_名称", "URL_名称"}
    ),

    fnFirstNonNull = (a as any, b as any) as any => if a <> null then a else b,
    WithResolvedMaster = Table.AddColumn(
        Table.AddColumn(
            Table.AddColumn(
                ExpandedByName,
                "住所",
                each fnFirstNonNull(Record.FieldOrDefault(_, "住所_コード", null), Record.FieldOrDefault(_, "住所_名称", null)),
                type nullable text
            ),
            "電話番号",
            each fnFirstNonNull(Record.FieldOrDefault(_, "電話番号_コード", null), Record.FieldOrDefault(_, "電話番号_名称", null)),
            type nullable text
        ),
        "URL",
        each fnFirstNonNull(Record.FieldOrDefault(_, "URL_コード", null), Record.FieldOrDefault(_, "URL_名称", null)),
        type nullable text
    ),

    FinalCleaned = Table.RemoveColumns(
        WithResolvedMaster,
        {"病院名_マッチ用", "住所_コード", "電話番号_コード", "URL_コード", "住所_名称", "電話番号_名称", "URL_名称"}
    ),

    // 前後・連続する全角/半角スペースを1つの半角スペースに正規化する(null安全)
    fnCollapseWhitespace = (v as any) as nullable text =>
        if v = null then
            null
        else
            let
                t = Text.Replace(Text.From(v), "　", " "),
                collapsed = Text.Combine(List.Select(Text.Split(t, " "), each _ <> ""), " "),
                trimmed = Text.Trim(collapsed)
            in
                if trimmed = "" then null else trimmed,

    // 取引先名から宛名の敬称「御中」を除去してから空白を正規化する(null安全)。
    // 「御中」の有無だけで同じ会社が別の取引先として集計されてしまうのを防ぐ。
    fnCleanDistributorText = (v as any) as nullable text =>
        let
            collapsed = fnCollapseWhitespace(v)
        in
            if collapsed = null then null else fnCollapseWhitespace(Text.Replace(collapsed, "御中", "")),

    // 取引先(代理店)マスタとの突き合わせ:
    //   ①伝票番号1(=貸出先コード。代理店コード2桁+病院コード3桁)が完全一致する行を優先
    //   ②見つからなければ、取引先の表記を正規化した上で代理店略称と一致するものを正式名称に変換
    //   ③どちらも見つからなければ、正規化だけした元の取引先表記を残す(消さない)
    //
    // 注意: 取引先マスタは「代理店×病院」の組み合わせごとに1行あるため、同じ代理店が
    // 何十行にもわたって登場する(名称もコードも一意ではない)。そのままTable.NestedJoinで
    // 突き合わせると、一致した行数だけ売上実績側の行が複製されてしまう重大なバグになるため、
    // ①は一意な「貸出先コード」で、②は代理店略称だけを重複排除した小さい対応表を作ってから
    // 突き合わせる(どちらも売上実績側の行数を絶対に増やさない設計にする)。
    DistributorMaster = 取引先マスタ,

    AbbrevToNameLookup = Table.Distinct(
        Table.SelectColumns(DistributorMaster, {"代理店略称", "代理店名称"}),
        {"代理店略称"}
    ),

    WithSlipNoText = Table.AddColumn(
        FinalCleaned,
        "伝票番号1_テキスト",
        each fnCollapseWhitespace(Record.FieldOrDefault(_, "伝票番号1", null)),
        type nullable text
    ),
    MergedByDistCode = Table.NestedJoin(
        WithSlipNoText, {"伝票番号1_テキスト"},
        DistributorMaster, {"貸出先コード"},
        "マスタ_伝票番号一致", JoinKind.LeftOuter
    ),
    ExpandedByDistCode = Table.ExpandTableColumn(
        MergedByDistCode, "マスタ_伝票番号一致",
        {"代理店名称"},
        {"取引先_伝票一致"}
    ),

    WithTrimmedDistributorName = Table.AddColumn(
        ExpandedByDistCode,
        "取引先_マッチ用",
        each fnCleanDistributorText(Record.FieldOrDefault(_, "取引先", null)),
        type nullable text
    ),
    MergedByDistAbbrev = Table.NestedJoin(
        WithTrimmedDistributorName, {"取引先_マッチ用"},
        AbbrevToNameLookup, {"代理店略称"},
        "マスタ_略称一致", JoinKind.LeftOuter
    ),
    ExpandedByDistAbbrev = Table.ExpandTableColumn(
        MergedByDistAbbrev, "マスタ_略称一致",
        {"代理店名称"},
        {"取引先_略称一致"}
    ),

    WithResolvedDistributor = Table.AddColumn(
        ExpandedByDistAbbrev,
        "取引先_統一",
        each
            let
                byCode = Record.FieldOrDefault(_, "取引先_伝票一致", null),
                byAbbrev = Record.FieldOrDefault(_, "取引先_略称一致", null)
            in
                if byCode <> null then
                    byCode
                else if byAbbrev <> null then
                    byAbbrev
                else
                    fnCleanDistributorText(Record.FieldOrDefault(_, "取引先", null)),
        type nullable text
    ),

    FinalWithDistributor = Table.RemoveColumns(
        WithResolvedDistributor,
        {"伝票番号1_テキスト", "取引先_伝票一致", "取引先_マッチ用", "取引先_略称一致"}
    ),

    // メーカー名の半角カナを全角カナに統一する(マスタが無いため機械的な正規化のみ)
    HalfToFullKanaCombos = {
        {"ｶﾞ", "ガ"}, {"ｷﾞ", "ギ"}, {"ｸﾞ", "グ"}, {"ｹﾞ", "ゲ"}, {"ｺﾞ", "ゴ"},
        {"ｻﾞ", "ザ"}, {"ｼﾞ", "ジ"}, {"ｽﾞ", "ズ"}, {"ｾﾞ", "ゼ"}, {"ｿﾞ", "ゾ"},
        {"ﾀﾞ", "ダ"}, {"ﾁﾞ", "ヂ"}, {"ﾂﾞ", "ヅ"}, {"ﾃﾞ", "デ"}, {"ﾄﾞ", "ド"},
        {"ﾊﾞ", "バ"}, {"ﾋﾞ", "ビ"}, {"ﾌﾞ", "ブ"}, {"ﾍﾞ", "ベ"}, {"ﾎﾞ", "ボ"},
        {"ﾊﾟ", "パ"}, {"ﾋﾟ", "ピ"}, {"ﾌﾟ", "プ"}, {"ﾍﾟ", "ペ"}, {"ﾎﾟ", "ポ"},
        {"ｳﾞ", "ヴ"}
    },
    HalfToFullKanaSingles = {
        {"ｱ", "ア"}, {"ｲ", "イ"}, {"ｳ", "ウ"}, {"ｴ", "エ"}, {"ｵ", "オ"},
        {"ｶ", "カ"}, {"ｷ", "キ"}, {"ｸ", "ク"}, {"ｹ", "ケ"}, {"ｺ", "コ"},
        {"ｻ", "サ"}, {"ｼ", "シ"}, {"ｽ", "ス"}, {"ｾ", "セ"}, {"ｿ", "ソ"},
        {"ﾀ", "タ"}, {"ﾁ", "チ"}, {"ﾂ", "ツ"}, {"ﾃ", "テ"}, {"ﾄ", "ト"},
        {"ﾅ", "ナ"}, {"ﾆ", "ニ"}, {"ﾇ", "ヌ"}, {"ﾈ", "ネ"}, {"ﾉ", "ノ"},
        {"ﾊ", "ハ"}, {"ﾋ", "ヒ"}, {"ﾌ", "フ"}, {"ﾍ", "ヘ"}, {"ﾎ", "ホ"},
        {"ﾏ", "マ"}, {"ﾐ", "ミ"}, {"ﾑ", "ム"}, {"ﾒ", "メ"}, {"ﾓ", "モ"},
        {"ﾔ", "ヤ"}, {"ﾕ", "ユ"}, {"ﾖ", "ヨ"},
        {"ﾗ", "ラ"}, {"ﾘ", "リ"}, {"ﾙ", "ル"}, {"ﾚ", "レ"}, {"ﾛ", "ロ"},
        {"ﾜ", "ワ"}, {"ｦ", "ヲ"}, {"ﾝ", "ン"},
        {"ｧ", "ァ"}, {"ｨ", "ィ"}, {"ｩ", "ゥ"}, {"ｪ", "ェ"}, {"ｫ", "ォ"},
        {"ｬ", "ャ"}, {"ｭ", "ュ"}, {"ｮ", "ョ"},
        {"ｯ", "ッ"}, {"ｰ", "ー"},
        {"｡", "。"}, {"｢", "「"}, {"｣", "」"}, {"､", "、"}, {"･", "・"}
    },
    fnHalfToFullKana = (v as any) as text =>
        let
            t = if v = null then "" else Text.From(v),
            step1 = List.Accumulate(HalfToFullKanaCombos, t, (state, pair) => Text.Replace(state, pair{0}, pair{1})),
            step2 = List.Accumulate(HalfToFullKanaSingles, step1, (state, pair) => Text.Replace(state, pair{0}, pair{1}))
        in
            step2,
    fnNormalizeMakerName = (v as any) as nullable text =>
        if v = null then
            null
        else
            fnCollapseWhitespace(fnHalfToFullKana(v)),

    MakerColumnCandidates = {"メーカー①", "メーカー②", "メーカー③"},
    ExistingMakerCols = List.Intersect({Table.ColumnNames(FinalWithDistributor), MakerColumnCandidates}),
    FinalWithMaker = Table.TransformColumns(
        FinalWithDistributor,
        List.Transform(ExistingMakerCols, each {_, fnNormalizeMakerName})
    )
in
    FinalWithMaker
