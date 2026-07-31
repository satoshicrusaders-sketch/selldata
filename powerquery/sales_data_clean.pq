// ============================================================================
// 販売実績（月次一覧）スクリーニング用 Power Query M スクリプト
//
// 対象: OneDrive for Business / SharePoint 上の
//   07. 売上関連 / 月次一覧 - コピー / [8期|9期|10期|11期|12期|13期] / YYYY.MM.xlsx (or YYYYMM.xlsx)
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
    ExistingDateCols = List.Intersect({Table.ColumnNames(TypedNumeric), DateColumnCandidates}),
    CleanedDateValues = Table.TransformColumns(
        TypedNumeric,
        List.Transform(ExistingDateCols, each {_, fnCleanDateValue})
    ),
    TypedDates = Table.TransformColumnTypes(
        CleanedDateValues,
        List.Transform(ExistingDateCols, each {_, type date}),
        "ja-JP"
    ),

    // 期によって「売上日」列自体が存在しない期間があるため(11-12期あたりはOP日しかない)、
    // 3つの日付候補から存在するものを優先順(売上日 > 伝票日付 > OP日)で1本化した
    // 分析用の日付列を追加する。時系列の軸にはこちらを使うことを推奨。
    WithUnifiedDate = Table.AddColumn(
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
    ExcludedIncompleteRows = Table.SelectRows(WithUnifiedDate, each not IsIncompletePendingRow(_))
in
    ExcludedIncompleteRows
