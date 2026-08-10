// ============================================================================
// 製品価格マスタ（仕入先価格表統合） 読み込み用 Power Query M スクリプト
//
// 対象: C:\Users\sfuru\OneDrive - M'z company\デスクトップ\Claude\Claudeへ\
//       診療報酬情報\仕入先価格表\<メーカー名>\*.xlsx
//
// 目的（最終ゴール）:
//   ① 償還記号(略称。例:HF-4)で検索  → 現在の償還価格を表示
//   ② 製品番号で検索              → 現在の償還価格・定価を表示
//   ③ メーカー×償還記号で絞り込み  → 該当製品一覧を表示
//   ができる、メーカー横断の「製品価格マスタ」を作る。
//   (①②③はいずれもPower BI側で本クエリの列にフィルタ/スライサーを
//    かけるだけで実現できるので、このクエリ自体は1本の統合テーブルを
//    作ることだけに専念する)
//
// 使い方:
//   1. Power BI Desktop > Power Query エディター > 新しいクエリ > 空のクエリ
//   2. クエリ名を「製品価格マスタ」に変更
//   3. 詳細エディターにこの中身を貼り付けて「完了」
//   4. 下のRootFolderを、実際にOneDriveが同期されているローカルパスに
//      書き換える
//
// 設計方針:
//   メーカーごとに価格表のシート構成・列名がまったく違うため、
//   フォルダ名(仕入先価格表\<メーカー名>\)でメーカーを判定し、
//   メーカーごとの専用処理関数(fnParseXxx)に振り分ける方式にしている。
//   新しいメーカーの価格表を追加する場合は、
//     (a) 仕入先価格表フォルダの下に新しいメーカー名のサブフォルダを作り
//         そこにメーカーの価格表(.xlsx)を保存する
//     (b) このクエリに専用の fnParseXxx 関数を追加し、MakerParsers に
//         1行登録する
//   の2点で対応できる。未対応メーカーのフォルダは自動的にスキップされる
//   (エラーにはならないが製品価格マスタにも載らないので注意)。
//
// 現状の対応状況:
//   ・マイクロポート … 「価格表」シートのみ対応。
//     「価格表_KNEE_EVL」等その他シートは列構成が別のため未対応(要追加)。
//   ・ナカシマ … シート名に「材料」を含むシートのみ対応。
//     「立会い」「レンタル」「ナビゲーション」を含むファイルは
//     特定保険医療材料の償還価格表ではないため、そもそも読み込み対象から
//     除外している。
//
// 既知の注意点(未解決):
//   ナカシマの「材料（その他製品）」と「材料・レンタル（その他製品）」の
//   ように、同じ取得日で似た名前のファイルが複数存在するケースがある。
//   もし同じ製品番号が両方に載っていた場合、後述の重複排除ロジックでは
//   どちらが残るか決め打ちにならない。運用上どちらか一方に統一されて
//   いくことを想定しているが、当面は実データを見ながら要確認。
// ============================================================================

let
    RootFolder = "C:\Users\sfuru\OneDrive - M'z company\デスクトップ\Claude\Claudeへ\診療報酬情報\仕入先価格表",

    // ヘッダー等の全角/半角スペース・改行を除去して比較しやすくする(null安全)
    fnCleanName = (n as any) as text =>
        let
            t = if n = null then "" else Text.From(n),
            noHalfSpace = Text.Replace(t, " ", ""),
            noFullSpace = Text.Replace(noHalfSpace, "　", ""),
            noLineBreak = Text.Replace(Text.Replace(noFullSpace, "#(lf)", ""), "#(cr)", "")
        in
            Text.Trim(noLineBreak),

    // ファイル名の中から「連続する8桁の数字(YYYYMMDD)」を探して日付化する。
    // マイクロポートのファイル名のように"2026.6"や"v3"等、日付以外の数字が
    // 混在していても誤爆しないよう、単純に数字だけ抜き出すのではなく
    // 「8桁連続」かつ「妥当な年月日(2020〜2035年)」の箇所だけを採用する。
    // 見つからない場合はファイルの更新日時にフォールバックする。
    fnExtractDateFromName = (fileName as text, modifiedDate as datetime) as date =>
        let
            len = Text.Length(fileName),
            startPositions = if len < 8 then {} else List.Numbers(0, len - 7),
            isDigitRun = (pos as number) as logical =>
                let
                    candidate = Text.Range(fileName, pos, 8),
                    allDigits = List.AllTrue(List.Transform(
                        Text.ToList(candidate), each List.Contains({"0".."9"}, _)
                    ))
                in
                    allDigits,
            fnTryParse = (pos as number) as nullable date =>
                let
                    candidate = Text.Range(fileName, pos, 8),
                    parsed = try Date.FromText(
                        Text.Range(candidate, 0, 4) & "-" & Text.Range(candidate, 4, 2) & "-" & Text.Range(candidate, 6, 2)
                    ) otherwise null,
                    inRange = parsed <> null and parsed >= #date(2020, 1, 1) and parsed <= #date(2035, 12, 31)
                in
                    if inRange then parsed else null,
            candidatePositions = List.Select(startPositions, isDigitRun),
            candidateDates = List.RemoveNulls(List.Transform(candidatePositions, fnTryParse)),
            found = if List.IsEmpty(candidateDates) then null else candidateDates{0}
        in
            if found <> null then found else DateTime.Date(modifiedDate),

    // --------------------------------------------------------------------
    // マイクロポート「価格表」シート専用パーサ
    //
    // 償還価格は「20XX年XX月　償還価格」という列が改定のたびに増えていく
    // 想定なので、特定の列名をハードコードせず、列名に含まれる数字(年月)を
    // 読み取って一番新しいものを「現在の償還価格」として動的に採用する。
    // --------------------------------------------------------------------
    fnParseMicroPort = (content as binary, fileName as text, modifiedDate as datetime) as table =>
        let
            workbook = Excel.Workbook(content, null),
            sheet = workbook{[Item = "価格表", Kind = "Sheet"]}[Data],
            // 1〜9行目は会社名・住所・単位表記等の見出しのため、10行目をヘッダーに昇格
            skipped = Table.Skip(sheet, 9),
            promoted = Table.PromoteHeaders(skipped, [PromoteAllScalars = true]),
            cleaned = Table.TransformColumnNames(promoted, fnCleanName),

            allCols = Table.ColumnNames(cleaned),
            priceCols = List.Select(allCols, each Text.Contains(_, "償還価格")),
            priceColsWithKey = List.Transform(priceCols, each [name = _, key = Text.Select(_, {"0".."9"})]),
            // 列名から取れる数字(年月)が一番大きい = 一番新しい列を現在価格とみなす
            latestPriceCol =
                if List.IsEmpty(priceColsWithKey) then
                    null
                else
                    List.Last(List.Sort(priceColsWithKey, (a, b) => Number.From(a[key]) - Number.From(b[key])))[name],

            renamed =
                if latestPriceCol = null then
                    cleaned
                else
                    Table.RenameColumns(cleaned, {{latestPriceCol, "償還価格_現在"}}, MissingField.Ignore),

            withMeta = Table.AddColumn(
                Table.AddColumn(
                    Table.AddColumn(renamed, "メーカー", each "マイクロポート・オーソペディックス", type text),
                    "取得日", each fnExtractDateFromName(fileName, modifiedDate), type date
                ),
                "ソースファイル", each fileName, type text
            ),

            neededCols = {
                "メーカー", "製品番号", "販売名", "製品名", "販売価格",
                "償還価格_現在", "特定保険医療材料名1", "特定保険医療材料名2",
                "取得日", "ソースファイル"
            },
            existingNeeded = List.Intersect({Table.ColumnNames(withMeta), neededCols}),
            selected = Table.SelectColumns(withMeta, existingNeeded),

            standardized = Table.RenameColumns(
                selected,
                {
                    {"販売名", "品名"},
                    {"製品名", "規格"},
                    {"販売価格", "定価"},
                    {"特定保険医療材料名1", "償還記号"},
                    {"特定保険医療材料名2", "償還記号2"}
                },
                MissingField.Ignore
            )
        in
            standardized,

    // --------------------------------------------------------------------
    // ナカシマ「材料(◯◯)」シート専用パーサ
    //
    // 新/旧の2値しか持たないため、「新償還価格」＝そのファイル取得時点での
    // 現在価格とみなす。複数カテゴリ(股関節・膝関節等)が別ファイルに
    // 分かれているため、1ファイル内の「材料」を含む全シートを結合する。
    // --------------------------------------------------------------------
    fnParseNakashima = (content as binary, fileName as text, modifiedDate as datetime) as table =>
        let
            workbook = Excel.Workbook(content, null),
            targetSheets = Table.SelectRows(workbook, each Text.Contains([Name], "材料")),
            parsedSheets = List.Select(
                List.Transform(
                    targetSheets[Data],
                    each
                        let
                            skipped = Table.Skip(_, 3),
                            promoted = Table.PromoteHeaders(skipped, [PromoteAllScalars = true]),
                            cleaned = Table.TransformColumnNames(promoted, fnCleanName)
                        in
                            cleaned
                ),
                each Table.RowCount(_) > 0
            ),
            combined = if List.IsEmpty(parsedSheets) then null else Table.Combine(parsedSheets),

            result =
                if combined = null then
                    null
                else
                    let
                        withMeta = Table.AddColumn(
                            Table.AddColumn(
                                Table.AddColumn(combined, "メーカー", each "ナカシマ", type text),
                                "取得日", each fnExtractDateFromName(fileName, modifiedDate), type date
                            ),
                            "ソースファイル", each fileName, type text
                        ),
                        neededCols = {
                            "メーカー", "品番", "商品名", "サイズ等",
                            "新希望小売価格", "新償還価格", "償還分類番号",
                            "取得日", "ソースファイル"
                        },
                        existingNeeded = List.Intersect({Table.ColumnNames(withMeta), neededCols}),
                        selected = Table.SelectColumns(withMeta, existingNeeded)
                    in
                        Table.RenameColumns(
                            selected,
                            {
                                {"品番", "製品番号"},
                                {"商品名", "品名"},
                                {"サイズ等", "規格"},
                                {"新希望小売価格", "定価"},
                                {"新償還価格", "償還価格_現在"},
                                {"償還分類番号", "償還記号"}
                            },
                            MissingField.Ignore
                        )
        in
            result,

    // --------------------------------------------------------------------
    // フォルダ走査(メーカー名サブフォルダを含めて再帰的に取得)
    // --------------------------------------------------------------------
    SourceFiles = Folder.Files(RootFolder),
    XlsxFiles = Table.SelectRows(
        SourceFiles,
        each
            Text.EndsWith([Extension], ".xlsx")
            and not Text.StartsWith([Name], "~$")
            // 立会い料・器械レンタル料・ナビゲーション(目次)は
            // 特定保険医療材料の償還価格表ではないため対象外
            and not Text.Contains([Name], "立会い")
            and not Text.Contains([Name], "レンタル")
            and not Text.Contains([Name], "ナビゲーション")
    ),

    // フォルダパスの「仕入先価格表\」の直後の1階層目 = メーカー名とみなす
    WithMaker = Table.AddColumn(
        XlsxFiles, "メーカー_フォルダ",
        each
            let
                afterRoot = Text.AfterDelimiter([Folder Path], RootFolder & "\"),
                firstSeg = Text.BeforeDelimiter(afterRoot, "\")
            in
                if firstSeg = "" then Text.TrimEnd(afterRoot, "\") else firstSeg,
        type text
    ),

    // メーカー名(フォルダ名) → 専用パーサ の対応表。
    // 新しいメーカーを追加する場合はここに1行足すだけでよい。
    MakerParsers = [
        マイクロポート = fnParseMicroPort,
        ナカシマ = fnParseNakashima
    ],

    WithParsedTable = Table.AddColumn(
        WithMaker, "ParsedTable",
        each
            let
                parser = Record.FieldOrDefault(MakerParsers, [メーカー_フォルダ], null)
            in
                if parser = null then
                    null // 未対応メーカー(専用パーサ未実装)は今のところスキップする
                else
                    try parser([Content], [Name], [Date modified]) otherwise null,
        type nullable table
    ),

    ValidParsed = Table.SelectRows(WithParsedTable, each [ParsedTable] <> null and Table.RowCount([ParsedTable]) > 0),
    Combined = if Table.RowCount(ValidParsed) = 0 then #table({}, {}) else Table.Combine(ValidParsed[ParsedTable]),

    // 製品番号は結合キーになるため、前後の全角/半角スペースを除去してテキスト型に統一
    TrimmedKeys = Table.TransformColumns(
        Combined,
        {
            {"製品番号", each Text.Trim(Text.From(_)), type text},
            {"償還記号", each if _ = null then null else Text.Trim(Text.From(_)), type nullable text}
        }
    ),

    // 同じメーカー×製品番号が複数ファイル(=複数取得日)に重複して存在する
    // 場合、取得日が一番新しいものだけを残す(改定のたびにファイルが
    // 増えていく運用のため、最新版だけが「製品価格マスタ」の1行になる)
    // Table.Last(...)はレコード型(1行)を返してしまいTable.Combineと組み合わせられないため、
    // テーブル型のまま「末尾1行」を取れるTable.LastNを使う
    Grouped = Table.Group(
        TrimmedKeys, {"メーカー", "製品番号"},
        {{"最新行", each Table.LastN(Table.Sort(_, {{"取得日", Order.Ascending}}), 1), type table}}
    ),
    Latest = Table.Combine(Grouped[最新行])
in
    Latest
