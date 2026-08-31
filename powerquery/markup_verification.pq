// ============================================================================
// 仕入掛率・納入掛率 検算用 Power Query M スクリプト
//
// 目的:
//   「価格1」=納入価、「価格2」=定価、「掛率」=納入価/定価(納入掛率)であることは
//   確認済み。一方「仕入掛率」(定価に対する仕入値の割合)は明細行単位の列が存在せず、
//   「税抜仕入計」が症例(伝票番号1+伝票番号2)単位の値であるため単純には出せない。
//   このクエリは症例単位で「定価合計」「税抜仕入計」「税抜売上計」を集計し、
//   そこから仕入掛率・納入掛率を逆算して1症例1行の検算用テーブルを作る。
//   出来上がった数値を、知っている症例と見比べて実態と合っているか確認するための
//   診断用クエリであり、レポートのビジュアルには使わない想定。
//
// 前提:
//   同じPower BIファイル内に「売上実績クエリ」(sales_data_clean.pq)が
//   既に作成・実行済みであること
//
// 使い方:
//   1. Power Query エディター > 新しいクエリ > 空のクエリ
//   2. 詳細エディターにこのファイルの中身を丸ごと貼り付け、クエリ名を
//      「仕入掛率_検算」などに変更して更新
//   3. できたテーブルを「掛率差(納入-仕入)」や「売上金額合計_vs_税抜売上計差」で
//      並べ替え、値が大きくずれている症例を個別に確認する
//
// 読み方の注意:
//   - 「定価合計」は明細行の「定価」(製品マスタ由来)を症例内で合算したもの。
//     合計行(税抜仕入計・税抜売上計が入っている行)自体には製品番号が無いことが
//     多いため定価はnullになり、合算には影響しない想定。
//   - 「税抜仕入計」「税抜売上計」はsales_data_clean.pq側の修正により、症例内で
//     合計行にしか値が入っていない(明細行はnull化済み)ため、List.Sumで
//     合算しても実質的にその1件の値がそのまま採用される。
//   - 「仕入掛率_検算」「納入掛率_検算」は症例全体をひとまとめにした平均的な
//     掛け率であり、症例内に複数製品がある場合、製品ごとの内訳までは分からない。
//   - 「売上金額合計_vs_税抜売上計差」がゼロから大きく離れる症例があれば、
//     売上金額(明細行の内訳額)と税抜売上計(症例の合計額)の対応関係自体を
//     再確認する必要がある。
// ============================================================================

let
    Source = 売上実績クエリ,

    // 症例キー: sales_data_clean.pq の fnMarkLineItems と同じグループ化基準
    // (伝票番号1 + 伝票番号2 + ソースファイル)
    Grouped = Table.Group(
        Source,
        {"伝票番号1", "伝票番号2", "ソースファイル"},
        {
            {"年月", each List.Max(List.RemoveNulls(Table.Column(_, "年月"))), type nullable text},
            {"病院名_統一", each List.Max(List.RemoveNulls(Table.Column(_, "病院名_統一"))), type nullable text},
            {"代理店_統一", each List.Max(List.RemoveNulls(Table.Column(_, "代理店_統一"))), type nullable text},
            {"明細行数", each Table.RowCount(_), Int64.Type},
            {
                "定価合計",
                each
                    let vals = List.RemoveNulls(Table.Column(_, "定価"))
                    in if List.IsEmpty(vals) then null else List.Sum(vals),
                type nullable number
            },
            {
                "売上金額合計",
                each
                    let vals = List.RemoveNulls(Table.Column(_, "売上金額"))
                    in if List.IsEmpty(vals) then null else List.Sum(vals),
                type nullable number
            },
            {
                "税抜仕入計",
                each
                    let vals = List.RemoveNulls(Table.Column(_, "税抜仕入計"))
                    in if List.IsEmpty(vals) then null else List.Sum(vals),
                type nullable number
            },
            {
                "税抜売上計",
                each
                    let vals = List.RemoveNulls(Table.Column(_, "税抜売上計"))
                    in if List.IsEmpty(vals) then null else List.Sum(vals),
                type nullable number
            },
            {
                "掛率_単純平均",
                each
                    let vals = List.RemoveNulls(Table.Column(_, "掛率"))
                    in if List.IsEmpty(vals) then null else List.Average(vals),
                type nullable number
            }
        }
    ),

    fnSafeDivide = (numerator as nullable number, denominator as nullable number) as nullable number =>
        if numerator = null or denominator = null or denominator = 0 then
            null
        else
            numerator / denominator,

    WithRatios = Table.AddColumn(
        Table.AddColumn(
            Grouped,
            "仕入掛率_検算",
            each fnSafeDivide([税抜仕入計], [定価合計]),
            type nullable number
        ),
        "納入掛率_検算",
        each fnSafeDivide([税抜売上計], [定価合計]),
        type nullable number
    ),

    WithDiffs = Table.AddColumn(
        Table.AddColumn(
            WithRatios,
            "掛率差(納入-仕入)",
            each
                if [納入掛率_検算] = null or [仕入掛率_検算] = null then
                    null
                else
                    [納入掛率_検算] - [仕入掛率_検算],
            type nullable number
        ),
        "売上金額合計_vs_税抜売上計差",
        each
            if [売上金額合計] = null or [税抜売上計] = null then
                null
            else
                [売上金額合計] - [税抜売上計],
        type nullable number
    )
in
    WithDiffs
