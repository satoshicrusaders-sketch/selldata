// ============================================================================
// 医療機関マスタ 読み込み用 Power Query M スクリプト
//
// 対象: C:\Users\sfuru\OneDrive - M'z company\07. 売上関連\病院マスタ_更新済み.xlsx
//
// 使い方:
//   1. Power BI Desktop > Power Query エディター > 新しいクエリ > 空のクエリ
//   2. クエリ名を「病院マスタ」に変更
//   3. 詳細エディターにこの中身を貼り付けて「完了」
//
// 売上実績クエリとは病院名をキーにモデル画面でリレーションを作る。
// (この病院マスタは1ファイルだけを読むので、月次一覧フォルダの
//  再帰スキャン対象には含めていない)
// ============================================================================

let
    FilePath = "C:\Users\sfuru\OneDrive - M'z company\07. 売上関連\病院マスタ_更新済み.xlsx",
    // useHeaders は null にしておく。true にすると Excel.Workbook がここで先に
    // 1行目をヘッダー昇格してしまい、直後の Table.PromoteHeaders と二重昇格になって
    // 本来のヘッダーの次の行(実データ1件目)が消えてしまう。
    Source = Excel.Workbook(File.Contents(FilePath), null),
    Sheet1 = Source{[Item = "Sheet1", Kind = "Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(Sheet1, [PromoteAllScalars = true]),

    // 売上実績側の「病院名」と結合するキーになるため、
    // 全角/半角スペースの混入によるズレを防ぐためにトリムしておく
    TrimmedName = Table.TransformColumns(
        PromotedHeaders,
        {
            {"名称", each Text.Trim(Text.From(_)), type text},
            {"正式名称", each Text.Trim(Text.From(_)), type text}
        }
    ),

    ChangedType = Table.TransformColumnTypes(
        TrimmedName,
        {
            {"コード", type text},
            {"名称", type text},
            {"正式名称", type text},
            {"住所", type text},
            {"電話番号", type text},
            {"URL", type text}
        }
    ),

    // 万一「名称」が完全に同じ行が複数あると、後でリレーションが1対多にならず
    // エラーになるため、重複があれば1件だけに絞る
    Deduplicated = Table.Distinct(ChangedType, {"名称"})
in
    Deduplicated
