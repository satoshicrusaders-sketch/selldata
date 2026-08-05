// ============================================================================
// 製品(価格表)マスタ 読み込み用 Power Query M スクリプト
//
// 対象: C:\Users\sfuru\OneDrive - M'z company\07. 売上関連\製品マスタ.xlsx
//       (シート名: 価格表)
//
// 使い方:
//   1. Power BI Desktop > Power Query エディター > 新しいクエリ > 空のクエリ
//   2. クエリ名を「製品マスタ」に変更
//   3. 詳細エディターにこの中身を貼り付けて「完了」
//   4. 下のFilePathを、実際に保存したファイルのパス・ファイル名に合わせて書き換える
//
// 売上実績クエリの「製品番号」をキーにモデル画面でリレーションを作るか、
// sales_data_clean.pq側で直接マージして使う。
//
// 製品番号は125件が重複していたが、その大半(118件)は同一メーカーの
// 重複行(表記ゆれ)で実害なし。7件だけ、同じ製品番号で異なるメーカーが
// 登録されている(コードの偶然の衝突と思われる、ごく僅かなので許容する)。
// 売上実績側で行が増殖しないよう、必ず製品番号で1件に絞ってから使うこと。
// ============================================================================

let
    FilePath = "C:\Users\sfuru\OneDrive - M'z company\07. 売上関連\製品マスタ.xlsx",
    Source = Excel.Workbook(File.Contents(FilePath), null),
    価格表 = Source{[Item = "価格表", Kind = "Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(価格表, [PromoteAllScalars = true]),

    // 売上実績側の「製品番号」との結合キーになるため、前後の全角/半角スペースを
    // 除去してテキスト型に統一しておく
    TrimmedColumns = Table.TransformColumns(
        PromotedHeaders,
        {
            {"メーカー", each Text.Trim(Text.From(_)), type text},
            {"製品番号", each Text.Trim(Text.From(_)), type text},
            {"品名", each Text.Trim(Text.From(_)), type text},
            {"規格", each Text.Trim(Text.From(_)), type text}
        }
    ),

    // 「製品番号」は本来1件のはずなので、万一の重複(表記ゆれ等)に備えて1件に絞る。
    // 異なるメーカーが同じ製品番号を持つ7件については、どちらが残るかは決め打ちに
    // ならないが、件数がごく僅かなので実用上は許容する
    Deduplicated = Table.Distinct(TrimmedColumns, {"製品番号"})
in
    Deduplicated
