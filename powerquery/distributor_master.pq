// ============================================================================
// 代理店(取引先)・病院 組み合わせマスタ 読み込み用 Power Query M スクリプト
//
// 対象: C:\Users\sfuru\OneDrive - M'z company\07. 売上関連\代理店_病院マスタ.xlsx
//
// 使い方:
//   1. Power BI Desktop > Power Query エディター > 新しいクエリ > 空のクエリ
//   2. クエリ名を「取引先マスタ」に変更
//   3. 詳細エディターにこの中身を貼り付けて「完了」
//
// 「貸出先コード」列(例:"18001")は 代理店コード(2桁) + 病院コード(3桁) の組み合わせで、
// これは売上実績の「伝票番号1」の値そのものと一致する。伝票番号1が存在する行は、
// このマスタ1件と直接一致させることで、取引先(代理店名称)・病院名称・病院住所を
// まとめて正確に取得できる。
// ============================================================================

let
    FilePath = "C:\Users\sfuru\OneDrive - M'z company\07. 売上関連\代理店_病院マスタ.xlsx",
    Source = Excel.Workbook(File.Contents(FilePath), null),
    Sheet1 = Source{[Item = "Sheet1", Kind = "Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(Sheet1, [PromoteAllScalars = true]),

    // 伝票番号1・売上実績の「病院名」「取引先」との突き合わせで使うため、
    // キーになる列はテキスト型に統一し、前後の全角/半角スペースを除去しておく
    TrimmedColumns = Table.TransformColumns(
        PromotedHeaders,
        {
            {"貸出先コード", each Text.Trim(Text.From(_)), type text},
            {"代理店コード", each Text.Trim(Text.From(_)), type text},
            {"代理店略称", each Text.Trim(Text.From(_)), type text},
            {"代理店名称", each Text.Trim(Text.From(_)), type text},
            {"病院コード", each Text.Trim(Text.From(_)), type text},
            {"病院名称", each Text.Trim(Text.From(_)), type text}
        }
    ),

    ChangedType = Table.TransformColumnTypes(
        TrimmedColumns,
        {
            {"郵便番号", type text},
            {"住所１", type text},
            {"住所２", type text},
            {"Mz担当", type text},
            {"電話番号", type text},
            {"FAX", type text},
            {"病院住所", type text},
            {"病院電話番号", type text},
            {"病院URL", type text}
        }
    ),

    // 「貸出先コード」は代理店×病院の組み合わせごとに1件のはずなので、
    // 万一の重複に備えて1件に絞る
    Deduplicated = Table.Distinct(ChangedType, {"貸出先コード"})
in
    Deduplicated
