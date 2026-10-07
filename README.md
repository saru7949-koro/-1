# -1
営業案件の分析メモと資料をまとめるリポジトリ

## パイプライン不足分析(代理店軸)

上期▲7台の挽回と今期パイプライン不足(9F時点の理論値21件)の解消に向け、代理店別の分析と2〜4週間のアクションを出す。

```
原本/                 元ファイル(読み取り専用・git管理外)
作業/snapshots/<日付>/ 実行時に原本をコピーしたもの(分析はこのコピーのみ。git管理外)
出力/                 分析結果(Excel / Markdown。施設名・個人名は含めない)
config.toml           ファイル名・期間・算出式パラメータ・代理店表記ゆれ
scripts/run.py        スナップショット → 分析 → 出力
```

実行:

```
python3 scripts/run.py 2026-10-06
```

- 原本ファイルを更新するときは `原本/` に置き、`config.toml` の `[source]` のファイル名を書き換える
- 21件の算出式(Sep資料 slide2): 着地理論値 = 成約 + FCT×モデル勝率 + Backup×10%、必要リード = 不足 ÷ モデル勝率
- 仮置きの前提は 出力 Excel の「00_前提と算出式」シートに一覧がある

役員報告スライド(3枚)の作り直し:

```
python3 scripts/deck_metrics.py 2026-10-07            # 出力/deck_metrics_20261007.json を作る
NODE_PATH=<pptxgenjsのnode_modules> node scripts/build_exec_deck.js 出力/役員報告_パイプライン挽回計画_20261007.pptx 出力/deck_metrics_20261007.json
```

- 新しいFCT版の案件Excelを受け取ったら、`原本/` に置いて `config.toml` の `pipeline_latest`・`latest_label`・`as_of` を書き換え、上の手順を実行する
- 区分「nFCT」は、確度100%か上期の案件を成約、それ以外(10FCTで計上された下期案件など)を未成約のFCTとして扱う
