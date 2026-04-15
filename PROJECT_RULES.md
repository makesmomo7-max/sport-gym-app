# momo fit（ジム用）運用ルール

## このフォルダの役割（正本）
- 対象: スポーツジム会員向けアプリ（デイサービス用は別スレッドで別ファイル予定）
- 主なファイル:
  - `gym.html`（会員向けホーム・入口）
  - `check.html`（日次コンディション記録）
  - `share_trainer.html`（トレーナー共有）
  - `chat.html`（AIチャット）
  - `yoyaku_employee.html`（予約フォーム）
  - `thanks.html`（送信完了画面）
  - `code.gs`（GAS連携）

## 変更ルール
1. 今後の修正は **このフォルダのみ** に実施する
2. `health-app` は参照用（過去版）として扱い、直接運用しない
3. デプロイ前に `yoyaku_employee.html` の `<form id="reserveForm">` が壊れていないか確認する

## 差分取り込み方針（health-app から）
- 採用するのは「文言・UIアイデア」のみ
- 送信処理・連携処理は本フォルダ（momo fit ジム）を優先
- 予約ロジック（時刻制限・検証）は本フォルダを正とする
