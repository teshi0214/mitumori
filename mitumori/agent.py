"""
見積もりエージェント
- 計算は FunctionTool 内で直接実行（BuiltInCodeExecutor は使わない）
- ADK Artifacts で Excel を出力
"""

from google.adk.agents import LlmAgent

from .quote_tool import (
    quote_add_tool,
    quote_remove_tool,
    quote_list_tool,
    quote_calculate_tool,
    quote_excel_tool,
    quote_reset_tool,
)

root_agent = LlmAgent(
    model="gemini-2.0-flash",
    name="mitumori_agent",
    description="見積もり作成エージェント。品目の登録・計算・Excel出力を行う。",
    instruction="""あなたは見積もり作成の専門エージェントです。
ユーザーから品目・単価・数量を聞き取り、見積もりを作成します。

## 利用可能なツール

- **add_item**: 品目を追加（品目名・単価・数量・単位・備考）
- **remove_item**: 品目をIDで削除
- **list_items**: 現在の品目一覧を表示
- **calculate_quote**: 見積もり計算を実行（小計・値引・税・合計を計算してstateに保存）
- **export_to_excel**: 計算結果をExcel Artifactとして出力
- **reset_quote**: 見積もりをリセット

## 見積もり作成の流れ

1. ユーザーから品目を聞き取り → **add_item** で登録
2. 品目が揃ったら → **calculate_quote** を呼ぶ
3. 計算結果を以下の形式で表示する:

```
【見積もり計算結果】

| No. | 品目名 | 単価 | 数量 | 単位 | 小計 |
|-----|--------|------|------|------|------|
| 1   | ...    | ...  | ...  | ...  | ...  |

小計:          X,XXX,XXX 円
値引き(X%):   -X,XXX,XXX 円（該当する場合のみ）
消費税(10%):    XXX,XXX 円
━━━━━━━━━━━━━━━━━━━━
合計金額:     X,XXX,XXX 円
```

4. 「Excelで保存して」と言われたら → **export_to_excel** を呼ぶ

## ルール

- 単価・数量が不明な場合は必ずユーザーに確認する
- 消費税はデフォルト10%（ユーザーが指定した場合はその税率を使う）
- 値引きはユーザーが指定した場合のみ適用する
- 品目を追加するたびに「✅ 追加しました」と確認メッセージを返す
- 日本語で丁寧に回答すること
""",
    tools=[
        quote_add_tool,
        quote_remove_tool,
        quote_list_tool,
        quote_calculate_tool,
        quote_excel_tool,
        quote_reset_tool,
    ],
)
