# scriptsフォルダ一覧

場所：`03_コンテンツ＆マーケティング部\AI-note\scripts\`

note投稿・画像生成を自動化するスクリプト群。
設定方法は `03_コンテンツ＆マーケティング部\AI-note\config\README.md` を参照。

---

## スクリプト一覧

| ファイル | 役割 | 使い方 |
|---|---|---|
| `generate_images.py` | DALL-E 3でマンガ画像を生成する | `python scripts/generate_images.py` |
| `generate_eyecatch.py` | アイキャッチ画像を生成する | `python scripts/generate_eyecatch.py` |
| `create_note_draft.py` | note下書きを作成する | `python scripts/create_note_draft.py` |
| `update_note_draft.py` | note下書きに画像を挿入する | `python scripts/update_note_draft.py --note-id [ID]` |
| `post_draft_mac.py` | note下書きを投稿する（Mac用） | Mac環境で使用 |
| `start_chrome.bat` | Chromeを起動するバッチ（Windows用） | ダブルクリックで起動 |

---

## 必要な設定

1. `03_コンテンツ＆マーケティング部\AI-note\.env.local` にAPIキーを設定（中身は非表示・絶対に出力しない）
2. 必要なライブラリのインストール：
   ```
   pip install openai pandas openpyxl playwright python-dotenv
   playwright install chromium
   ```
3. 詳細は `03_コンテンツ＆マーケティング部\AI-note\config\README.md` を参照

---

## 注意事項

- `.env.local` には note.com のパスワードと OpenAI API Key が入っている
- 中身は絶対にコード・チャット上に表示しない
- スクリプトは `03_コンテンツ＆マーケティング部\AI-note\` フォルダをカレントディレクトリにして実行する
- Windows では `python`、Mac では `python3` コマンドを使う
