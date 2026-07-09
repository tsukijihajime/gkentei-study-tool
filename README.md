# G検定 学習ツール（GitHub Pages公開版）

`G検定/G検定_学習ツール/`（社内配布用）とは別に、GitHub Pagesで一般公開することを想定して作成したバージョン。用語の説明文（desc）を市販の参考書の文章表現から独立させ、著作権リスクを解消してある。

## 社内配布版との違い

| | 社内配布版 | 本フォルダ（公開版） |
|---|---|---|
| 想定公開範囲 | 社内のみ（社外秘） | GitHub Pagesで一般公開 |
| 用語のdesc（説明文） | 一部が市販参考書の文章表現と一致 | 84語のdescを独自表現に書き直し済み |
| トップページ | 「社外秘」バッジあり | なし、代わりに免責事項を明記 |

## 著作権対応の経緯

用語集は市販の参考書『徹底攻略 G検定 ジェネラリスト問題集 第3版』（インプレス、2024年8月発行）の索引を参考に用語を選定して作成した。用語名（見出し語）自体は一般的な技術用語であり著作物ではないが、**説明文（desc）の一部が同書の文章表現と一致していた**ことが判明したため、以下の対応を行った。

1. 黒本PDF全ページのテキストと全659語のdescについて、最長共通部分文字列（LCS）を機械的に検出（`difflib.SequenceMatcher`）
2. 20文字以上の日本語の連続一致があった84語を特定
3. 元の技術的な意味・正確性を保ったまま、参考書の言い回しを使わない独自の説明文に書き直し（Agentによる並列作業＋人手での再検証）
4. 書き直し後に再度LCS検証を行い、日本語の一致が解消されたことを確認（残る一致は英語の技術用語名のみで、これは著作物ではないため問題ない）
5. 記述問題集（`questions.json`、166問）についても同様の検証を行い、解説文中で一致していた箇所を書き直し済み

検証・書き直しの詳細な作業ログは社内環境の `_scratch/gkentei/` に残っている（本フォルダには同梱していない）。

**免責事項**：上記の対応は機械的な検証と人手（AI支援）による書き直しに基づくものであり、著作権侵害が完全に存在しないことを法的に保証するものではない。公開前に法務レビューを受けることを推奨する。

## 収録ファイル

| ファイル | 内容 |
|---|---|
| `index.html` | トップページ |
| `glossary.html` | 用語集検索サイト（659語） |
| `quiz.html` | 4択クイズ（説明文→用語当て） |
| `exam.html` | 記述問題集（正誤判定・空欄穴埋め、166問） |
| `public_data.json` | 用語集データ本体（84語のdescを書き直し済み） |
| `questions.json` | 記述問題集の設問データ |
| `gkentei_glossary_public.xlsx` | 用語集xlsx（`public_data.json`から生成） |
| `build_glossary_public.py` | `public_data.json` → xlsx |
| `build_glossary_site.py` | xlsx → `glossary.html` |
| `build_quiz.py` | xlsx → `quiz.html` |
| `build_exam.py` | `questions.json` → `exam.html` |

## ビルド手順

用語集や問題を追加・修正した場合は、以下の順で再生成する。

```powershell
cd G検定\GitHub公開版
python build_glossary_public.py
python build_glossary_site.py
python build_quiz.py
python build_exam.py
```

前提：Python 3.x + openpyxl（`pip install openpyxl`）

## GitHub Pagesへの公開手順（概要）

1. このフォルダの中身（`index.html`・`glossary.html`・`quiz.html`・`exam.html`など）をリポジトリのルートまたは`docs/`フォルダに配置する
2. GitHubリポジトリの Settings → Pages で公開元ブランチ・フォルダを設定する
3. 公開後のURLでトップページ（`index.html`）が表示されることを確認する

`public_data.json`・`questions.json`・各`build_*.py`はビルド用のソースであり、公開ページの動作there自体には不要（同梱してもリスクはないが、リポジトリを軽くしたい場合は`.gitignore`等で除外してもよい）。

## 用語を追加・修正する場合の注意

新しい用語やdescを追加する際は、市販の参考書等の文章をそのまま使わず、独自の言い回しで説明文を書くこと。用語の技術的な正確性はWeb検索や一般的な技術知識で確認し、根拠のない年代・数値・固有名詞を創作しないこと。
