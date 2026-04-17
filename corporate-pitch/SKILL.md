---
name: corporate-pitch
description: >
  上��企業向け提案資料（PPTX + Figma）を生成するスキル。
  ALWAYS use this skill whenever the user mentions any of the following:
  提案資料作成, プレゼン資料, 企業向けプレゼン, コーポレートピッチ, ピッチデック作成,
  提案書を作りたい, 上場企業向け資料, 営業資料作成, corporate pitch, corporate proposal,
  formal presentation, 法人提案, ビジネスプレゼン, 取引先向け資料, client proposal deck,
  提案プレゼン, 取引先プレゼン, 営業プレゼン, 法人向けスライド, ピッチ資料.
  Accepts source files (PDF, DOCX, PPTX, images, TXT) as base material.
  Generates AIVALIX-branded slides with client company colors, AI-generated images
  (Nano Banana Pro / Gemini 3 Pro), Figma design file, and PPTX output.
  Builds one slide at a time for crash resilience.
  Pauses for user review at structure and mid-generation stages.
---

# Corporate Pitch — 上場企業向け提案資料生成スキル

## Overview
AIVALIX株式会社が上場企業に対してフォーマルな提案を行うための、洗練されたプレゼンテーション資料を生成する。

**成果物**: Figmaデザインファイル + PowerPoint (.pptx)  
**ブランド**: AIVALIX（黒 #000000 / 白 #FFFFFF）+ クライアント企業のテーマカラー  
**画像生成**: Nano Banana Pro (gemini-3-pro-image-preview)

### スクリプトパス
全スクリプトは `corporate-pitch/scripts/` に配置:
- `install_deps.py` — 依存パッケージ確認・インストール
- `extract_source_material.py` — ソース資料テキスト抽出
- `color_extractor.py` — ロゴからドミナントカラー抽出
- `generate_image_pro.py` — Nano Banana Pro画像生成
- `generate_pptx_corporate.py` — PPTX生成エンジン（1枚ずつ）

### リファレンス
- [slide_templates.md](references/slide_templates.md) — スライドタイプ別レイアウト仕様
- [japanese_corporate_guide.md](references/japanese_corporate_guide.md) — 上場企業向け提案書の作法

---

## ワークフロー

**重要**: 各ステップを順番に実行すること。ユーザー確認ゲートを飛ばさないこと。

---

### STEP 0 — 依存チェック

```bash
python "corporate-pitch/scripts/install_deps.py"
```

失敗した場合はユーザーに手動インストール手順を案内して続行。

---

### STEP 1 — 入力収集

ユーザーから以下の情報を収集する。不足情報はAskUserQuestionで確認。

| 項目 | 必須 | デフォルト |
|------|------|----------|
| ソース資料ファイル (PDF/DOCX/PPTX/TXT/画像) | Yes | - |
| クライアント企業名 (日本語・英語) | Yes | - |
| 提案目的・テーマ | Yes | - |
| クライアント企業ロゴ (PNG/JPG) | Recommended | なし |
| クライアントブランドカラー (hex) | Optional | ロゴから自動抽出 |
| 対象者 (部門・役職レベル) | Optional | 経営層 |
| AIVALIXロゴ | Optional | `G:\マイドライブ\AIVALIX関連\資料作成用素材\AIVALIXロゴ白 (1).PNG` |

**クライアントカラー自動抽出** (ロゴ提供時):
```bash
python "corporate-pitch/scripts/color_extractor.py" --logo "[クライアントロゴパス]"
```

---

### STEP 2 — コンテンツ合成

ソース資料からテキスト抽出:
```bash
python "corporate-pitch/scripts/extract_source_material.py" \
  --files "[file1]" "[file2]" ... \
  --output "extracted_content.json"
```

抽出結果を読み、以下のフォーマットで**構造化サマリー**をユーザーに提示:

```
## 収集情報サマリー

### 基本情報
- 提案元: AIVALIX株式会社
- 提案先: [クライアント名]
- 提案目的: [...]
- カラースキーム: AIVALIX黒白 + [クライアントアクセント色]

### セクション別詳細
1. [セクション名]: [要約]
2. ...

### 不足情報
- [不足している情報があれば列挙]
```

**🔒 ユーザー確認必須**: 「上記の情報で問題ありませんか？追加・修正があればお知らせください。」

---

### STEP 3 — 構成提案

スライド構成をテーブルで提案。[japanese_corporate_guide.md](references/japanese_corporate_guide.md)の標準構成を参考に。

```
## スライド構成案

| # | タイプ | タイトル | 画像 | 内容概要 |
|---|--------|----------|------|----------|
| 1 | cover | [提案タイトル] | background | ロゴ2つ + 日付 |
| 2 | agenda | 本日のご提案内容 | none | セクション一覧 |
| 3 | section | 課題認識 | background | セクション区切り |
| 4 | content | 貴社の現状と課題 | right | バレット5項目 |
| ... | | | | |
| N | back_cover | Thank You | none | 連絡先 + ロゴ |

**カラースキーム**:
- AIVALIX: #000000 / #FFFFFF
- クライアント: #[抽出色]
- 総スライド数: N枚

**画像生成計画**: X枚の画像をNano Banana Proで生成予定
```

**🔒 ユー���ー確認必須**: 「上記の構成でよろしいですか？スライドの追加・削除・順序変更があればお知らせください。」

---

### STEP 4 — PPTXスケルトン生成（画像スロット確保）

**新ワークフロー**: まずPPTXを画像なしで生成し、画像を入れるべき場所の正確なピクセルサイズを確定する。

```bash
PYTHONIOENCODING=utf-8 python "corporate-pitch/scripts/generate_pptx_corporate.py" \
  --structure "slide_structure.json" \
  --output "[ClientName]_Proposal_AIVALIX_skeleton.pptx" \
  --aivalix-logo "aivalix_logo_white_16x9.png" \
  --client-logo "[クライアントロゴパス]" \
  --client-color "#[アクセント色]" \
  --emit-image-slots
```

これにより `image_slots.json` が出力される:
```json
[
  {
    "slide_idx": 0,
    "w_px": 1280, "h_px": 720,
    "style_hint": "corporate",
    "prompt_hint": ""
  },
  {
    "slide_idx": 2,
    "w_px": 365, "h_px": 720,
    "style_hint": "diagram",
    "prompt_hint": ""
  }
]
```

---

### STEP 4.5 — 画像生成（正確なサイズで）

`image_slots.json` の各スロットに対して、**正確なピクセルサイズ**で画像を**1枚ずつ**生成。

```bash
python "corporate-pitch/scripts/generate_image_pro.py" \
  --prompt "[日本語テキストを含むビジネス図表の詳細プロンプト]" \
  --output "images/slide_N.jpg" \
  --width [w_px from slot] \
  --height [h_px from slot] \
  --style diagram
```

**重要な原則**:
- **トリミング不要**: スロットの正確なピクセルサイズを `--width` / `--height` で指定するため、生成後のクロップ処理が不要
- **日本語テキスト**: 画像内のすべてのテキスト・ラベル・注釈は**日本語**で出力される（デフォルト有効、`--no-japanese` で無効化可能）
- **フォント統一**: スライドと同じ **Noto Sans JP** フォントをプロンプトで指定
- **プロンプトにこだわる**: 単なるキーワードではなく、配置・構成・色・テキスト内容まで具体的に記述する

**プロンプト設計指針**:
- 表紙/裏表紙: 都市夜景、モダンオフィス → `--style corporate`
- 提案内容(技術説明): デジタルインターフェース → `--style technology`
- 課題認識: 産業設備、インフラ → `--style infrastructure`
- セクション区切り: 幾何学模様 → `--style abstract`
- **ビジネス概念図・構造説明**: コンサル風インフォグラフィック → `--style diagram` ← **優先使用**
- **人物の顔を避ける**。抽象的・環境的な画像を優先
- コンテンツスライドで概念や構造を説明する場合は、抽象画像より `diagram` スタイルを優先

**プロンプト品質チェックリスト** (diagram スタイル):
1. 図表の種類を明記（フローチャート、放射状図、マトリクス、ピラミッド等）
2. 各ボックス/ノードの日本語ラベルを具体的に記述
3. 矢印・接続線の方向と意味を説明
4. 色指示（黒・白・赤アクセント、クライアントカラー）
5. 全体のレイアウト構成（左→右フロー、中心→放射、上→下階層等）
6. アイコンの種類（歯車、グラフ、地図、データベース等）

**失敗時**: グラデーションプレースホルダーを使用して続行（スクリプトが自動対応）。

生成した画像パスを `image_map.json` に保存:
```json
{
  "0": {"path": "images/slide_cover.jpg"},
  "2": {"path": "images/slide_market_diagram.jpg"},
  ...
}
```

---

### STEP 5 — Figmaデザイン構築（メイン成果物）

Figma MCPツールを使用して洗練されたデザインを構築する。

#### 5.1 Figmaファイル作成
`create_new_file` で新規Figmaファイルを作成:
- ファイル名: `[ClientName]_Proposal_AIVALIX`

#### 5.2 デザインシステム設定
`use_figma` で以下を設定:
- カラー変数: AIVALIX黒(#000000), 白(#FFFFFF), クライアントアクセント
- テキストスタイル: タイトル(48pt), サブタイトル(18pt), 本文(14pt), キャプション(11pt)
- フォント: Noto Sans JP / Noto Sans

#### 5.3 スライドデザイン構築
各スライドを**1枚ずつ** 1920x1080 フレームとして構築:

`use_figma` で各フレームに:
1. 背景色/画像を設定
2. テキスト要素を配置（タイトル、本文、日付等）
3. ロゴを配置（AIVALIX + クライアント）
4. 図形・アクセント要素を配置（区切り線、バレットドット等）
5. 生成画像を配置

**1フレームごとに `get_screenshot` でビジュアル確認。**

#### 5.4 中間レビュー
全スライドの約50%が完成したら:

```
## 中間レビュー

スライド1〜[N]まで構築しました。Figmaファイルでご確認ください。

### 確認ポイント:
- デザインのトーンは適切ですか？
- テキスト量は多すぎ/少なすぎませんか？
- 画像の方向性は合っていますか？
- 色味の調整が必要ですか？

修正指示があればお知らせください。残りのスライドに反映します。
```

**🔒 ユーザーフィードバック収集**: フィードバックを受けて残りのスライドに反映。

#### 5.5 残りスライド完成
フィードバックを反映して残りのスライドを構築。

---

### STEP 6 — PPTX最終生成（画像挿入）

画像生成が完了したら、`image_map.json` を使って最終PPTXを生成する。

```bash
PYTHONIOENCODING=utf-8 python "corporate-pitch/scripts/generate_pptx_corporate.py" \
  --structure "slide_structure.json" \
  --output "[ClientName]_Proposal_AIVALIX.pptx" \
  --aivalix-logo "aivalix_logo_white_16x9.png" \
  --client-logo "[クライアントロゴパス]" \
  --client-color "#[アクセント色]"
```

**注意事項**:
- `PYTHONIOENCODING=utf-8` を設定して日本語テキストのエンコードエラーを防ぐ
- `--aivalix-logo` には黒背景用の白色版ロゴ（透過PNG、16:9）を指定
- `image_map.json` は `slide_structure.json` と同じディレクトリに配置すれば自動検出される
- 画像は `image_slots.json` で確定した正確なサイズで生成済みのため、**トリミング不要**
- `accent_light` はクライアントカラーから自動算出される（手動指定不要）

**1枚ずつ生成 + チェックポイント保存** — 途中でクラッシュしても最後のチェックポイントから回復可能。

---

### STEP 7 — 最終レビュー

成果物をユーザーに提示:

```
## 成果物

### 1. Figmaデザインファイル
📎 [Figma URL]
- 全スライドのデザインが閲覧・編集可能

### 2. PowerPointファイル  
📎 [ファイル名].pptx
- そのまま使用可能なPPTX形式

### 修正対応
改善したい箇所があれば具体的にお知らせください:
- テキスト修正 → 該当スライドのみ再生成
- 画像差し替え → 新プロンプトで再生成  
- レイアウト変更 → Figma + PPTXの該当スライドを修正
- 色調整 → カラーパラメータ変更して再生成
```

---

## エラーハンドリング

| 障害ポイント | 対処 |
|------------|------|
| パッケージ不足 | install_deps.pyで自動インストール。失敗時は手動手順を案内 |
| ファイル読み取り不可 | スキップ+警告、残りのソースで続行 |
| GEMINI_API_KEY未設定 | 警告表示。グラデーションプレースホルダーでPPTX生成は続行 |
| 画像生成失敗 | 3回リトライ後、グラデーションプレースホルダー使用 |
| スライド生成クラッシュ | 例外キャッチ、エラーPH挿入、パイプライン続行 |
| PPTX保存失敗 | 最終チェックポイントから回復 |
| Figma MCP利用不可 | PPTXのみ生成。ユーザーに通知 |
| ロゴファイル未発見 | ロゴなしで続行、警告表示 |
| 色抽出失敗 | デフォルトカラー #1A365D を使用 |

---

## スライド構造JSONの仕様

### image_prompt / image_style フィールド
各スライドに `image_prompt` と `image_style` を設定することで、画像生成のプロンプトとスタイルをスライド構造JSONから制御できる。

```json
{
  "type": "content",
  "title": "水道DX市場",
  "image_placement": "right",
  "image_style": "diagram",
  "image_prompt": "水道DX市場の構造を示す放射状図。中央に「水道DX市場 数千億円」、周囲に「AI管路診断」「GIS統合」「EBPM」「衛星データ」「IoTセンサー」の5つのノード。各ノードから中央へ矢印。黒・白・赤のカラースキーム。",
  "content": [...]
}
```

- `image_style`: `generate_image_pro.py` の `--style` に対応（corporate/technology/infrastructure/abstract/diagram）
- `image_prompt`: 画像生成の詳細プロンプト。日本語ラベル・図表構成を具体的に記述
- これらは `image_slots.json` の `style_hint` / `prompt_hint` にも出力される

### key_message フィールド
コンテンツスライド（cover, executive summary, back_cover 以外）には `key_message` フィールドを設定する。

```json
{
  "type": "content",
  "title": "スライドタイトル",
  "key_message": "このスライドで最も伝えたい示唆を1〜2行で記述",
  "bullets": [...]
}
```

- **配置**: ヘッダー（黒帯）の**下**、白背景コンテンツエリアの最上部に配置
- **色**: `text_dark`（#1A1A1A / 黒）— ヘッダー外なので黒文字で視認性を確保
- **フォントサイズ**: **20pt / 太字**（本文13ptより大きく、タイトル26pt以下。auto_shrink で最小16ptまで縮小）
- **高さ**: 0.72" のテキストボックス + 薄いセパレータ線（0.015"）で本文と区切り
- **効果**: `_draw_content_header` が `content_y` を返し、後続コンテンツが自動で下にシフトする
- **対象外**: cover, executive summary (slide index 1), back_cover

### accent_light カラー自動生成
クライアントカラー（`--client-color`）から自動的にaccent_lightが算出される:
```
lr = min(255, r + int((255 - r) * 0.55))
lg = min(255, g + int((255 - g) * 0.55))
lb = min(255, b + int((255 - b) * 0.55))
```
暗い背景（黒ヘッダー）上でセクションラベルやキーメッセージに使用する。

### ヘッダーレイアウト仕様（v7）

#### コンパクトヘッダー（1.0"）
- **高さ**: `header_h = Inches(1.0)` — タイトル1行に最適化したコンパクト設計
- **セクションラベル**: 12pt、y=0.08"、高さ0.30"、accent_light色
- **タイトル**: 26pt太字白、y=0.36"、高さ0.52"
  - **必ず1行**: `word_wrap=False` + `auto_shrink=True`（min_size=16pt）で強制
  - タイトルが長い場合はフォントサイズを自動縮小して1行に収める
- **アクセントライン**: ヘッダー直下に0.05"の赤帯

#### AIVALIXロゴ
- **ヘッダー内配置**: 全コンテンツスライドのヘッダー右上に配置
- **サイズ**: ロゴPNGに白余白パディングが含まれるため、基本サイズの**2.5倍**で配置
  - `logo_h = 0.90"` (基本0.36" × 2.5)
  - `logo_w = 2.62"` (16:9比率で算出)
  - `logo_y = 0.05"` (ヘッダー内上寄せ)
- **タイトル幅の予約**: ロゴとタイトルテキストの重なりを防ぐため、タイトル・セクションラベルの幅から `logo_reserve = 3.0"` を差し引く
- **ロゴファイル**: 黒背景用に白色版（透過PNG）を使用。`aivalix_logo_white_16x9.png`
- **白色ロゴ生成**: 元ロゴの暗いピクセルを白に、明るいピクセルを透過に変換

#### 画像配置ルール（ヘッダー外）
- **画像は黒ヘッダー帯を含まない** — `content_start_y`（ヘッダー + キーメッセージの下端）以降のみ
- `_content_layout_below_header()` が `content_start_y` を起点に画像領域を計算
- **適応型画像幅**: コンテンツ項目数に応じて最適化
  - 3項目以下: `img_w = 5.0"` （余裕ある図表表示）
  - 4〜5項目: `img_w = 4.5"`
  - 6項目以上: `img_w = 4.0"` （テキスト側を優先）
- **画像高さ**: `content_h = SLIDE_H - content_start_y - footer_margin(0.55")` で残り全高を使用
- **表紙 (cover)**: 唯一の例外。`placement=background` で全画面画像

### テーブルの適応型サイジング
テーブルスライド (`make_content_table`) では、内容量に応じてフォントサイズと行高さを自動調整する:

| 行数 | 最大セル文字数 | フォントサイズ |
|------|---------------|---------------|
| ≤4行 | ≤30文字 | 14pt |
| ≤5行 | — | 13pt |
| ≤7行 | — | 12pt |
| >7行 | — | 11pt |

- **行高さ**: テキスト行数（推定）に基づく最小行高さ (0.42"〜0.72") とスライド内利用可能高さの按分から算出
- **最大行高さ**: 1.05" を上限として余白の肥大化を防止
- **ヘッダー行**: 固定 0.50"、アクセントカラー背景

### 画像生成 — 2パスワークフロー

**原則**: 画像はスライドレイアウト確定後に、**スロットの正確なピクセルサイズ**で生成する。トリミングは行わない。

#### ワークフロー
1. `generate_pptx_corporate.py --emit-image-slots` でPPTXスケルトンを生成
2. 出力された `image_slots.json` から各スロットの `w_px` × `h_px` を取得
3. `generate_image_pro.py --width W --height H` で正確なサイズの画像を生成
4. `image_map.json` に画像パスを記録
5. 再度 `generate_pptx_corporate.py`（`--emit-image-slots` なし）で最終PPTXを生成

#### 画像テキストの原則
- **すべてのテキストは日本語** (デフォルト有効、`--no-japanese` で無効化)
- **フォント指定**: スライドと同じ **Noto Sans JP** をプロンプトで指定
- diagram スタイルでは自動的に日本語＋Noto Sans JP指定が含まれる

#### 画像スタイル
| スタイル | 用途 | 指定 |
|---------|------|------|
| corporate | 表紙/裏表紙: 都市夜景、オフィス | `--style corporate` |
| technology | 技術説明: デジタルUI | `--style technology` |
| infrastructure | 課題認識: 産業設備 | `--style infrastructure` |
| abstract | セクション区切り: 幾何学模様 | `--style abstract` |
| **diagram** | **ビジネス概念図・構造説明** | `--style diagram` ← **優先** |

- **優先方針**: コンテンツスライドで概念や構造を説明する場合は `diagram` スタイルを優先
- **人物の顔を避ける**。抽象的・環境的な画像を優先

#### プロンプト品質基準 (diagram スタイル)
プロンプトは単なるキーワードではなく、以下を具体的に記述すること:
1. 図表の種類（フローチャート、放射状図、マトリクス、ピラミッド等）
2. 各ボックス/ノードの**日本語ラベル**を具体的に列挙
3. 矢印・接続線の方向と意味
4. 色指示（黒・白・赤アクセント）
5. 全体のレイアウト構成（左→右、中心→放射、上→下等）
6. アイコンの種類（歯車、グラフ、地図、データベース等）

---

## 重要な制約

1. **1枚ずつ処理**: 画像生成もスライド構築も必ず1枚ずつ実行。並行処理しない
2. **ユーザー確認ゲート**: STEP 2, 3, 5.4 の3箇所で必ずユーザー確認を取る
3. **敬語レベル**: 提案書のテキストはです/ます調。「貴社」「弊社」を使用
4. **カラールール**: AIVALIX黒白をベースに、クライアント色はアクセントのみ。accent_lightはヘッダーのセクションラベルに自動適用。**キーメッセージは黒色(text_dark)**
5. **フォント**: Noto Sans JP / Noto Sans を統一使用
6. **キーメッセージ**: コンテンツスライドには必ず key_message を設定する（cover/executive summary/back_cover を除く）。白背景エリアに**20pt太字黒**で配置（本文13ptより大きく、タイトル26pt以下）
7. **タイトル1行**: コンテンツスライドのタイトルは**必ず1行**。`word_wrap=False` + `auto_shrink`（min 16pt）で強制。2行に折り返さない
8. **ヘッダーコンパクト化**: ヘッダー高さは1.0"。タイトル1行化により従来より低くなり、コンテンツ領域を最大化
9. **画像はヘッダー外**: 画像は黒ヘッダー帯（+ キーメッセージ）の**下**にのみ配置。`content_start_y` 以降の領域で適応的にサイズ決定
10. **ロゴサイズ**: AIVALIXロゴはPNG白余白を考慮し基本サイズの2.5倍 (0.90"×2.62") で配置。タイトル幅は logo_reserve=3.0" を確保してロゴとの重なりを防止
11. **テーブル最適化**: テーブルは内容量に応じてフォントサイズ(11〜14pt)と行高さ(0.42"〜1.05")を適応的に調整。余白の肥大化を防止
12. **画像スタイル**: コンテンツスライドの概念図・構造説明には `--style diagram` を優先使用。抽象画像のみに頼らない
13. **画像2パス生成**: 必ずPPTXスケルトン→image_slots.json→正確なサイズで画像生成→最終PPTX の順序で実行。画像のトリミングは行わない
14. **画像テキスト日本語**: 画像内のテキストはすべて日本語。フォントはスライドと同じNoto Sans JPを指定
15. **プロンプト品質**: diagramスタイルでは図表種類・ラベル・矢印・色・レイアウトを具体的に記述。キーワード羅列ではなく構成を説明する
