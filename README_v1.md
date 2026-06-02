# 応募書類セルフチェックツール ご利用の手引き（はじめての方向け）

本ツールは、第2回 SPReAD 公募の応募書類を、提出前にご自身で確認するためのものです。本手引きの手順に沿って進めると、確認結果が Microsoft Excel ファイルとして出力されます。

- かかる時間の目安：初回は 15〜20 分ほど（Python の準備を含む）／2回目以降は数分です。
- 用意するもの：インターネットにつながったパソコン（Windows または Mac）と、確認する応募書類。

> **この手引きの読み方について**
> 本手引きは、はじめて使う方でも確実に最後までたどり着けるよう、操作を一つずつ細かく説明しています。
> 操作に慣れている方は、各ステップの詳しい説明は読み飛ばし、下の「上級者向けクイックスタート」のみで進めていただいて構いません。ご自身に合うところからお読みください。

### 上級者向けクイックスタート（操作に慣れている方は本項のみで完了できます）

1. Python 3.11 以上を導入（Windows はインストール時に **「Add python.exe to PATH」にチェック**）。
2. 必要なライブラリを導入：`pip install openpyxl pdfminer.six pypdf`（Mac は `pip3`）。
3. ツール（`.py`）と確認する書類を **同じフォルダ** に置き、そのフォルダで実行：
   - 様式1（Excel）：`python research_plan_self_check.py`
   - 様式0・2・3・4（PDF）：`python form_self_check.py`
   - ※ Mac は `python3` を使用します。2つのツールは、それぞれ別のフォルダで実行してください。
4. 同じフォルダに出力される結果 Excel（`判定サマリー`／`詳細チェック`）を確認し、「要確認」の項目を修正のうえ再実行します。
   - 確認先・出力先を変更する場合は、各スクリプト冒頭の `INPUT_PATH`／`OUTPUT_FILE`／`RECURSIVE` を編集します（詳細は末尾の「上級者向け」を参照）。

以下では、上記の各手順を、はじめてご利用になる方向けに一つずつ説明します。

---

## 使用するツールの選択

確認する書類によって、使用するツール（プログラムのファイル）が異なります。

| 確認する書類 | 書類のファイル形式 | 使用するツール（ファイル名） |
| --- | --- | --- |
| 様式1（研究計画調書） | Microsoft Excel | `research_plan_self_check.py` |
| 様式0・2・3・4（チェックリスト・各種同意確認書） | PDF | `form_self_check.py` |

- **様式1を確認する** → `research_plan_self_check.py` を使用します。
- **様式0・2・3・4を確認する** → `form_self_check.py` を使用します。
- **両方を確認する** → 2つのツールを、それぞれ別のフォルダで実行します（後述）。

> **英語版について**：`research_plan_self_check.py` は様式1の日本語版・英語版の両方に対応します（言語を自動で判別し、結果も同じ言語で出力します）。`form_self_check.py` に英語版はありません（様式0・2・3・4 の日本語のみ対応）。

---

## 全体の流れ（3ステップ）

1. **ステップ1：Python をインストールする**（初回のみ）
2. **ステップ2：必要な部品（ライブラリ）をインストールする**（初回のみ）
3. **ステップ3：ツールを実行して、書類を確認する**（毎回）

2回目以降はステップ3のみで完了します。以下、順に説明します。

---

## 【ステップ1】Python をインストールする（初回のみ）

このツールを動かすには「Python（パイソン）」という無料のソフトウェアが必要です。すでにインストール済みの場合は、このステップは不要です。

### Windows をお使いの場合

1. インターネットで次のページを開きます → <https://www.python.org/downloads/>
2. ページの上のほうにある **黄色い「Download Python 3.x.x」ボタン** をクリックします。
   （3.x.x はその時点の最新版の数字です。3.11 以上であれば問題ありません。）
3. ダウンロードされたファイル（`python-3.x.x.exe`）を**ダブルクリックして開きます**。
   ※ 多くの場合、画面の下や「ダウンロード」フォルダに保存されています。
4. **【重要】** インストール画面が表示されたら、最も下にある **「Add python.exe to PATH」**（または「Add Python to PATH」）のチェックボックスに **必ずチェックを入れてください。**
   このチェックを入れ忘れると、後の手順でコマンドが正しく動作しません。入れ忘れてインストールした場合は、いったんアンインストールし、チェックを入れたうえで再度インストールしてください。
5. **「Install Now」** をクリックします。インストールが完了するまで、しばらくお待ちください。
6. **「Setup was successful」** と表示されたら **「Close」** をクリックして終了です。

### Mac をお使いの場合

1. インターネットで次のページを開きます → <https://www.python.org/downloads/>
2. **黄色い「Download Python 3.x.x」ボタン** をクリックします（3.11 以上であれば問題ありません）。
3. ダウンロードされたファイル（`python-3.x.x.pkg`）を**ダブルクリックして開きます**。
4. インストーラーの案内に従い、「続ける」→「同意する」→「インストール」と進めます。途中で Mac のパスワードを求められたら入力します。
5. 「インストールが完了しました」と表示されたら終了です。

> **【Mac の方への注意】** このあとの説明で出てくるコマンドは、Mac では `python` を `python3` に、`pip` を `pip3` に読み替えてください（本手引きの Mac 用の例ではすでに `python3`／`pip3` と書いてあります）。

---

## 【ステップ2】必要な部品（ライブラリ）をインストールする（初回のみ）

ツールが動くために必要な「部品」をまとめて入れます。ここでは初めて、コマンド（命令文）を入力するためのアプリを使います。このアプリは、**Windows では「コマンドプロンプト」、Mac では「ターミナル」** という名前です（本手引きでは、以降「コマンドプロンプト／ターミナル」と表記します）。手順に沿って操作すれば問題ありません。

### まず「コマンドプロンプト／ターミナル」を開きます

**Windows の場合：**
1. キーボードの **Windows キー**（窓のマーク）を押します。
2. そのまま `cmd` と入力します。
3. 「コマンド プロンプト」が表示されたら、**Enter キー** を押して開きます。

**Mac の場合：**
1. キーボードの **command キー + スペースキー** を同時に押します。
2. `ターミナル`（または `Terminal`）と入力します。
3. 「ターミナル」が表示されたら **Enter キー** を押して開きます。

### 次に、下のコマンドを入力して Enter を押します

コマンドプロンプト／ターミナルに、次の1行を**そのまま入力（コピー＆ペースト推奨）**して、Enter キーを押します。

**Windows：**
```
pip install openpyxl pdfminer.six pypdf
```

**Mac：**
```
pip3 install openpyxl pdfminer.six pypdf
```

しばらく文字が表示されたあと、最後に **`Successfully installed ...`** と表示されれば完了です。これで2種類のツールに必要な部品がそろいます。次回以降、このステップは不要です。

> **コピー＆ペーストのコツ**：コマンドプロンプト／ターミナルに貼り付けるときは、Windows のコマンドプロンプトでは画面内で **右クリック** すると貼り付けられます。Mac のターミナルでは **command + V** で貼り付けられます。

---

## 【ステップ3】ツールを実行して、書類を確認する（毎回）

ここからが実際の確認作業です。**最も簡単で確実な方法**をご案内します。

### ① 確認用のフォルダを用意します

デスクトップなどに、確認する書類ごとにフォルダを作成します。**ツール（.py ファイル）と確認する書類を、同じフォルダに入れる**ことが重要です。

次のように **2つのフォルダ** を作成することを推奨します（混在すると正しく集計できない場合があるため、ツールごとに分けてください）。

```
様式1チェック          … research_plan_self_check.py と 研究計画調書（Excel）を入れる
様式0・2・3・4チェック … form_self_check.py と 様式0・2・3・4 の PDF を入れる
```

> 確認対象が片方のみの場合は、フォルダは1つで構いません。

### ② そのフォルダの場所で「コマンドプロンプト／ターミナル」を開きます

**Windows の場合（簡単な方法）：**
1. 上で作ったフォルダ（例：「様式1チェック」）を開きます。
2. ウィンドウ上部の **アドレスバー**（今いる場所が「PC > デスクトップ > 様式1チェック」のように表示されている細長い欄）を1回クリックします。文字が選択された状態になります。
3. そこに `cmd` と入力して **Enter キー** を押します。
   → そのフォルダの場所でコマンドプロンプトが開きます。

**Mac の場合：**
1. ターミナルを開きます（ステップ2と同じ開き方）。
2. `cd ` と入力します（**cd のあとに半角スペースを1つ**入れます）。
3. 続けて、上で作ったフォルダを **Finder からターミナルの画面へドラッグ＆ドロップ** します（フォルダの場所が自動で入力されます）。
4. **Enter キー** を押します。→ そのフォルダに移動できます。

### ③ 実行のコマンドを入力して Enter を押します

開いたコマンドプロンプト／ターミナルに、使用するツールに合わせて次を入力し、Enter キーを押します。

**様式1（研究計画調書・Excel）を確認する場合：**

| | コマンド |
| --- | --- |
| Windows | `python research_plan_self_check.py` |
| Mac | `python3 research_plan_self_check.py` |

**様式0・2・3・4（PDF）を確認する場合：**

| | コマンド |
| --- | --- |
| Windows | `python form_self_check.py` |
| Mac | `python3 form_self_check.py` |

### ④ 結果ファイルを確認します

処理が終わると、**同じフォルダの中に確認結果の Excel ファイルが作成されます。**

- 様式1の結果：`研究計画調書_セルフチェック結果_(日時).xlsx`
- 様式0・2・3・4の結果：`同意確認書_セルフチェック結果_(日時).xlsx`

ファイル名には実行した日時が自動的に付くため、再度実行しても前回の結果は消えず、別ファイルとして保存されます。

### ⑤ 両方を確認する場合は、もう一方のフォルダでも同じ操作を行います

「様式0・2・3・4チェック」フォルダで①〜④を繰り返します。

> **【重要】** `.py` ファイルを**ダブルクリックで実行しないでください。** ウィンドウが一瞬だけ開いてすぐに閉じ、正しく動作しません。必ず上記の手順のとおり、コマンドプロンプト／ターミナルから `python ...` で実行してください。

---

## 使い方の例（最初から最後までの流れ）

実際の流れを把握しやすいよう、具体例で説明します。ここでは、様式1（研究計画調書・Excel）を Windows で確認する場合を例とします。

### 例：様式1（研究計画調書）を確認する（Windows）

**準備：フォルダにファイルを2つ入れる**

デスクトップに「様式1チェック」というフォルダを作り、その中にツールと確認する書類を入れます。中身は次のようになります。

```
様式1チェック
├─ research_plan_self_check.py   ← 配布されたツール
└─ 研究計画調書.xlsx              ← 確認する書類
```

**手順：**

1. 「様式1チェック」フォルダを開きます。
2. ウィンドウ上部のアドレスバーをクリックし、`cmd` と入力して Enter キーを押します。
   → コマンドプロンプトが、このフォルダの場所で開きます。画面には次のように、今いる場所が表示されています。
   ```
   C:\Users\user\Desktop\様式1チェック>
   ```
3. その表示の右側に、次のコマンドを入力して Enter キーを押します。
   ```
   python research_plan_self_check.py
   ```
4. 数行のメッセージが表示され、処理が終わると、最後に結果ファイルが作成された旨が表示されます（表示される文言は一例であり、実際とは異なる場合があります）。
   ```
   研究計画調書.xlsx を確認しました
   結果を 研究計画調書_セルフチェック結果_20260601_1430.xlsx に出力しました
   ```
5. 「様式1チェック」フォルダに戻ると、結果の Excel ファイルが新たに作成されています。
   ```
   様式1チェック
   ├─ research_plan_self_check.py
   ├─ 研究計画調書.xlsx
   └─ 研究計画調書_セルフチェック結果_20260601_1430.xlsx   ← これが結果ファイル
   ```
   （`20260601_1430` の部分は実行した日時です。例では 2026年6月1日 14時30分。）
6. この結果ファイルを開いて内容を確認します（次の「結果の確認」を参照）。

> **Mac の場合**：手順2は「ターミナルを開く → `cd ` と入力 → 『様式1チェック』フォルダをドラッグ → Enter」、手順3のコマンドは `python3 research_plan_self_check.py` になります。それ以外は同じです。

### 例：様式0・2・3・4（PDF）を確認する（Windows）

別のフォルダ「様式0・2・3・4チェック」を作り、ツールと PDF を入れます。

```
様式0・2・3・4チェック
├─ form_self_check.py
├─ 様式0_チェックリスト.pdf
├─ 様式2_同意確認書.pdf
└─ …（確認する PDF をまとめて格納できます）
```

操作は様式1の例と同じで、コマンドのみが異なります。

```
python form_self_check.py
```

実行後、同じフォルダに `同意確認書_セルフチェック結果_(日時).xlsx` が作成されます。

> PDF はフォルダにまとめて格納すれば、一度の実行でまとめて確認できます。文字を選択できる PDF をご使用ください（スキャンした画像の PDF は「判断不可」と表示されます）。

---

## 結果の確認

出力された Excel ファイルには、**「判定サマリー」** と **「詳細チェック」** の2つのシートがあります。

1. まず「判定サマリー」シートで、各ファイルが **「OK」** か **「要確認」** かを確認します。表示例は次のとおりです。

   | ファイル | 判定 |
   | --- | --- |
   | 研究計画調書.xlsx | 要確認 |

2. 「要確認」と表示された場合は、「詳細チェック」シートで該当する項目を確認します（表示例）。

   | 項目 | 判定 | 内容 |
   | --- | --- | --- |
   | 研究目的の文字数 | OK | — |
   | 研究計画の文字数 | 要確認 | 文字数が不足しています |

3. 指摘された項目を、元の書類で修正します。
4. 修正後、再度ツールを実行して確認します。すべて「OK」となれば、形式面の確認は完了です。

> 本ツールは **形式面の確認を補助する** ものです。「OK」と表示されても、記載内容が公募要領の要件を満たしているかは別途ご確認ください。提出の最終判断は、応募者ご自身の責任で行ってください。

---

## 想定される事象と対処

| 症状・表示 | 対処 |
| --- | --- |
| `python is not recognized`／`python` が認識されない | Python が入っていないか、ステップ1の「Add python.exe to PATH」のチェックを入れ忘れています。ステップ1をやり直してください。コマンドプロンプト／ターミナルをいったん閉じて開き直す、またはパソコンを再起動すると直ることもあります。Mac の方は `python` ではなく `python3` をお使いください。 |
| `ModuleNotFoundError` と表示される | 必要な部品が入っていません。ステップ2のコマンド（`pip install ...`／Mac は `pip3 install ...`）をもう一度実行してください。 |
| ダブルクリックするとウィンドウが一瞬で閉じる／何も起きない | ダブルクリックでは動きません。ステップ3の手順どおり、コマンドプロンプト／ターミナルから `python ...` で実行してください。 |
| 様式1で、記入済みなのに「未記入」「文字数不足」と判定される | Excel でその書類をいったん上書き保存し直してから、もう一度実行してください。 |
| PDF が「判断不可」と表示される | 文字を選択できる PDF が必要です。スキャン（画像として取り込んだ）PDF は文字を読み取れません。文字を選択できる形で作り直してください。 |
| 結果ファイルが開けない／文字化けする | 最新版の Microsoft Excel で開いてください。 |

---

## 【上級者向け・通常は不要】より柔軟に利用する場合

ここから先は、PC 操作に慣れている方向けの補足です。**通常はステップ3の方法のみで十分**であり、読み飛ばして構いません。

### Jupyter Notebook から実行する

```
from ツール名 import run
run(r"確認するフォルダのパス")
```

- 様式1：`from research_plan_self_check import run` のあと `run(r"C:\Users\(ユーザー名)\Desktop\様式1チェック")`
- 様式0・2・3・4：`from form_self_check import run` のあと `run(r"C:\Users\(ユーザー名)\Desktop\様式0・2・3・4チェック")`

### 確認先・出力先を変更する場合（スクリプトの編集）

既定では「ツールを置いて実行したフォルダ内の書類」を確認し、結果も同じフォルダに出力します。別の場所を指定する場合は、各スクリプト冒頭の設定（`INPUT_PATH`／`OUTPUT_FILE`／`RECURSIVE`）を書き換えます。

- `INPUT_PATH`：確認対象。書類が入ったフォルダ、または個別ファイルのパスを指定します。既定値の `Path(".")` は「ツールを実行したフォルダ」を指します。
  - 例：`INPUT_PATH = Path(r"C:\Users\(ユーザー名)\Desktop\様式1チェック")`
- `OUTPUT_FILE`：結果ファイルの保存先（フォルダ）。ファイル名には実行のたびに日時が自動で付くため、前回の結果は上書きされません。
  - 例：`OUTPUT_FILE = Path(r"C:\Users\(ユーザー名)\Desktop\結果\研究計画調書_セルフチェック結果.xlsx")`
- `RECURSIVE`：`True` にすると、指定フォルダ配下のサブフォルダもまとめて確認します（機関で複数の応募書類を一括で確認する場合にご利用ください）。

> Windows のパスは `r"..."`（先頭に `r` を付けた形）で書くと、`\`（円記号／バックスラッシュ）をそのまま指定できます。

---

## 機関向け：複数の応募書類を一括で確認する

機関（応募者所属機関）で多数の応募書類をまとめて確認する場合は、応募者ごとにサブフォルダを作成し、その親フォルダを一度の実行でまとめて確認できます。全員分の確認結果は、1つの Microsoft Excel ファイルに出力されます。

事前に、ステップ1（Python のインストール）およびステップ2（ライブラリのインストール）を済ませてください（初回のみ）。手順は次のとおりです。

1. 親フォルダ（例：「様式1一括チェック」）を作成し、その中に応募者ごとのサブフォルダを作って、それぞれに対象の書類を格納します。ツール（`.py`）は親フォルダに置きます。
2. スクリプト冒頭の `RECURSIVE` を `True` に変更します。
3. 親フォルダでコマンドプロンプト／ターミナルを開き、通常どおり実行します。

```
様式1一括チェック
├─ research_plan_self_check.py   （RECURSIVE = True に変更）
├─ 申請者A
│   └─ 研究計画調書.xlsx
├─ 申請者B
│   └─ 研究計画調書.xlsx
└─ 申請者C
    └─ 研究計画調書.xlsx
```

| 確認する書類 | コマンド（Windows） | コマンド（Mac） |
| --- | --- | --- |
| 様式1（Excel） | `python research_plan_self_check.py` | `python3 research_plan_self_check.py` |
| 様式0・2・3・4（PDF） | `python form_self_check.py` | `python3 form_self_check.py` |

親フォルダ配下のすべてのサブフォルダがまとめて確認され、全員分の結果が1つの Excel ファイル（「判定サマリー」シートに一覧）として出力されます。

- サブフォルダ名（例：申請者名や受付番号）は、結果ファイルで各書類を識別する手がかりになります。識別しやすい名称を推奨します。
- 結果を別の場所にまとめて保存する場合は、`OUTPUT_FILE` に保存先フォルダを指定できます（前掲「確認先・出力先を変更する場合」を参照）。
- 様式1と様式0・2・3・4 は、必ず別の親フォルダで実行してください（結果が混在し、正しく集計できない場合があります）。

---

## ご利用条件

本ツールは、SPReAD 第2回公募の応募者向けに配布するセルフチェック用ツールです。応募者本人、応募者所属機関による応募書類の確認を目的とした利用に限ります。

- 本ツールの商用利用は禁止します。
- ご自身の利用の範囲内での改変は可能ですが、改変したものを第三者へ配布することは禁止します。
- 本ツールは、必ず公式配布元から入手してください。第三者により再配布されたものの利用は認められません。

本ツールの使用により生じた損害について、文部科学省は責任を負いません。

---

## サードパーティ・ライブラリのライセンス

本ツールは、下記のオープンソースライブラリを利用します。これらのライブラリは、応募者ご自身が pip コマンドにより PyPI（Python Package Index）から取得するものであり、本配布物には同梱しておりません。各ライブラリの著作権およびライセンスは、それぞれの提供元（著作権者）に帰属します。ライセンスの全文は、下表の配布元ページに掲載されている LICENSE ファイル等でご確認いただけます。

**直接利用するライブラリ**（`pip install openpyxl pdfminer.six pypdf` で導入されるもの）

| ライブラリ | ライセンス | 配布元・ライセンス全文 |
| --- | --- | --- |
| openpyxl | MIT License | <https://foss.heptapod.net/openpyxl/openpyxl> |
| pypdf | BSD 3-Clause License | <https://github.com/py-pdf/pypdf> |
| pdfminer.six | MIT License | <https://github.com/pdfminer/pdfminer.six> |

**依存関係として自動的に導入されるライブラリ**

上記ライブラリは、動作に必要な依存ライブラリを pip により自動的に導入します。導入される依存ライブラリは実行環境（OS および Python のバージョン等）により異なる場合がありますが、代表的なものは下表のとおりであり、いずれも寛容（permissive）なオープンソースライセンスで配布されています。

| 依存ライブラリ | ライセンス | 配布元 |
| --- | --- | --- |
| et-xmlfile（openpyxl の依存） | MIT License | <https://pypi.org/project/et-xmlfile/> |
| charset-normalizer（pdfminer.six の依存） | MIT License | <https://pypi.org/project/charset-normalizer/> |
| cryptography（pdfminer.six の依存） | Apache License 2.0 または BSD 3-Clause License | <https://pypi.org/project/cryptography/> |
| cffi（cryptography の依存） | MIT License | <https://pypi.org/project/cffi/> |
| pycparser（cffi の依存） | BSD 3-Clause License | <https://pypi.org/project/pycparser/> |

> 上記の直接利用ライブラリおよびその依存ライブラリは、いずれも寛容（permissive）なオープンソースライセンス（MIT・BSD 3-Clause・Apache License 2.0）で提供されており、コピーレフト型ライセンス（GNU GPL／LGPL／AGPL 等）は含まれていません。

なお、本ツールの実行には Python 本体（Python Software Foundation License）が必要です。Python についても、応募者ご自身が python.org 等から取得するものであり、本配布物には同梱しておりません。

---

# Application Document Self-Check Tools — User Guide (for first-time users)

These tools help you check your application documents for the 2nd SPReAD call before submission.
If you follow this guide step by step, the result will be saved as a Microsoft Excel file.

- Time needed: about 15–20 minutes the first time (including setting up Python); just a few minutes after that.
- What you need: a computer (Windows or Mac) connected to the internet, and the documents you want to check.

> **How to read this guide**
> This guide explains every action step by step so that first-time users can reach the end with confidence.
> If you are already comfortable with these operations, you may skip the detailed explanations and proceed using only the "Quick start (for experienced users)" below. Please read from whichever section suits you.

### Quick start (for experienced users)

1. Install Python 3.11+ (on Windows, **check "Add python.exe to PATH"** during installation).
2. Install the libraries: `pip install openpyxl pdfminer.six pypdf` (use `pip3` on Mac).
3. Put the tool (`.py`) and the documents in the **same folder**, then run it from that folder:
   - Form 1 (Excel): `python research_plan_self_check.py`
   - Forms 0, 2, 3, 4 (PDF): `python form_self_check.py`
   - On Mac, use `python3`. Run the two tools in separate folders.
4. Open the result Excel (`判定サマリー` / `詳細チェック`) created in the same folder. Fix any "要確認" items and run again.
   - To change the input/output locations, edit `INPUT_PATH` / `OUTPUT_FILE` / `RECURSIVE` at the top of each script (see the "advanced users" section at the end).

The sections below explain each of these steps one at a time for first-time users.

---

## Selecting the tool to use

The tool you use depends on the document you want to check.

| Document to check | File format | Tool (file name) |
| --- | --- | --- |
| Form 1 (Research Plan) | Microsoft Excel | `research_plan_self_check.py` |
| Forms 0, 2, 3, and 4 (checklist and consent forms) | PDF | `form_self_check.py` |

- **To check Form 1** → use `research_plan_self_check.py`.
- **To check Forms 0, 2, 3, and 4** → use `form_self_check.py`.
- **To check both** → run the two tools in separate folders (explained below).

> **About English support:** `research_plan_self_check.py` supports both the Japanese and English versions of Form 1 (it detects the language automatically and outputs the results in the same language). `form_self_check.py` has no English version; it supports the Japanese Forms 0, 2, 3, and 4 only.

---

## Overview (3 steps)

1. **Step 1: Install Python** (first time only)
2. **Step 2: Install the required components (libraries)** (first time only)
3. **Step 3: Run a tool and check your documents** (every time)

From the second time onward, only Step 3 is required. The following sections explain each step in order.

---

## [Step 1] Install Python (first time only)

To run these tools, you need "Python," a free software application. If it is already installed, this step is not required.

### On Windows

1. Open this page in your browser → <https://www.python.org/downloads/>
2. Click the **yellow "Download Python 3.x.x" button** near the top (3.x.x is the latest version; 3.11 or later is fine).
3. **Double-click** the downloaded file (`python-3.x.x.exe`) to open it. It is usually in your "Downloads" folder.
4. **[Important]** On the installer screen, be sure to check the **"Add python.exe to PATH"** (or "Add Python to PATH") checkbox at the bottom. If you skip this, the commands will not work in later steps. If you forgot to check it, uninstall Python and reinstall it with the box checked.
5. Click **"Install Now"** and wait for it to finish.
6. When **"Setup was successful"** appears, click **"Close."**

### On Mac

1. Open this page in your browser → <https://www.python.org/downloads/>
2. Click the **yellow "Download Python 3.x.x" button** (3.11 or later is fine).
3. **Double-click** the downloaded file (`python-3.x.x.pkg`) to open it.
4. Follow the installer: "Continue" → "Agree" → "Install." Enter your Mac password if asked.
5. When it says the installation is complete, you are done.

> **[Note for Mac users]** In the steps below, replace `python` with `python3` and `pip` with `pip3` (the Mac examples in this guide already use `python3` / `pip3`).

---

## [Step 2] Install the required components (libraries) (first time only)

This installs the components the tools need to run. Here you will use, for the first time, the application for typing commands. It is called **"Command Prompt" on Windows** and **"Terminal" on Mac** (this guide refers to it as "Command Prompt / Terminal" from here on). Following the steps below will work without difficulty.

### First, open Command Prompt / Terminal

**On Windows:**
1. Press the **Windows key** (the window icon).
2. Type `cmd`.
3. When "Command Prompt" appears, press **Enter** to open it.

**On Mac:**
1. Press **command + space** together.
2. Type `Terminal`.
3. When "Terminal" appears, press **Enter** to open it.

### Then type the command below and press Enter

Type (copy & paste is recommended) the following single line and press Enter.

**Windows:**
```
pip install openpyxl pdfminer.six pypdf
```

**Mac:**
```
pip3 install openpyxl pdfminer.six pypdf
```

After some text is displayed, the process is complete when **`Successfully installed ...`** appears at the end. This installs the components required by both tools. You do not need to repeat this step from the next time onward.

> **Paste tip:** In Windows Command Prompt, **right-click** in the window to paste. In Mac Terminal, use **command + V**.

---

## [Step 3] Run a tool and check your documents (every time)

This is the actual checking process. The **simplest and most reliable** method is described below.

### 1. Prepare a folder

On your Desktop (or anywhere), create a folder for the documents you want to check. It is important to **place the tool (the .py file) and the documents in the same folder.**

We recommend creating **two folders** (separate them by tool, since mixing them may prevent the results from being aggregated correctly):

```
Form1_check          … put research_plan_self_check.py and the Research Plan (Excel)
Forms0-2-3-4_check   … put form_self_check.py and the PDFs of Forms 0, 2, 3, and 4
```

> If you only need to check one of them, one folder is enough.

### 2. Open Command Prompt / Terminal in that folder

**On Windows (easy way):**
1. Open the folder you created (e.g., "Form1_check").
2. Click once on the **address bar** (the thin bar at the top showing your location, like "PC > Desktop > Form1_check"). The text becomes selected.
3. Type `cmd` there and press **Enter**. → Command Prompt opens already inside that folder.

**On Mac:**
1. Open Terminal (same as in Step 2).
2. Type `cd ` (with **one space** after `cd`).
3. **Drag and drop** the folder from Finder onto the Terminal window (the folder's location is filled in automatically).
4. Press **Enter**. → You are now "inside" that folder.

### 3. Type the run command and press Enter

In Command Prompt / Terminal, type the command for your tool and press Enter.

**To check Form 1 (Research Plan, Excel):**

| | Command |
| --- | --- |
| Windows | `python research_plan_self_check.py` |
| Mac | `python3 research_plan_self_check.py` |

**To check Forms 0, 2, 3, and 4 (PDF):**

| | Command |
| --- | --- |
| Windows | `python form_self_check.py` |
| Mac | `python3 form_self_check.py` |

### 4. Check the result file

When it finishes, **a result Excel file is created in the same folder.**

- Form 1 result: `研究計画調書_セルフチェック結果_(date-time).xlsx` (Japanese form) or `Research_Plan_Self_Check_Result_(date-time).xlsx` (English form)
- Forms 0, 2, 3, 4 result: `同意確認書_セルフチェック結果_(date-time).xlsx`

A date-time stamp is added to the file name automatically, so running it again does not overwrite previous results (they accumulate as separate files).

### 5. If you need both, repeat in the other folder

Repeat steps 1–4 in the "Forms0-2-3-4_check" folder.

> **[Important]** Do **not** run the `.py` file by double-clicking it. A window will flash open and close, and it will not work. Always run it with `python ...` from Command Prompt / Terminal as shown above.

---

## A worked example (start to finish)

To make the flow concrete, here is a full example of checking Form 1 (Research Plan, Excel) on Windows.

### Example: checking Form 1 (Research Plan) on Windows

**Prepare: put two files in a folder**

Create a folder called "Form1_check" on the Desktop, and put the tool and your document inside it:

```
Form1_check
├─ research_plan_self_check.py   ← the distributed tool
└─ 研究計画調書.xlsx              ← the document you want to check
```

**Steps:**

1. Open the "Form1_check" folder.
2. Click the address bar at the top, type `cmd`, and press Enter.
   → Command Prompt opens inside this folder, showing your current location:
   ```
   C:\Users\user\Desktop\Form1_check>
   ```
3. To the right of that prompt, type the following and press Enter:
   ```
   python research_plan_self_check.py
   ```
4. A few lines of messages are displayed, and when it finishes, it reports that a result file has been created (the wording shown is an example and may differ from the actual output):
   ```
   Checked 研究計画調書.xlsx
   Saved the result to 研究計画調書_セルフチェック結果_20260601_1430.xlsx
   ```
5. Back in the "Form1_check" folder, a new result Excel file has appeared:
   ```
   Form1_check
   ├─ research_plan_self_check.py
   ├─ 研究計画調書.xlsx
   └─ 研究計画調書_セルフチェック結果_20260601_1430.xlsx   ← the result file
   ```
   (The `20260601_1430` part is the date and time of the run — in this example, 14:30 on 1 June 2026.)
6. Open this result file to review it (see "Checking the results" below).

> **On Mac:** for step 2, "open Terminal → type `cd ` → drag the Form1_check folder → Enter"; and the command in step 3 becomes `python3 research_plan_self_check.py`. Everything else is the same.

### Example: checking Forms 0, 2, 3, and 4 (PDF) on Windows

Create a separate folder "Forms0-2-3-4_check" and put the tool and the PDFs inside it:

```
Forms0-2-3-4_check
├─ form_self_check.py
├─ 様式0_チェックリスト.pdf
├─ 様式2_同意確認書.pdf
└─ … (put all the PDFs you want to check here)
```

The rest is exactly like the Form 1 example; only the command changes:

```
python form_self_check.py
```

After it runs, a `同意確認書_セルフチェック結果_(date-time).xlsx` file is created in the same folder.

> You can put all the PDFs in the folder and check them in a single run. Use PDFs whose text can be selected (scanned image PDFs will show "判断不可").

---

## Checking the results

The output Excel file has two sheets: **"判定サマリー (Summary)"** and **"詳細チェック (Detail Check)."**

1. In the Summary sheet, check whether each file is **"OK"** or **"要確認 (needs review)."** An example display is shown below:

   | File | Result |
   | --- | --- |
   | 研究計画調書.xlsx | 要確認 (needs review) |

2. If "要確認" appears, open the "詳細チェック (Detail Check)" sheet to see which item was flagged (example):

   | Item | Result | Note |
   | --- | --- | --- |
   | Character count: research objective | OK | — |
   | Character count: research plan | 要確認 | Too few characters |

3. Fix the flagged items in your original document.
4. Run the tool again to re-check. When everything shows "OK," the formatting check is complete.

> These tools only assist with **formatting checks.** Even if "OK" is shown, please separately confirm whether the content meets the requirements of the call guidelines. The final decision to submit is the applicant's own responsibility.

---

## Troubleshooting

| Symptom | Action |
| --- | --- |
| `python is not recognized` | Python is not installed, or you forgot to check "Add python.exe to PATH" in Step 1. Redo Step 1. Closing and reopening Command Prompt / Terminal, or restarting the PC, can also help. On Mac, use `python3` instead of `python`. |
| `ModuleNotFoundError` appears | The required components are missing. Run the Step 2 command again (`pip install ...`, or `pip3 install ...` on Mac). |
| Double-clicking flashes a window / does nothing | Double-clicking does not work. Run it with `python ...` from Command Prompt / Terminal as in Step 3. |
| Form 1 shows "未記入 (not filled in)" or "文字数不足 (too few characters)" even though it is filled in | Re-save the file in Excel, then run again. |
| A PDF shows "判断不可 (cannot determine)" | A text-selectable PDF is required. Scanned (image) PDFs cannot be read. Recreate it as a PDF whose text can be selected. |
| The result file will not open / characters are garbled | Open it in the latest version of Microsoft Excel. |

---

## [For advanced users — usually not needed] More flexible usage

The following is supplementary information for users comfortable with PC operations. **The Step 3 method above is normally sufficient,** so you may skip this section.

### Run from Jupyter Notebook

```
from <tool name> import run
run(r"path to the folder to check")
```

- Form 1: `from research_plan_self_check import run` then `run(r"C:\Users\(username)\Desktop\Form1_check")`
- Forms 0, 2, 3, 4: `from form_self_check import run` then `run(r"C:\Users\(username)\Desktop\Forms0-2-3-4_check")`

### Changing the input/output locations (editing the script)

By default, each tool checks the documents in the folder where it is placed and run, and saves the result in that same folder. To use different locations, edit the settings at the top of each script (`INPUT_PATH`, `OUTPUT_FILE`, `RECURSIVE`).

- `INPUT_PATH`: what to check. A folder containing the documents, or the path to an individual file. The default `Path(".")` refers to the folder where the tool is run.
  - Example: `INPUT_PATH = Path(r"C:\Users\(username)\Desktop\Form1_check")`
- `OUTPUT_FILE`: where the result file is saved (its folder). A date-time stamp is added automatically on each run, so previous results are not overwritten.
  - Example: `OUTPUT_FILE = Path(r"C:\Users\(username)\Desktop\results\研究計画調書_セルフチェック結果.xlsx")`
- `RECURSIVE`: set to `True` to also check subfolders under the specified folder (useful when an institution checks multiple applications at once).

> On Windows, writing a path as `r"..."` (with a leading `r`) lets you use `\` (backslash) directly.

---

## For institutions: checking multiple applications at once

To check many applications by an institution (an applicant's affiliated institution) at once, create one subfolder per applicant and check the entire parent folder in a single run. The results for every applicant are written to a single Microsoft Excel file.

Complete Step 1 (install Python) and Step 2 (install the libraries) beforehand (first time only). The procedure is as follows.

1. Create a parent folder (e.g., "Form1_batch_check"), and inside it create one subfolder per applicant, each containing that applicant's document. Place the tool (`.py`) in the parent folder.
2. Change `RECURSIVE` to `True` at the top of the script.
3. Open Command Prompt / Terminal in the parent folder and run the tool as usual.

```
Form1_batch_check
├─ research_plan_self_check.py   (set RECURSIVE = True)
├─ ApplicantA
│   └─ 研究計画調書.xlsx
├─ ApplicantB
│   └─ 研究計画調書.xlsx
└─ ApplicantC
    └─ 研究計画調書.xlsx
```

| Document to check | Command (Windows) | Command (Mac) |
| --- | --- | --- |
| Form 1 (Excel) | `python research_plan_self_check.py` | `python3 research_plan_self_check.py` |
| Forms 0, 2, 3, 4 (PDF) | `python form_self_check.py` | `python3 form_self_check.py` |

All subfolders under the parent are checked together, and the results for every applicant are written to a single Excel file (listed in the "判定サマリー (Summary)" sheet).

- Subfolder names (e.g., applicant name or reference number) help identify each document in the result file. Use names that are easy to distinguish.
- To save the results to a single separate location, specify a destination folder in `OUTPUT_FILE` (see "Changing the input/output locations" above).
- Always run Form 1 and Forms 0/2/3/4 in separate parent folders (otherwise the results may be mixed together and not aggregated correctly).

---

## Terms of use

These are self-check tools distributed to applicants of the 2nd SPReAD call. Use is limited to checking application documents by the applicant themselves or by the applicant's affiliated institution.

- Commercial use of these tools is prohibited.
- You may modify the tools for your own use, but distributing modified versions to third parties is prohibited.
- Always obtain these tools from the official distribution source. Use of copies redistributed by third parties is not permitted.

MEXT (the Ministry of Education, Culture, Sports, Science and Technology) is not liable for any damages arising from the use of these tools.

---

## Third-party library licenses

These tools use the open-source libraries listed below. Each applicant obtains these libraries via the pip command from PyPI (the Python Package Index); they are **not** bundled with this distribution. The copyright and license of each library belong to its respective project (copyright holder). The full text of each license can be found in the LICENSE file and on the project pages linked in the tables below.

**Libraries used directly** (installed by `pip install openpyxl pdfminer.six pypdf`)

| Library | License | Project / full license text |
| --- | --- | --- |
| openpyxl | MIT License | <https://foss.heptapod.net/openpyxl/openpyxl> |
| pypdf | BSD 3-Clause License | <https://github.com/py-pdf/pypdf> |
| pdfminer.six | MIT License | <https://github.com/pdfminer/pdfminer.six> |

**Dependencies installed automatically**

The libraries above automatically install the dependencies they require via pip. The exact set of dependencies may vary depending on the environment (OS and Python version), but the representative ones are listed below, and all are distributed under permissive open-source licenses.

| Dependency | License | Source |
| --- | --- | --- |
| et-xmlfile (dependency of openpyxl) | MIT License | <https://pypi.org/project/et-xmlfile/> |
| charset-normalizer (dependency of pdfminer.six) | MIT License | <https://pypi.org/project/charset-normalizer/> |
| cryptography (dependency of pdfminer.six) | Apache License 2.0 or BSD 3-Clause License | <https://pypi.org/project/cryptography/> |
| cffi (dependency of cryptography) | MIT License | <https://pypi.org/project/cffi/> |
| pycparser (dependency of cffi) | BSD 3-Clause License | <https://pypi.org/project/pycparser/> |

> All of the directly used libraries and their dependencies above are provided under permissive open-source licenses (MIT, BSD 3-Clause, or Apache License 2.0); no copyleft licenses (such as the GNU GPL, LGPL, or AGPL) are used.

Running these tools also requires Python itself (Python Software Foundation License). Python, too, is obtained by each applicant from python.org or a similar source and is not bundled with this distribution.
