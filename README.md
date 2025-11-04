# form-creator

GoogleFormをjsonから生成するApps Script。

# usage

## 初期ファイル作成

1. Google スプレッドシートを新規で作成する。
2. 「拡張機能」→「Apps Script」の順に開く。
3. 「ファイル」の中に`コード.gs`と`filepicker.html`というファイルを作り、このリポジトリ内の`src`ディレクトリ内の同名ファイルの内容を書き込む。
4. 保存してApps Scriptは閉じる。
5. スプレッドシートを保存し、一度閉じる。
6. 作成したスプレッドシートと同階層に`form-templates`というディレクトリを作成する。
7. 作成したスプレッドシートを開く。
8. メニューに「JSONからGoogleForm作成」が増えていることを確認する。
9.  「JSON選択」を選択します。
10. スクリプトの実行に許可が必要なので、内容を確認し、許可する。

## 使用方法

1. `form-templates`ディレクトリに元となるJSONを配置する。
2. スプレッドシートを開く。
3. B1セルにJSONファイルの探査先ディレクトリを記載しておく必要があるため、B1セルに`form-templates`と入力する。
4. 「JSONからGoogleForm作成」→「JSON選択」を開く。
5. フォーム作成に使用したいJSONを選択し、「フォーム作成」をクリックする。
6. 処理が正常に実行されれば「作成完了」とメッセージが表示され、スプレッドシートと同階層にフォームファイルが作成され、B2セルにそのURLが書き込まれる。

***任意  デフォルトではB1セルに探査先ディレクトリ、B2セルに生成されたフォームURLが書き込まれるため、それが分かるようにシートを整えておく。***


# JSON仕様

## 全体構造

```json
{
  "title": "アンケートタイトル",
  "description": "アンケートの説明",
  "choicesTemplates": [ ... ],
  "questions": [ ... ]
}
```

* **title**: アンケート全体のタイトル
* **description**: アンケート全体の説明文
* **choicesTemplates**: 繰り返し利用する選択肢群のテンプレート定義
* **questions**: 実際の質問リスト

## choicesTemplates

共通の選択肢を定義し、複数の質問で参照可能。

```json
{
  "name": "理解度チェック",
  "choices": [
    "1:分からない",
    "2:知っているが説明できない",
    "3:説明できるが実現方法が分からない/実装できない",
    "4:説明/実現できる",
    "5:説明/実現/適用判断ができる"
  ]
}
```

* **name**: テンプレート名
* **choices**: 選択肢の配列

## questions

質問ごとにオブジェクトを定義する。
`type` に応じて利用可能なフィールドが異なる。

### 共通フィールド

* **type**: 質問タイプ
  * pagebreak
  * multiplechoice
  * checkbox
  * list
  * text
  * paragraph
  * scale
  * date
  * time
  * grid
  * checkboxgrid
* **title**: 質問文
* **required**: 必須入力かどうか（true/false）
* **helpText**: 補助説明（任意）

### タイプ別仕様

#### 1. pagebreak

ページ区切りやセクションタイトルとして利用。

```json
{
  "type": "pagebreak",
  "title": "前半の質問",
  "helpText": "前半の質問です"
}
```

#### 2. multiplechoice

単一選択式。

```json
{
  "type": "multiplechoice",
  "title": "SpringBootの理解度",
  "choices": [
    "1:分からない",
    "2:知っているが説明できない",
    "3:説明できるが実現方法が分からない/実装できない",
    "4:説明/実現できる",
    "5:説明/実現/適用判断ができる"
  ]
}
```

choicesTemplatesを作っていれば、それを参照することもできます。
```json
{
  "type": "multiplechoice",
  "title": "SpringBootの理解度",
  "choices": "理解度チェック"
}
```



#### 3. checkbox

複数選択式。

```json
{
  "type": "checkbox",
  "title": "得意な言語",
  "choices": ["Java", "Python", "Go"]
}
```

#### 4. list

プルダウン形式の単一選択。

```json
{
  "type": "list",
  "title": "興味のある分野",
  "choices": ["Web", "AI", "IoT"]
}
```

#### 5. text

一行テキスト入力。

```json
{
  "type": "text",
  "title": "担当プロジェクト名"
}
```

#### 6. paragraph

複数行テキスト入力。

```json
{
  "type": "paragraph",
  "title": "自己PR"
}
```

#### 7. scale

数値レンジ評価。

```json
{
  "type": "scale",
  "title": "満足度",
  "min": 1,
  "max": 5,
  "minLabel": "不満",
  "maxLabel": "満足"
}
```

#### 8. date

日付入力。

```json
{
  "type": "date",
  "title": "入社日を入力してください"
}
```

#### 9. time

時刻入力。

```json
{
  "type": "time",
  "title": "起床時間を入力してください"
}
```

#### 10. grid

行×列のマトリクス形式。

```json
{
  "type": "grid",
  "title": "各用語に対する理解度を入力してください。",
  "rows": ["オブジェクト志向", "デザインパターン", "クリーンアーキテクチャ"],
  "columns": "理解度チェック"   // choicesTemplatesを参照可能
}
```

#### 11. checkboxgrid

複数選択可能なマトリクス形式。

```json
{
  "type": "checkboxgrid",
  "title": "好きなプログラミング言語を選んでください",
  "rows": ["フロントエンド", "バックエンド", "モバイル"],
  "columns": ["Java", "Python", "Go", "JavaScript"]
}
```
