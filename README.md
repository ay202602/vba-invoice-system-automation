# vba-invoice-system-automation

## 概要
売上データがあるシートから顧客名を絞り込み、請求書シートへ明細を自動転記するマクロです。
手作業のコピー & ペーストを、ボタン1クリックに置き換えます。

> [!NOTE]
> `sample/invoice.xlsm` のデータはすべて架空であり、実在する企業・団体とは無関係です。

## 機能

| 機能 | 内容 | 実装 |
|:---|:---|:---|
| 明細の転記 | 選択した顧客の売上データをすべて請求書シートへ値として貼り付け | `CInvoice.CellPaste` |
| 絞り込み | 売上シートのC列(顧客名)でオートフィルタを適用 | `CInvoice.FilterOn` |
| 貼り付け前のクリア | 前回の明細(A5:G17)を消去してから転記 | `CInvoice.CellClear` |
| 実行ボタン | 図形にマクロを割り当て、クリックで実行 | `CreateInvoice.CreateInvoice` |

### エラー処理

| 条件 | 表示メッセージ |
|:---|:---|
| B2セルが空欄 | 顧客名の入力を促すエラー |
| B2の顧客名が売上データに存在しない | 「(顧客名)は売上リストの顧客名には存在しません」 |
| 売上シートにデータがない | 売上リストが存在しない旨のエラー |

## 実行結果
![実行結果](image/result.png)

## 実行方法
1. ワークシート「請求書」B2セルの顧客名をプルダウンで選択（売上データ連動）
2. 「請求書作成」と書かれた図形（ボタン）をクリック
3. 該当する顧客名のデータが下の表（A5:G17）に転記されます

## 動作環境・使用技術
> **動作環境**： Microsoft Excel Home and Business 2024 （マクロの有効化が必要）

| 技術 | 内容 | 目的 | 該当箇所 |
|:---|:---|:---|:---|
| クラスモジュール | 検索 / コピー / 貼り付け / フィルタ をクラスに集約 | 保守性の向上 | `CInvoice.cls` |
| init メソッド | VBAのクラスは `Class_Initialize` に引数を取れないため、`Sub init` で対象セルを受け取る | 初期化処理の分離 | `CInvoice.init` |
| Private 変数(カプセル化) | `mCustomer` `mCellStart` `mPasteCells` を外部から直接触らせない | 意図しない書き換え防止 | `CInvoice.cls` 冒頭 |
| インスタンス変数を参照する Function | `CellSearch` が保持セルを基準に検索し `Range` を返す | 呼び出し側の引数削減 | `CInvoice.CellSearch` |
| 標準モジュールの薄型化 | 入力チェックと呼び出し順の制御のみ担当 | コード量・重複の削減 | `CreateInvoice.bas` |

## ディレクトリ構成

```
vba-invoice-system-automation
├── image
│   └── result.png           # 実行結果画像
├── sample
│   └── invoice.xlsm         # Excelファイル（架空データ）
└── src
    ├── CInvoice.cls         # クラスモジュール
    └── CreateInvoice.bas    # 標準モジュール
```
## invoice.xlsmについて
- 全て架空のデータであり、実際の企業・団体とは無関係です。
