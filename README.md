# word-to-pdf

フォルダ内の Word ファイル（`.doc` / `.docx`）を一括で PDF に変換する Windows 用ツールです。

## 使い方

1. `word_to_pdf.py` を変換したい Word ファイルと同じフォルダに置く
2. `word_to_pdf.py` をダブルクリックして実行
3. 完了すると、PDF がそのフォルダに生成され、元の Word ファイルは `converted/` フォルダに移動します

```
📁 実行前                       📁 実行後
├── word_to_pdf.py              ├── word_to_pdf.py
├── 資料A.docx                  ├── 資料A.pdf        ← 生成
└── 報告書B.docx                ├── 報告書B.pdf      ← 生成
                                └── converted/
                                      ├── 資料A.docx  ← 移動
                                      └── 報告書B.docx
```

## 必要な環境

| 必要なもの | 備考 |
|---|---|
| Windows | Mac・Linux 不可 |
| Microsoft Word | 変換の本体として使用 |
| Python 3.x | [python.org](https://www.python.org/) からインストール。インストール時に「Add Python to PATH」にチェックを入れること |
| pywin32 | 初回実行時に自動インストールを試みます |

## 仕組み

Python の `pywin32` ライブラリを通じて Microsoft Word を操作し、PDF として保存しています。変換品質は Word で手動保存したときと同じです。

## 注意事項

- 変換に失敗したファイルは `converted/` に移動しません（元の場所に残ります）
- `converted/` に同名ファイルが既にある場合、タイムスタンプを付けてリネームされます（上書きされません）
- Windowsのスマートアプリコントロールの制限により、`.bat` ファイルではなく `.py` ファイルとして配布しています
