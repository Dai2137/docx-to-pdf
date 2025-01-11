# docxファイルをpdf化し、結合して一つのpdfにするやつ

```
> python3 --version
Python 3.10.4
```

1. `./data/{適当な名前}/`に一括ダウンロードしたフォルダ（解凍済み）を配置
2. 以下をターミナルで実行（途中でアクセス許可が求められることがある）
    ```
    python -m venv .venv
    source .venv/Scripts/activate
    pip install -r requirements.txt
    python main.py -i ./data/{適当な名前}/
    ```
3. `./out.pdf`が作成される

## 設定

```
python main.py -h

usage: main.py [-h] -i INPUT_DIR [-o OUTPUT_DIR] [-d]

options:
  -h, --help            show this help message and exit
  -i INPUT_DIR, --input-dir INPUT_DIR
                        一括ダウンロードしたpdfファイルを配置したディレクトリ
  -o OUTPUT_DIR, --output-dir OUTPUT_DIR
                        結合したpdfを出力するディレクトリ
  -d, --delete-tmp-files
                        中間ファイルを削除するかどうか
```

## デモ

```
python main.py -i ./demo/
```

## 不具合メモ

- zipファイルが解凍できない場合
- 以下のコマンドを実行すると治るかも

```
# a.zipが壊れている場合
zip -FF a.zip --out a-fixed.zip
unzip -O CP932 a-fixed.zip
```

- 参考
  - https://qiita.com/ktateish/items/28b69c18d64d50d6d5b0
