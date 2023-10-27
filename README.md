# docxファイルをpdf化し、結合して一つのpdfにするやつ

1. `./data/{適当な名前}/`に一括ダウンロードしたフォルダ（解凍済み）を配置
2.
```
python -m venv ./.venv
source .venv/bin/activate
pip3 install -r requirements.txt
python3 main.py -i ./data/{適当な名前}/

（途中でアクセス許可が求められることがある）
```
3. `./out.pdf`が作成される

```
python3 main.py -h

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