import argparse
import os
import pathlib
import shutil

import docx2pdf
import pypdf

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-i",
        "--input-dir",
        type=str,
        required=True,
        help="一括ダウンロードしたpdfファイルを配置したディレクトリ",
    )
    parser.add_argument(
        "-o",
        "--output-file",
        type=str,
        default="./",
        help="結合したpdfを出力するファイル",
    )
    parser.add_argument(
        "-d",
        "--delete-tmp-files",
        action="store_true",
        help="中間ファイルを削除するかどうか",
    )

    args = parser.parse_args()

    input_dir = pathlib.Path(args.input_dir)
    word_files_dir = input_dir / "tmp" / "word"
    pdf_files_dir = input_dir / "tmp" / "pdf"
    output_file = pathlib.Path(args.output_file)

    if os.path.exists(word_files_dir):
        shutil.rmtree(word_files_dir)
    if os.path.exists(pdf_files_dir):
        shutil.rmtree(pdf_files_dir)

    os.makedirs(word_files_dir, exist_ok=True)
    os.makedirs(pdf_files_dir, exist_ok=True)

    print(f"Reading docx files in: {input_dir}")
    for word_file in input_dir.glob("*/*.docx"):
        joined_word_file = word_file.parent.name + "-" + word_file.name
        # たまに拡張子「.docx」を2つ付けたファイルが提出されている場合がある
        # 拡張子が正しくないと、docx2pdfがエラーを吐くので、不要な拡張子を削除する
        if joined_word_file.count(".docx") == 2:
            joined_word_file = joined_word_file[:-5]
        shutil.copy(word_file, word_files_dir / joined_word_file)

    print("Converting docx files to pdf...")
    docx2pdf.convert(word_files_dir, output_path=pdf_files_dir, keep_active=True)

    merger = pypdf.PdfMerger()
    for pdf_file in sorted(pdf_files_dir.glob("*.pdf")):
        merger.append(pdf_file)
    merger.write(output_file)

    print(f"Successed, created merged pdf file: {output_file}")

    if args.delete_tmp_files:
        shutil.rmtree(word_files_dir)
        shutil.rmtree(pdf_files_dir)
