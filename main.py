import argparse
import glob
import os
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
        "-o", "--output-dir", type=str, default="./", help="結合したpdfを出力するディレクトリ"
    )
    parser.add_argument(
        "-d", "--delete-tmp-files", action="store_true", help="中間ファイルを削除するかどうか"
    )

    args = parser.parse_args()

    input_dir = args.input_dir
    word_files_dir = f"{input_dir}/tmp/word"
    pdf_files_dir = f"{input_dir}/tmp/pdf"
    output_file = f"{args.output_dir}/out.pdf"

    if os.path.exists(word_files_dir):
        shutil.rmtree(word_files_dir)
    if os.path.exists(pdf_files_dir):
        shutil.rmtree(pdf_files_dir)

    os.makedirs(word_files_dir, exist_ok=True)
    os.makedirs(pdf_files_dir, exist_ok=True)

    print(f"Reading docx files in: {input_dir}")
    for word_file in glob.glob(f"{input_dir}/*/*.docx"):
        joined_word_file = "".join(word_file.split("/")[-2:])
        shutil.copy(word_file, f"{word_files_dir}/{joined_word_file}")

    print("Converting docx files to pdf...")
    docx2pdf.convert(word_files_dir, output_path=pdf_files_dir, keep_active=True)

    merger = pypdf.PdfMerger()
    for pdf_file in sorted(glob.glob(f"{pdf_files_dir}/*.pdf")):
        merger.append(pdf_file)
    merger.write(output_file)

    print(f"Successed, created merged pdf file: {output_file}")

    if args.delete_tmp_files:
        shutil.rmtree(word_files_dir)
        shutil.rmtree(pdf_files_dir)
