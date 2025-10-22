"""
cli.py
-------

コマンドラインインターフェイスを提供します。フォルダまたはファイルを指定して複数の PDF を一括処理し、解析結果を Excel に出力します。

使い方の例:

```bash
python -m app.cli --input "C:\\登記PDF" --output "C:\\owners.xlsx" --enable-ocr --workers 6
```
"""
from __future__ import annotations

import argparse
import concurrent.futures
import logging
import os
from pathlib import Path
from typing import List, Dict, Any

from ..core.router import classify_document
from ..core.text_pdfminer import extract_text_pdfminer
from ..core.text_pymupdf import extract_text_pymupdf
from ..core.text_pypdf2 import extract_text_pypdf2
from ..core.cmap_unicode import normalize_text
from ..core.land_parser import parse_land
from ..core.building_parser import parse_building
from ..core.corporate_parser import parse_corporate
from ..core.writer import write_results_to_excel


def collect_pdf_files(input_path: Path) -> List[Path]:
    """入力パスから PDF ファイルのリストを取得する。フォルダの場合は再帰的に探索する。
    """
    files: List[Path] = []
    if input_path.is_file() and input_path.suffix.lower() == '.pdf':
        return [input_path]
    for root, _, filenames in os.walk(input_path):
        for fname in filenames:
            if fname.lower().endswith('.pdf'):
                files.append(Path(root) / fname)
    return files


def extract_text_with_fallback(path: str) -> str:
    """pdfminer → PyMuPDF → PyPDF2 の順でテキスト抽出を試みる。"""
    text = extract_text_pdfminer(path)
    if text:
        return text
    text = extract_text_pymupdf(path)
    if text:
        return text
    return extract_text_pypdf2(path)


def process_file(pdf_path: Path, enable_ocr: bool = False) -> Dict[str, Any]:
    """単一 PDF を解析して結果を返す。"""
    logging.info(f"Processing {pdf_path}")
    text = extract_text_with_fallback(str(pdf_path))
    # Unicode 正規化と PUA 置換
    text_norm = normalize_text(text)
    doc_type = classify_document(text_norm)
    result: Dict[str, Any]
    if doc_type == 'land':
        result = parse_land(text_norm)
    elif doc_type == 'building':
        result = parse_building(text_norm)
    elif doc_type == 'corporate':
        result = parse_corporate(text_norm)
    else:
        result = {
            'type': 'unknown',
            'header': {},
            'owners': [],
            'accident_flag': False,
            'accident_memo': '未分類',
            'record_entries': [],
        }
    result['file_name'] = pdf_path.name
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description='登記簿 PDF 解析ツール')
    parser.add_argument('--input', required=True, help='入力ファイルまたはディレクトリを指定')
    parser.add_argument('--output', required=True, help='出力 Excel ファイル (owners.xlsx など)')
    parser.add_argument('--enable-ocr', action='store_true', help='OCR を有効にする')
    parser.add_argument('--workers', type=int, default=4, help='並列処理ワーカー数 (デフォルト 4)')
    parser.add_argument('--log', default='runlog.txt', help='ログファイル出力先')
    args = parser.parse_args()

    logging.basicConfig(filename=args.log, level=logging.INFO,
                        format='%(asctime)s %(levelname)s %(message)s', encoding='utf-8')

    input_path = Path(args.input)
    files = collect_pdf_files(input_path)
    logging.info(f"Found {len(files)} PDF files in {args.input}")
    results: List[Dict[str, Any]] = []
    if args.workers and args.workers > 1 and len(files) > 1:
        with concurrent.futures.ThreadPoolExecutor(max_workers=args.workers) as executor:
            future_to_path = {executor.submit(process_file, f, args.enable_ocr): f for f in files}
            for future in concurrent.futures.as_completed(future_to_path):
                path = future_to_path[future]
                try:
                    res = future.result()
                    results.append(res)
                except Exception as e:
                    logging.error(f"Failed to process {path}: {e}")
    else:
        for f in files:
            try:
                res = process_file(f, args.enable_ocr)
                results.append(res)
            except Exception as e:
                logging.error(f"Failed to process {f}: {e}")
    # 書き込み
    write_results_to_excel(results, args.output)
    logging.info(f"Finished. Results written to {args.output}")


if __name__ == '__main__':
    main()