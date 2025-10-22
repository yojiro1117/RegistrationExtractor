"""
gui.py
------

Streamlit を用いた GUI アプリケーション。本アプリでは PDF ファイルをドラッグ＆ドロップすることで登記簿を解析し、Excel に出力します。OCR の有効／無効や出力ファイル名の指定、進捗表示などを備えています。
"""
from __future__ import annotations

import io
import tempfile
from pathlib import Path
from typing import List, Dict, Any

import streamlit as st

from ..core.cmap_unicode import normalize_text
from ..core.router import classify_document
from ..core.text_pdfminer import extract_text_pdfminer
from ..core.text_pymupdf import extract_text_pymupdf
from ..core.text_pypdf2 import extract_text_pypdf2
from ..core.land_parser import parse_land
from ..core.building_parser import parse_building
from ..core.corporate_parser import parse_corporate
from ..core.writer import write_results_to_excel


def extract_text_with_fallback_bytes(data: bytes) -> str:
    """メモリ上の PDF データに対してテキスト抽出を実施する。"""
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=True) as tmp:
        tmp.write(data)
        tmp.flush()
        text = extract_text_pdfminer(tmp.name)
        if text:
            return text
        text = extract_text_pymupdf(tmp.name)
        if text:
            return text
        return extract_text_pypdf2(tmp.name)


def process_uploaded_file(name: str, data: bytes) -> Dict[str, Any]:
    text = extract_text_with_fallback_bytes(data)
    text_norm = normalize_text(text)
    doc_type = classify_document(text_norm)
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
    result['file_name'] = name
    return result


def main():
    st.set_page_config(page_title='RegistrationExtractor', layout='wide')
    st.title('登記簿 PDF 解析アプリ')
    st.write('PDF をドラッグ＆ドロップして解析を実行し、所有者一覧を Excel に出力します。')

    uploaded_files = st.file_uploader('PDF ファイルをアップロード（複数可）', type=['pdf'], accept_multiple_files=True)
    enable_ocr = st.checkbox('OCR を有効にする（画像 PDF 用）', value=False)
    output_file_name = st.text_input('出力ファイル名', value='owners.xlsx')

    if st.button('実行'):
        if not uploaded_files:
            st.warning('PDF をアップロードしてください。')
        else:
            progress = st.progress(0)
            results: List[Dict[str, Any]] = []
            total = len(uploaded_files)
            for idx, uploaded in enumerate(uploaded_files):
                # 進捗更新
                progress.progress((idx) / total)
                data = uploaded.read()
                result = process_uploaded_file(uploaded.name, data)
                results.append(result)
                progress.progress((idx + 1) / total)
            # 書き込み
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_excel:
                write_results_to_excel(results, tmp_excel.name)
                tmp_excel.flush()
                tmp_excel_path = Path(tmp_excel.name)
            st.success('解析が完了しました。')
            st.write(f'処理件数: {len(results)} 件')
            # プレビュー表示 (先頭 5 行)
            try:
                import pandas as pd  # type: ignore
                from openpyxl import load_workbook  # type: ignore
                wb = load_workbook(tmp_excel_path)
                ws = wb.active
                data = list(ws.values)
                headers = data[0]
                rows = data[1:6]
                df = pd.DataFrame(rows, columns=headers)
                st.dataframe(df)
            except Exception:
                pass
            # ダウンロードリンク
            with open(tmp_excel_path, 'rb') as f:
                st.download_button('Excel をダウンロード', f, file_name=output_file_name, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    main()