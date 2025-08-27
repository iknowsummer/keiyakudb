import tempfile
import subprocess
import sys
import shutil
import os
from io import BytesIO
from datetime import datetime

from docxtpl import DocxTemplate


# メモリ上にWordファイルを生成してBytesIOで返す関数
def generate_docx_stream(template_path, context):
    """
    Word契約書をテンプレートから生成し、BytesIOで返す

    :param template_path: テンプレートファイルのパス
    :param context: 差し込みデータを格納した辞書
    :return: BytesIOオブジェクト
    """
    tpl = DocxTemplate(template_path)
    tpl.render(context)
    file_like = BytesIO()
    tpl.save(file_like)
    file_like.seek(0)
    return file_like


# docxのストリームをPDFストリームに変換する関数
def convert_stream_docx2pdf(docx_stream, timeout_sec: int = 20):
    """
    WordドキュメントのストリームをPDFに変換してBytesIOで返す
    LibreOffice(soffice)をサブプロセスで呼び出して変換

    :param docx_stream: WordドキュメントのBytesIOオブジェクト
    :return: PDFのBytesIOオブジェクト
    """

    # 一時ファイル・一時ディレクトリ作成
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "input.docx")
        pdf_path = os.path.join(tmpdir, "input.pdf")

        # docx_streamを一時ファイルに書き出し
        with open(docx_path, "wb") as f:
            f.write(docx_stream.getvalue())

        # OSごとにsofficeコマンドのパスを決定
        if sys.platform.startswith("win"):
            soffice_cmd = (
                shutil.which("soffice.exe")
                or "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
            )
        else:
            soffice_cmd = shutil.which("soffice") or "/usr/bin/soffice"

        # LibreOfficeでPDF変換（ヘッドレス）
        cmd = [
            soffice_cmd,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            tmpdir,
            docx_path,
        ]
        try:
            subprocess.run(
                cmd,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=timeout_sec,
            )
        except subprocess.TimeoutExpired:
            raise RuntimeError("PDF変換がタイムアウトしました")
        except Exception as e:
            raise RuntimeError(f"PDF変換失敗: {e}")

        # 変換後のPDFをBytesIOで返す
        if not os.path.exists(pdf_path):
            raise FileNotFoundError("PDF変換に失敗しました")
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()
        return BytesIO(pdf_bytes)
