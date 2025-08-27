from io import BytesIO
from docxtpl import DocxTemplate
from datetime import datetime
import os


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


# ファイル保存用ラッパー
def generate_docx_file(template_path, output_dir, context):
    """
    Word契約書をテンプレートから生成し、日付+時刻サフィックス付きで保存

    :param template_path: テンプレートファイルのパス
    :param output_dir: 出力ディレクトリ
    :param context: 差し込みデータを格納した辞書
    :return: 保存したファイルのパス
    """
    # 日付＋時刻のサフィックスを生成
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    template_basename = os.path.splitext(os.path.basename(template_path))[0]
    output_filename = f"{template_basename}_{timestamp}.docx"
    output_path = os.path.join(output_dir, output_filename)

    os.makedirs(output_dir, exist_ok=True)
    file_like = generate_docx_stream(template_path, context)

    with open(output_path, "wb") as f:
        f.write(file_like.getvalue())

    return output_path
