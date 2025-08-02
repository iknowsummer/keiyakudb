from docxtpl import DocxTemplate
from datetime import datetime
import os


def generate_docx(template_path, output_dir, context):
    """
    Word契約書をテンプレートから生成し、日付+時刻サフィックス付きで保存

    :param template_path: テンプレートファイルのパス
    :param output_dir: 出力ディレクトリ
    :param context: 差し込みデータを格納した辞書
    """
    # テンプレート読み込み
    tpl = DocxTemplate(template_path)

    # 差し込みデータ適用
    tpl.render(context)

    # 日付＋時刻のサフィックスを生成
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # テンプレートファイル名（拡張子なし）を取得
    template_basename = os.path.splitext(os.path.basename(template_path))[0]

    # 出力先ファイル名を作成
    output_filename = f"{template_basename}_{context['client_name']}_{timestamp}.docx"
    output_path = os.path.join(output_dir, output_filename)

    # 出力先フォルダが存在しなければ作成
    os.makedirs(output_dir, exist_ok=True)

    # Wordファイルとして保存
    tpl.save(output_path)

    print(f"✅ 契約書を生成しました: {output_path}")


# --- 使用例 ---
if __name__ == "__main__":
    template_path = "./template/outsourcing.docx"  # テンプレートファイル
    output_dir = "./dist"  # 出力フォルダ

    # 差し込みデータ
    context = {
        "company_name": "株式会社サンプル",
        "company_address": "東京都港区1-2-3",
        "representative_name": "代表取締役 佐藤一郎",
        "client_name": "山田太郎",
        "client_address": "大阪府大阪市中央区1-2-3",
        "service_description": "Webサイト制作業務",
        "price": "150000",
        "payment_method": "銀行振込",
        "start_date": "2025年8月10日",
        "end_date": "2025年12月31日",
        "year": "7",
        "month": "8",
        "day": "2",
    }

    # 契約書生成
    generate_docx(template_path, output_dir, context)
