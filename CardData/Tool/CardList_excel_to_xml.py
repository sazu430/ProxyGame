
from openpyxl import load_workbook
from lxml import etree
from pathlib import Path
import sys

# exeまたはPythonスクリプトの配置場所
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent
# ファイルパス
EXCEL_PATH = BASE_DIR / "CardList.xlsx"
XSD_PATH = BASE_DIR / "Card.xsd"
OUTPUT_PATH = BASE_DIR / "CardList.xml"

OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
# XMLのルート要素
root = etree.Element("Card")

# Excel読み込み
workbook = load_workbook(EXCEL_PATH, data_only=True)
sheet = workbook["CardList"]

# ID重複チェック用
used_ids = set()

# 2行目から読み込む（1行目はヘッダー）
for row in sheet.iter_rows(min_row=2, values_only=True):

    # IDが空なら行を無視
    if row[0] is None:
        continue

    card_id = str(row[0]).strip()

    # ID重複チェック
    if card_id in used_ids:
        raise ValueError(f"IDが重複しています: {card_id}")

    used_ids.add(card_id)

    # カード要素を生成
    card = etree.SubElement(
        root,
        "Card",
        id=card_id,
        type=str(row[2]).strip(),
        rarity=str(row[3]).strip()
    )

    # 子要素の生成
    fields = [
        ("Name", row[1]),
        ("ActionCost", row[4]),
        ("Description", row[7])
    ]

    for tag, value in fields:
        if value is None:
            raise ValueError(
                f"{card_id}: {tag} が未入力です"
            )

        element = etree.SubElement(card, tag)
        element.text = str(value)

# XML生成
xml_tree = etree.ElementTree(root)

# XSD読み込み
schema_doc = etree.parse(XSD_PATH)
schema = etree.XMLSchema(schema_doc)

# XSD検証
if not schema.validate(xml_tree):
    print("XML検証エラー")
    print(schema.error_log)
    raise SystemExit(1)

# 検証成功時のみXML出力
xml_tree.write(
    OUTPUT_PATH,
    encoding="UTF-8",
    xml_declaration=True,
    pretty_print=True
)

print("XML生成成功:", OUTPUT_PATH)
print("カード枚数:", len(used_ids))
