from docx import Document
import xml.etree.ElementTree as ET
import json

target_dir = "samples"


def export_template_info(template_path, fileName):
    output = []
    doc = Document(template_path)
    for para in doc.paragraphs:
        xml_content = para._element.xml
        output.append(xml_content)

    # 保存到json文件
    with open(f"{target_dir}/{fileName}.json", "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=4)


if __name__ == "__main__":
    template_path = "templates/千古情.docx"
    fileName = "千古情"
    export_template_info(template_path, fileName)
