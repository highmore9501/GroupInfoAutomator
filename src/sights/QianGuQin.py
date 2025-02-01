from docx import Document
from datetime import datetime
from src.utils.utils import copy_row_and_insert, clear_table, replace_in_docx, replace_img_in_docx
import os


def process_table(doc_path, new_guests, output_path):
    doc = Document(doc_path)

    # 定位目标表格（假设是第一个包含6列的表格）
    target_table = None
    for table in doc.tables:
        if len(table.columns) == 6:  # 根据列数识别表格
            target_table = table
            break

    if not target_table:
        raise ValueError("未找到符合条件的表格")

    # 清空第2-6行的数据
    clear_table(target_table, 2, 6)

    # 如果guests数量少于等于4个，直接填充客人信息
    if len(new_guests) <= 4:
        for idx, (name, id_num, region) in enumerate(new_guests, start=2):
            target_table.rows[idx].cells[0].text = str(idx - 1)
            target_table.rows[idx].cells[1].text = name
            target_table.rows[idx].cells[2].text = id_num
            target_table.rows[idx].cells[3].text = region
            target_table.rows[idx].cells[4].text = ""
            target_table.rows[idx].cells[5].text = "门票+14：00演出"
        doc.save(output_path)
        return

    # 如果guests数量大于4个，复制足够数量的第2行，插入到表格中
    copy_row_and_insert(target_table, 2, len(new_guests) - 5)

    # 填充客人信息
    for idx, (name, id_num, region) in enumerate(new_guests, start=1):
        target_table.rows[idx].cells[0].text = str(idx)
        target_table.rows[idx].cells[1].text = name
        target_table.rows[idx].cells[2].text = id_num
        target_table.rows[idx].cells[3].text = region
        target_table.rows[idx].cells[4].text = ""
        target_table.rows[idx].cells[5].text = "门票+14：00演出"

    # 保存修改后的文档
    doc.save(output_path)


def editQianGuQinDocx(guests, driver, visit_date, driver_photo_path):
    """
    编辑千古情报备文件
    """
    today = datetime.now().strftime("%Y年%m月%d日")
    template_path = "templates/千古情.docx"
    guest_number = len(guests)
    temp_doc_path = "out/temp.docx"
    output_path = f"out/{driver['司机']}_{visit_date}三亚千古情.docx"
    replace_dict = {
        "胡成": driver["司机"],
        "4人小包团": f"{guest_number}人小包团",
        "18562219199": driver["电话"],
        "琼BD96963": driver["车牌号"],
        "2月1日": visit_date,
        "2025年1月30日": today,
        "370602197812251632": driver["身份证"],
    }
    # 替换文件中的内容
    replace_in_docx(
        input_path=template_path,
        output_path=temp_doc_path,
        replacements=replace_dict
    )
    # 处理表格
    process_table(
        doc_path=temp_doc_path,
        new_guests=[(g["name"], g["id"], g["location"]) for g in guests],
        output_path=output_path
    )
    # 替换docx中的图片
    replace_img_in_docx(
        output_path, "word\media\image2.jpeg", driver_photo_path)

    # 删除掉中间文件

    os.remove(temp_doc_path)
    return f"千古情报备文件已生成：{output_path},默认是14:00场次"


# 使用示例
if __name__ == "__main__":
    guests = [
        {"name": "陶常磊", "id": "130638198709108535",
            "phone": "15192765518", "location": "随便填的"},
        {"name": "付佳", "id": "130638198904160048",
            "phone": "15192765517", "location": "只是测试"},
        {"name": "陶禹翰", "id": "13063820181203003X",
            "phone": "", "location": "野猪林"},
        {"name": "陶常磊", "id": "130638198709108535",
         "phone": "15192765518", "location": "随便填的"},
        {"name": "付佳", "id": "130638198904160048",
            "phone": "15192765517", "location": "只是测试"},
    ]
    driver = {"司机": "由爱华", "电话": "13889471002",
              "车牌号": "琼BD02144", "身份证": "21021119820303512X"}
    driver_photo_path = "C:/Users/BigHippo78/Pictures/由爱华.png"
    visit_date = "2月3日"

    editQianGuQinDocx(
        doc_path="templates/千古情.docx",
        guests=guests,
        driver=driver,
        driver_photo_path=driver_photo_path,
        visit_date=visit_date
    )
