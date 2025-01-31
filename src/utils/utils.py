from docx import Document


def copy_row(source_row, target_row):
    """将源行的内容和格式复制到目标行"""
    for src_cell, tgt_cell in zip(source_row.cells, target_row.cells):
        # 复制文本
        tgt_cell.text = src_cell.text

        # 复制段落格式（如字体、大小、加粗等）
        source_para = src_cell.paragraphs[0]
        target_para = tgt_cell.paragraphs[0]

        # 复制段落样式
        target_para.style = source_para.style

        # 复制字体属性
        source_font = source_para.runs[0].font
        target_font = target_para.runs[0].font
        target_font.name = source_font.name
        target_font.size = source_font.size
        target_font.bold = source_font.bold
        target_font.italic = source_font.italic


def copy_row_and_insert(table, source_row_index, copy_count):
    """
    复制源行并插入到表格中
    """
    source_row = table.rows[source_row_index]

    for i in range(copy_count):  # 插入新行

        new_row = table.add_row()
        # 获取新行和源行的XML元素
        new_row_element = new_row._element
        source_row_element = source_row._element
        # 将新行插入到源行的位置（源行会自动后移）
        source_row_element.getparent().insert(source_row_index+1, new_row_element)

        # 复制源行的内容和格式到新行
        copy_row(source_row, table.rows[source_row_index])


def clear_table(target_table, start_row, end_row):
    """
    清空表格中指定行的数据
    """
    for i in range(start_row, end_row):
        for cell in target_table.rows[i].cells:
            cell.text = ""


def replace_in_docx(input_path, output_path, replacements):
    """
    在Word文档中执行批量替换
    """
    doc = Document(input_path)

    def replace_text(container):
        # 处理跨Run的文本替换
        for paragraph in container.paragraphs:
            merged_text = ''.join([run.text for run in paragraph.runs])
            if any(old in merged_text for old in replacements):
                # 执行批量替换
                for old, new in replacements.items():
                    merged_text = merged_text.replace(old, new)

                # 清除原有Run
                for run in paragraph.runs:
                    run.text = ''

                # 添加新Run并保留第一个Run的格式
                if paragraph.runs:
                    paragraph.runs[0].text = merged_text
                else:
                    paragraph.add_run(merged_text)

        # 处理表格中的嵌套表格
        if hasattr(container, "tables"):
            for table in container.tables:
                for row in table.rows:
                    for cell in row.cells:
                        replace_text(cell)

    # 处理文档各组成部分
    replace_text(doc)

    # 处理页眉页脚
    for section in doc.sections:
        replace_text(section.header)
        replace_text(section.footer)

    doc.save(output_path)
