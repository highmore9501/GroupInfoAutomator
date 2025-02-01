from src.utils.utils import copy_xlsx_row_and_insert, replace_in_xlsx, count_guests, get_trip_slice
import shutil
from openpyxl import load_workbook


def editYaTeDocx(
        flight_info, attraction_info, guests, driver, visit_date):

    template_path = "templates/亚特兰蒂斯.xlsx"
    driver_name = driver['司机']
    driver_phone = driver['电话']
    driver_car = driver['车牌号']
    target_path = f"out/{driver_name}+({visit_date})亚特兰蒂斯小包团(望海旅行社).xlsx"

    # 复制模板文件
    shutil.copyfile(template_path, target_path)

    outbound_flight = flight_info[0]['航班号']
    outbound_date = flight_info[0]['日期']
    outbound_from = flight_info[0]['出发地']
    outbound_to = flight_info[0]['到达地']
    outbound_info = f"{outbound_date} {outbound_flight} {outbound_from}-{outbound_to}"

    inbound_flight = flight_info[1]['航班号']
    inbound_date = flight_info[1]['日期']
    inbound_from = flight_info[1]['出发地']
    inbound_to = flight_info[1]['到达地']
    inbound_info = f"{inbound_date} {inbound_flight} {inbound_from}-{inbound_to}"

    guest_row_count_in_template = 6  # 模板中的客人信息行数
    trip_row_count_in_template = guest_row_count_in_template - 2  # 模板中的行程信息行数

    # 计算需要写在表格中的行程信息
    trip_slice = get_trip_slice(
        attraction_info, visit_date, trip_row_count_in_template)

    new_guest_count = len(guests)

    adult_count, child_count = count_guests(guests)
    total_guest_count = f"{adult_count}大{child_count}小"

    if new_guest_count > 6:  # 如果客人数量大于6
        insert_row_count = new_guest_count - guest_row_count_in_template
        # 复制第5行，插入到表格中
        copy_xlsx_row_and_insert(target_path, 5, insert_row_count)

    # 替换表格中的内容
    replacements = {
        "胡成": driver_name,
        "18562219199": driver_phone,
        "琼BD96963": driver_car,
        "1/31 FM9521 上海-三亚": outbound_info,
        "2/5 C7328 亚龙湾-海口": inbound_info,
        "3大1小 ": total_guest_count,
    }

    replace_in_xlsx(target_path, replacements)

    # 获取客人联系方式，一个就行
    guest_contact = ""
    for guest in guests:
        if guest['phone']:
            guest_contact = f"{guest['name']}{guest['phone']}"
            break

    wb = load_workbook(target_path)
    sheet = wb['Sheet1']

    # 填充客人联系方式
    if guest_contact:
        sheet["C3"] = guest_contact

    # 填充客人信息
    for i in range(new_guest_count):
        guest = guests[i]
        name = guest['name']
        id_num = guest['id']
        cell_address = f"E{i+3}"
        sheet[cell_address] = f"{name}{id_num}"

    # 填充行程信息
    for j in range(trip_row_count_in_template):
        date_address = f"G{j+4}"
        sight_address = f"H{j+4}"
        sight_text = "+".join(trip_slice[j]['景点'])
        sight_text = sight_text.replace("三亚", "")
        sight_text = sight_text.replace("海南", "")
        sight_text = sight_text.replace("欢迎来到+", "")
        sight_text = sight_text.replace("温馨提示+", "")
        sight_text = sight_text.replace("美食推荐+", "")
        sight_text = sight_text.replace("文化旅游区", "")
        sight_text = sight_text.replace("凤凰岛", "")
        sight_text = sight_text.replace("赶海体验", "")
        sheet[date_address] = f"{trip_slice[j]['日期']}"
        sheet[sight_address] = sight_text

    wb.save(target_path)

    return f"亚特兰蒂斯报备文件已生成: {target_path}，默认票种是水世界+水族馆"
