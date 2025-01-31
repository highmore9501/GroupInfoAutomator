import gradio as gr
from src.infos.getGuestsInfo import getGuestsInfo
from src.infos.getTripInfo import getTripInfo
from src.infos.getDriverInfo import getDriverInfo
from src.sights.QianGuQin import editQianGuQinDocx
from src.sights.HouDao import editHouDaoDocx
from src.searchSight import searchSight


def generate_report(trip_info, guest_info, driver_info, driver_photo_path):
    report = ""
    guests = getGuestsInfo(guest_info)
    if not guests:
        report += "客人信息有误，请查看终端输出结果\n"
        return report
    trip_info_dict = getTripInfo(trip_info)
    flight_info = trip_info_dict.get('航班信息', [])
    hotel_info = trip_info_dict.get('酒店信息', [])
    attraction_info = trip_info_dict.get('景点信息', [])
    car_info = trip_info_dict.get('用车信息', [])
    driver = getDriverInfo(driver_info)

    qianGuQing = searchSight("千古情", attraction_info)
    if not qianGuQing:
        report += "未找到千古情景点信息\n"
    else:
        visit_date = qianGuQing['日期']
        report += editQianGuQinDocx(guests, driver, visit_date) + "\n"

    houDao = searchSight("猴岛", attraction_info)
    if not houDao:
        report += "未找到猴岛景点信息\n"
    else:
        visit_date = houDao['日期']
        report += editHouDaoDocx(guests, driver, visit_date) + "\n"

    # yaTe = searchSight("亚特兰蒂斯")
    # if not yaTe:
    #     report += "未找到亚特兰蒂斯景点信息\n"
    # else:
    #     report += editYaTeDocx(
    #         flight_info, hotel_info, attraction_info, car_info, guests, driver, driver_photo_path, yaTe) + "\n"

    return report


with gr.Blocks() as demo:
    gr.Markdown("## 报备信息录入")

    trip_info = gr.Textbox(label="团队行程录入处")
    guest_info = gr.Textbox(label="客人信息录入处")
    driver_info = gr.Textbox(label="司机信息录入处")

    output = gr.Textbox(label="输出信息")

    generate_button = gr.Button("生成报备文件")
    generate_button.click(generate_report, inputs=[
                          trip_info, guest_info, driver_info], outputs=output)

demo.launch()
