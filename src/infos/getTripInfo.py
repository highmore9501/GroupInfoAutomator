from typing import Dict, Any, List, Optional
import json


def parse_flight_info(flight_str: str) -> List[Dict[str, str]]:
    flights = []
    lines = [line.strip() for line in flight_str.split('\n') if line.strip()]
    for line in lines:
        try:
            parts = line.split()
            date = parts[0]
            departure, arrival = parts[1].split('-')
            start_time, end_time = parts[2].split('-')
            flight_number = parts[3]
            flights.append({
                '日期': date,
                '出发地': departure,
                '到达地': arrival,
                '起飞时间': start_time,
                '到达时间': end_time,
                '航班号': flight_number
            })
        except Exception as e:
            print(f"Error parsing flight info line: {line}. Error: {e}")
    return flights


def parse_hotel_info(hotel_str: str) -> List[Dict[str, str]]:
    hotels = []
    lines = [line.strip() for line in hotel_str.split('\n') if line.strip()]
    for line in lines:
        try:
            date_part, hotel_part = line.split(' ', 1)
            hotels.append({
                '日期': date_part,
                '酒店': hotel_part.strip()
            })
        except Exception as e:
            print(f"Error parsing hotel info line: {line}. Error: {e}")
    return hotels


def parse_attraction_info(attraction_str: str) -> List[Dict[str, Any]]:
    attractions = []
    lines = [line.strip()
             for line in attraction_str.split('\n') if line.strip()]
    for line in lines:
        try:
            date, _, items_str = line.partition(' ')
            items = [item.strip() for item in items_str.split('+')]
            attractions.append({
                '日期': date,
                '景点': items
            })
        except Exception as e:
            print(f"Error parsing attraction info line: {line}. Error: {e}")
    return attractions


def parse_car_info(car_str: str) -> List[Dict[str, str]]:
    cars = []
    lines = [line.strip() for line in car_str.split('\n') if line.strip()]
    for line in lines:
        try:
            parts = [p.strip() for p in line.split('|')]
            date_part = parts[0].split()[0]
            car_type = parts[0].split(' ', 1)[1]
            passengers = parts[1].split('·')[0].strip()
            luggage = parts[1].split('·')[1].strip()
            duration = parts[2]
            cars.append({
                '日期': date_part,
                '车型': car_type,
                '乘客': passengers,
                '行李': luggage,
                '用车时长': duration
            })
        except Exception as e:
            print(f"Error parsing car info line: {line}. Error: {e}")
    return cars


def getTripInfo(trip_info: str) -> Optional[Dict[str, Any]]:
    """
    解析行程信息
    """
    trip_info_dict = {}
    blocks = trip_info.strip().split("\n\n")

    for block in blocks:
        block = block.strip()
        if not block:
            continue
        key, value = block.split("：", 1)
        value = value.strip()
        if key == '航班信息':
            trip_info_dict[key] = parse_flight_info(value)
        elif key == '酒店信息':
            trip_info_dict[key] = parse_hotel_info(value)
        elif key == '景点信息':
            trip_info_dict[key] = parse_attraction_info(value)
        elif key == '用车信息':
            trip_info_dict[key] = parse_car_info(value)
        else:
            trip_info_dict[key] = value

    return trip_info_dict


if __name__ == "__main__":
    trip_info = """
    航班信息：
    1月13日 青岛-三亚 07:00-13:10 CA4779
    1月18日 三亚-青岛 06:50-12:40 HU7663

    酒店信息：
    1月13日 三亚亚特兰蒂斯酒店 (海景大床房 ¥2350 x 1间 x 1晚)
    1月14日-1月16日 三亚亚龙湾爱琴海套房度假酒店 (开放式海景亲子套房简单厨房 ¥928 x 1间 x 3晚)
    1月17日 维也纳酒店(三亚湾店) (豪华大床房（观景阳台.宽敞明亮.超大空间） ¥510 x 1间 x 1晚)

    景点信息：
    1月13日 欢迎来到海南+温馨提示+美食推荐+三亚亚特兰蒂斯失落的空间水族馆
    1月14日 三亚亚特兰蒂斯水世界+三亚亚特兰蒂斯C秀
    1月15日 蜈支洲岛+太阳湾路
    1月16日 南山文化旅游区+凤凰岛直升机基地+小东海赶海体验
    1月17日 亚龙湾热带天堂森林公园+三亚游艇旅游中心
    1月18日 三亚送机

    用车信息：
    1月15日 经济5座  | 乘客5 · 行李2 | 用车10h
    1月16日 经济5座  | 乘客5 · 行李2 | 用车10h
    1月17日 经济5座  | 乘客5 · 行李2 | 用车10h

    """

    trip_data = getTripInfo(trip_info)
    print(json.dumps(trip_data, ensure_ascii=False, indent=2))
