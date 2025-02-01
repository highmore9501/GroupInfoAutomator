from typing import List, Dict


def getDriverInfo(driver_info: str) -> List[Dict[str, str]]:
    """
    解析司机信息
    """
    try:
        driver_info = driver_info.strip().split("\n")
        driver_info = [info.strip().split("：") for info in driver_info]
        driver_info = dict(driver_info)
        return driver_info
    except Exception as e:
        print(f"Error parsing driver info: {driver_info}. Error: {e}")
        return {}


if __name__ == "__main__":
    driver_info = """
    司机：胡成
    电话：18562219199
    车牌号：琼BD96963
    身份证：370602197812251632
    """

    print(getDriverInfo(driver_info))
