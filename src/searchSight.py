from typing import Dict, Any, List, Optional


def searchSight(sight: str, attraction_info: List[Dict[str, Any]]) -> bool:
    """
    在行程的景点信息中搜索指定的景点信息，返回是否包含，如果包含的话返回日期
    """
    for item in attraction_info:
        if any(sight in sight_item for sight_item in item['景点']):
            return item
    return None
