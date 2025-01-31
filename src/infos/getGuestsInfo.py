from typing import List, Dict
from id_validator import validator


def getGuestsInfo(guests_info: str) -> List[Dict[str, str]]:
    guests = []
    try:
        for guest in guests_info.strip().split("\n"):
            guest = guest.strip().split(" ")
            id_number = guest[1]
            if not validator.is_valid(id_number):
                print(f"Invalid ID number: {id_number}")
                continue
            guest_info = {
                "name": guest[0],
                "id": id_number,
                "phone": guest[2] if len(guest) == 3 else "",
                "location": validator.get_info(id_number)['address']
            }
            guests.append(guest_info)
    except Exception as e:
        print(f"Error parsing guest info: {guests_info}. Error: {e}")
    return guests


if __name__ == "__main__":
    guests_info = """
    于双民 370283198312204515 15192765518
    林柳 371083198501205525 15192765517
    于林溪 370211201710162640
    """

    print(getGuestsInfo(guests_info))
