import os
from dotenv import load_dotenv
import asyncio
import aiohttp
import openpyxl

load_dotenv()
api_key = os.getenv("LOSTARK_API_KEY")
excel_directory = os.getenv("LOCAL_EXCEL_DIRECTORY")

# API 호출을 위한 URL과 헤더 설정
url = "https://developer-lostark.game.onstove.com/markets/items"
headers = {
    "accept": "application/json",
    "authorization": f"Bearer {api_key}"
}

''' 
특정 아이템의 고유 코드로 마켓 가격 정보 조회 parameters
아래는 full format임
'''
# data = {
#     "Sort": "GRADE",
#     "CategoryCode": 90300,
#     "CharacterClass": "바드",
#     "ItemTier": None,
#     "ItemGrade": "희귀",
#     "ItemName": "아비도스 목재",
#     "PageNo": 0,
#     "SortCondition": "ASC"
# }


# 거래소 최저가 출력
async def fetch_item_min_price(session, category_code, item_name):
    # 생활 재료는 CategoryCode와 ItemName만 있으면 정상적으로 동작
    data = {
        "CategoryCode": category_code,
        "ItemName": item_name
    }
    
    async with session.post(url, headers = headers, json = data) as response:
        # HTTP Code 200: OK
        if response.status == 200:
            response_data = await response.json()
            items = response_data.get("Items", [])
            
            if items:
                return items[0].get("CurrentMinPrice")
            else:
                return f"Not found data for {item_name}"

        else:
            return f"Error: {response.status} for {item_name}"


async def fetch_multiple_items(items_list):
    async with aiohttp.ClientSession() as session:
        tasks = []
        for item in items_list:
            task = fetch_item_min_price(session, item['category_code'], item['item_name'])
            tasks.append(task)
            
        prices = await asyncio.gather(*tasks)
        print(prices)
        return prices

#category_code: 50010 - 재련 재료 / 90200 - 식물채집 전리품 / 90300 - 벌목 전리품 / 90400 - 채광 전리품 / 90500 - 수렵 전리품 / 90600 - 낚시 전리품 / 90700 - 고고학 전리품 / 90800 - 기타
items_list = [
    {"category_code": 50010, "item_name": "오레하 융화 재료"},
    {"category_code": 50010, "item_name": "상급 오레하 융화 재료"},
    {"category_code": 50010, "item_name": "최상급 오레하 융화 재료"},
    {"category_code": 50010, "item_name": "아비도스 융화 재료"},
    {"category_code": 90300, "item_name": "목재"},
    {"category_code": 90400, "item_name": "철광석"},
    {"category_code": 90200, "item_name": "들꽃"},
    {"category_code": 90600, "item_name": "생선"},
    {"category_code": 90500, "item_name": "두툼한 생고기"},
    {"category_code": 90700, "item_name": "고대 유물"},
    {"category_code": 90200, "item_name": "수줍은 들꽃"},
    {"category_code": 90400, "item_name": "묵직한 철광석"},
    {"category_code": 90300, "item_name": "부드러운 목재"},
    {"category_code": 90600, "item_name": "붉은 살 생선"},
    {"category_code": 90700, "item_name": "희귀한 유물"},
    {"category_code": 90500, "item_name": "다듬은 생고기"},
    {"category_code": 90700, "item_name": "아비도스 유물"},
    {"category_code": 90400, "item_name": "아비도스 철광석"},
    {"category_code": 90200, "item_name": "아비도스 들꽃"},
    {"category_code": 90600, "item_name": "아비도스 태양 잉어"},
    {"category_code": 90500, "item_name": "아비도스 두툼한 생고기"},
    {"category_code": 90600, "item_name": "오레하 태양 잉어"},
    {"category_code": 90500, "item_name": "오레하 두툼한 생고기"},
    {"category_code": 90700, "item_name": "오레하 유물"},
    {"category_code": 90300, "item_name": "아비도스 목재"}
]

file_name = "_융화 재료 손익.xlsx"
file_path = os.path.join(excel_directory, file_name)
def save_prices_to_excel(prices, file_path = file_path):
    workbook = openpyxl.load_workbook(file_path, read_only = False)
    sheet = workbook.active
    
    # Q9 오레하 융화 재료부터 Q12 아비도스 융화 재료까지 4개의 가격 채워넣음
    # 이후 Z14 목재부터 Z34 아비도스 목재까지 21개의 가격을 채워넣음
    for i, price in enumerate(prices):
        if i <= 3:
            cell = f"Q{9 + i}"
        else:
            cell = f"Z{10 + i}"
        sheet[cell] = price
        
    workbook.save(file_path)
    
if __name__ == "__main__":
    prices = asyncio.run(fetch_multiple_items(items_list))
    save_prices_to_excel(prices, file_path)