import http.client
import ssl
import json
import urllib.parse
import pandas as pd


def make_request(method, url):
    conn.request(method, url)
    res = conn.getresponse()
    data = res.read()
    return json.loads(data.decode('utf-8'))


region_input = input("Введите регион (пример: Пензенская область): ")
city_input = input("Введите город (пример: Пенза): ")
street_input = input("Введите улицу (пример: Плеханова): ")
house_start = int(input("Введите номер дома (старт): "))
house_finish = int(input("Введите номер дома (финиш): "))
apart_start = int(input("Введите номер квартиры (старт): "))
apart_finish = int(input("Введите номер квартиры (финиш): "))

context = ssl.create_default_context()
context.check_hostname = False
context.verify_mode = ssl.CERT_NONE
conn = http.client.HTTPSConnection("rosreestr.gov.ru", context=context)

regions = make_request('GET', '/api/online/macro_regions/')
region_id = [m['id'] for m in regions if m['name'].lower() == region_input.lower()][0]
cities = make_request("GET", f'/api/online/regions/{region_id}')
city_id = [m['id'] for m in cities if m['name'].lower() == city_input.lower()][0]

result = []
for house in range(house_start, house_finish + 1):
    for apart in range(apart_start, apart_finish + 1):
        print(f"Обработка дома {house} квартиры {apart}.")
        try:
            information = make_request("GET",
                                       f"https://rosreestr.gov.ru/api/online/address/fir_objects/?macroRegionId={region_id}&regionId={city_id}&street={urllib.parse.quote(street_input)}&house={house}&building=&structure=&apartment={apart}")
            print(f"Информация получена! Суммарно: {len(information)}")
            idx = 0
            for info in information:
                idx += 1
                print(f"Загрузка объектов. № {idx} / {len(information)}")
                obj_id = info['objectId']
                object_info = make_request("GET", f"https://rosreestr.gov.ru/api/online/fir_object/{obj_id}/")
                objectData = object_info['objectData']
                objectAddress = objectData['objectAddress']
                parcelData = object_info['parcelData']
                del object_info['parcelData']
                del objectData['objectAddress']
                del object_info['objectData']
                try:
                    object_info.update(parcelData)
                except:
                    pass
                try:
                    object_info.update(objectData)
                except:
                    pass

                try:
                    object_info.update(objectAddress)
                except:
                    pass
                result.append(object_info)
        except:
            break

df = pd.DataFrame(result)
df.to_excel("result.xlsx", index=True)
