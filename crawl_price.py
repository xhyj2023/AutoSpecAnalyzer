import requests
import pandas as pd
import time

# ===== 材质字典 =====
material_list = [
    {"materialId": 3, "material": "304"},
    {"materialId": 7, "material": "201/J1"},
    {"materialId": 9, "material": "201/J2"},
    {"materialId": 18, "material": "201/J5"},
    {"materialId": 27, "material": "430"},
    {"materialId": 28, "material": "316L"},
    {"materialId": 29, "material": "201"},
    {"materialId": 30, "material": "S32001"},
    {"materialId": 31, "material": "201/J3"},
    {"materialId": 32, "material": "201/J1A"}
]

# ===== 请求基础信息 =====
url = "http://admin.tomals.com/Api/Price/getCompareList"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Content-Type": "application/x-www-form-urlencoded"
}

# ===== 常量 =====
TOKEN = "7ZpK0G0OQvKkpMo06BvQkLm9WwZierceVJIwQyqo+valiY6QWwIQtuv0vCHwuwWF"
DATE = "2025-10-17"

# ===== 存储所有结果 =====
all_records = []

# ===== 动态生成 SN =====
def generate_sn():
    return str(int(time.time()))  # 返回当前时间戳（秒）

# ===== 获取分页数据 =====
def fetch_data_with_pagination(material_id, material_name):
    page = 1  # 从第一页开始
    while True:
        print(f"正在抓取：{material_name}，第 {page} 页")

        SN = generate_sn()  # 每次请求时生成新的 SN
        payload = {
            "materialId": material_id,
            "time": DATE,
            "companyId": "",  # 不指定公司
            "apiType": 1,
            "token": TOKEN,
            "sn": SN,
            "pv": "web",
            "page": page,  # 分页参数
            "pageSize": 20  # 每页抓取 20 条数据，根据实际情况调整
        }

        try:
            response = requests.post(url, data=payload, headers=headers)
            print("状态码:", response.status_code)
            print("前200字符:", response.text[:200])

            result = response.json()

            if response.status_code == 200 and result.get("status") == 1:
                data_list = result.get("data", [])
                if not data_list:
                    print(f"⚠️ 材质 {material_name} 第 {page} 页 返回空数据")
                    break

                # 处理数据
                for item in data_list:
                    # 获取
                    type_name = item.get("typeName", "未知规格")
                    company_name = item.get("companyName", "未知公司")
                    thickness = item.get("thickness", "未知厚度")
                    place = item.get("place", "未知产地")
                    adjust_status = item.get("adjustStatus", "未知状态")
                    price = item.get("price", "未知价格")
                    price_id = item.get("priceId", "未知编号")

                    # 打印调试输出：抓取到的完整数据
                    print(f"抓取到公司：{company_name}, 规格：{type_name}, 厚度：{thickness}, 价格：{price}")

                    all_records.append({
                        "材质": material_name,
                        "materialId": material_id,
                        "规格": type_name,  # 添加规格信息
                        "公司": company_name,
                        "厚度": thickness,
                        "产地": place,
                        "状态": adjust_status,
                        "价格": price,
                        "编号": price_id,
                        "时间": DATE
                    })

                page += 1  # 获取下一页数据
            else:
                print(f"❌ {material_name} 返回错误: {result.get('msg')}")
                break

        except requests.exceptions.RequestException as e:
            print(f"请求异常 - {material_name} 第 {page} 页: {e}")
            break
        except Exception as e:
            print(f"未知异常 - {material_name} 第 {page} 页: {e}")
            break

        time.sleep(0.3)  # 控制请求频率

# ===== 主循环：按材质遍历 =====
for material_item in material_list:
    material_id = material_item["materialId"]
    material_name = material_item["material"]
    fetch_data_with_pagination(material_id, material_name)

# ===== 去重和保存结果 =====
if all_records:
    df = pd.DataFrame(all_records)
    df = df.drop_duplicates(subset=["公司", "材质", "规格", "厚度", "价格"], keep="first")  # 去重时加上规格
    df.to_excel("compare_data.xlsx", index=False)
    print(f"\n✅ 所有数据已保存到 compare_data.xlsx，共 {len(df)} 条记录")
else:
    print("\n⚠️ 没有抓取到任何数据，请检查 token 或接口参数")
