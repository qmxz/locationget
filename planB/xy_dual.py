import xlrd
import requests
import xml.etree.ElementTree as ET
import xlwt
import time
import threading
import os
from datetime import datetime
#powered by lqy
#for xy
# 定义参数及并发数
api_key='<your api key>'
api_url = 'https://restapi.amap.com/v3/geocode/geo'
concurrent_requests = 15

# 打开XLS表格
workbook = xlrd.open_workbook('address.xls')
worksheet = workbook.sheet_by_index(0)
# 创建"results"子文件夹
if not os.path.exists('results'):
    os.makedirs('results')

# 获取当前时间戳，用于生成文件名
current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
# 创建新的XLS表格来存储结果
output_workbook = xlwt.Workbook()
output_sheet = output_workbook.add_sheet('output_sheet')

# 写入标题行
output_sheet.write(0, 0, 'source address')
output_sheet.write(0, 1, 'city')
output_sheet.write(0, 2, 'current address')
output_sheet.write(0, 3, 'longitude')
output_sheet.write(0, 4, 'latitude')

# 定义线程处理函数
def process_data(row_index):
    cell_value_1 = worksheet.cell_value(row_index, 0)
    cell_value_2 = worksheet.cell_value(row_index, 1)
    # 构建API请求的参数字符串
    params_str = f'key={api_key}&address={cell_value_1}&city={cell_value_2}&output=XML'
    full_api_url = f'{api_url}?{params_str}'
    
    # 调用Web API获取XML数据
    response = requests.get(full_api_url)
    xml_data = response.content
    
 # 解析XML数据获取所需的值
    root = ET.fromstring(xml_data)
    # 假设目标值在<result>标签下的<value>元素中
    result_element1 = root.find('geocodes/geocode/formatted_address')
    if result_element1 is not None:
             target_value1 = result_element1.text
    else:
             target_value1 = "Not Found"  # 如果找不到目标值，则给出一个默认值
    result_element2 = root.find('geocodes/geocode/location')
    if result_element2 is not None:
        location_str = result_element2.text
        # 以逗号为分隔符拆分字符串
        longitude, latitude = location_str.split(',')
    else:
        longitude, latitude = "Not Found", "Not Found"  # 如果找不到目标值，则给出默认值
    # 将目标值写入新表格
    output_sheet.write(row_index, 0, cell_value_1)
    output_sheet.write(row_index, 1, cell_value_2)
    output_sheet.write(row_index, 2, target_value1)
    output_sheet.write(row_index, 3, longitude)
    output_sheet.write(row_index, 4, latitude)

# 定义主函数
def main():
    print("-----------\n")
    print("Powered By LQY\n")
    print("For XY\n")
    print("-----------\n")
    print("开始...")
    start_time = time.time()

    # 遍历第二行及以后的数据（忽略标题行）
    for row_index in range(1, worksheet.nrows):  # 从索引1开始，忽略标题行
        # 控制并发数
        while threading.active_count() >= concurrent_requests:
            time.sleep(0.1)
        # 创建并启动新线程
        threading.Thread(target=process_data, args=(row_index,)).start()

    # 等待所有线程执行完成
    while threading.active_count() > 1:
        time.sleep(0.1)

    # 获取结果保存的文件路径
    result_file = os.path.join('results', f'results_{current_time}.xls')
    
    # 保存结果到新的XLS表格
    output_workbook.save(result_file)

    end_time = time.time()
    print("完成!")
    print("总耗时：", end_time - start_time, "秒")
    input("按任意键终止程序...")
if __name__ == '__main__':
    main()
