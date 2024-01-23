import openpyxl
import matplotlib.pyplot as plt
import numpy as np

# 用户输入文件名（不含后缀）
file_name = input("请输入文件名（不含后缀）: ")

# 构建完整的文件路径
file_path = f"C:\\Users\\ThinkPad\\SynologyDrive\\Trademe\\{file_name}.xlsx"

try:
    # 读取Excel文件
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # 获取销量和日期列数据
    sales_column = sheet['D'][1:]  # 第一行是标题，所以从第二行开始读取
    date_column = sheet['A'][1:]

    # 将Excel日期转换为Python日期
    dates = [date.value for date in date_column]

    # 获取销量数据
    sales = [sale.value for sale in sales_column]

    # 处理缺货，用均值回归填充
    def fill_missing_sales(sales):
        for i in range(len(sales)):
            if sales[i] == 0:
                start_idx = max(0, i - 7)
                end_idx = min(len(sales), i + 7)
                non_zero_values = [v for v in sales[start_idx:end_idx] if v != 0]
                if non_zero_values:
                    sales[i] = np.mean(non_zero_values)
        return sales

    sales = fill_missing_sales(sales)

    # 计算均线
    def calculate_moving_average(data, window_size):
        weights = np.ones(window_size) / window_size
        return np.convolve(data, weights, mode='valid')

    # 绘制销量和均线图
    plt.figure(figsize=(10, 6))
    plt.plot(dates, sales, label='Sales')
    plt.plot(dates[9:], calculate_moving_average(sales, 10), label='10-day MA')
    plt.plot(dates[29:], calculate_moving_average(sales, 30), label='30-day MA')
    plt.plot(dates[89:], calculate_moving_average(sales, 90), label='90-day MA')
    plt.plot(dates[179:], calculate_moving_average(sales, 180), label='180-day MA')
    plt.plot(dates[359:], calculate_moving_average(sales, 360), label='360-day MA')

    plt.xlabel('Date')
    plt.ylabel('Sales')
    plt.title('Sales and Moving Averages')
    plt.legend()
    plt.show()

except FileNotFoundError:
    print(f"文件 '{file_name}.xlsx' 未找到。请检查文件路径是否正确。")
except Exception as e:
    print(f"发生错误：{e}")

finally:
    if 'workbook' in locals():
        # 关闭Excel文件
        workbook.close()
