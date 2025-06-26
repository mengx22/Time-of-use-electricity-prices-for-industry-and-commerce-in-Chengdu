import pandas as pd
import os

def create_units_df():
    """
    创建电价单位表
    
    返回:
        pd.DataFrame: 包含电价各列单位的DataFrame
    """
    # 定义单位数据
    units = ['年', '月', 'kV', '元/千瓦时', '元/千瓦时', '元/千瓦时', '元/千瓦时', '元/千瓦·月', '元/千伏安·月']
    columns = ['年份', '月份', '电压等级', '尖峰', '高峰', '平段', '低谷', '需量电价', '容量电价']
    # 创建DataFrame，使用完整的列名
    df_units = pd.DataFrame([units], columns=columns)
    return df_units

def xlsxToEfile():
    """
    读取四川省省级代理购电电工商业用户电价.xlsx文件的两张表
    并返回电价数值表和电价区间表的DataFrame
    
    返回:
        tuple: (电价数值DataFrame, 电价区间DataFrame)
    """
    try:
        # 获取当前工作目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(current_dir, "四川省省级代理购电电工商业用户电价.xlsx")
        
        # 读取两张表
        df_prices = pd.read_excel(file_path, sheet_name='2022-2025分时电价值')  # 第一张表：电价数值
        df_ranges = pd.read_excel(file_path, sheet_name='2022-2025分时区间')  # 第二张表：电价区间
        
        print(f"成功读取文件: {os.path.basename(file_path)}")
        print("\n电价数值表预览:")
        print(df_prices.head())
        print("\n电价区间表预览:")
        print(df_ranges.head())
        
        return df_prices, df_ranges
    except Exception as e:
        print(f"读取文件时发生错误: {str(e)}")
        return None, None

def save_to_special_format(df, output_path):
    """
    将DataFrame保存为特定格式的文件
    
    参数:
        df (pd.DataFrame): 要保存的DataFrame数据
        output_path (str): 输出文件的路径
    """
    try:
        # 获取列名作为特性
        features = list(df.columns)
        
        # 准备输出内容
        with open(output_path, 'w', encoding='utf-8') as f:
            # 写入开始标记
            f.write("<xx>\n")
            
            # 写入特性部分（特性用空格分隔）
            f.write("@ " + " ".join(features) + "\n")
            
            # 写入数据部分（直接跟数据）
            f.write("# ")
            # 将DataFrame转换为字符串，去除索引，对齐数据
            data_str = df.to_string(index=False)
            f.write(data_str)
            f.write("\n</xx>")
            
        print(f"文件已成功保存到: {output_path}")
    except Exception as e:
        print(f"保存文件时发生错误: {str(e)}")

def save_all_tables(df_units, df_prices, df_ranges, output_path):
    """
    将所有DataFrame保存到同一个文件中
    
    参数:
        df_units (pd.DataFrame): 单位表
        df_prices (pd.DataFrame): 电价数值表
        df_ranges (pd.DataFrame): 电价区间表
        output_path (str): 输出文件的路径
    """
    try:
        # 转换年份和月份为整数格式
        df_prices = df_prices.copy()
        df_ranges = df_ranges.copy()
        
        # 电价数值表格式转换
        df_prices['年份'] = df_prices['年份'].astype(float).astype(int)
        df_prices['月份'] = df_prices['月份'].astype(float).astype(int)
        
        # 电价区间表格式转换（注意这里月份列名可能是'月'）
        month_col = '月' if '月' in df_ranges.columns else '月份'
        df_ranges['年份'] = df_ranges['年份'].astype(float).astype(int)
        df_ranges[month_col] = df_ranges[month_col].astype(float).astype(int)
        
        # 准备输出内容
        with open(output_path, 'w', encoding='utf-8') as f:
            # 写入单位表
            f.write("<电价单位>\n")
            f.write("@ " + " ".join(list(df_units.columns)) + "\n")
            # 每行数据前加#号
            for _, row in df_units.iterrows():
                f.write("# " + " ".join(map(str, row.values)) + "\n")
            f.write("</电价单位>\n\n")
            
            # 写入电价数值表
            f.write("<电价数值>\n")
            f.write("@ " + " ".join(list(df_prices.columns)) + "\n")
            # 每行数据前加#号，确保年份和月份为整数格式
            for _, row in df_prices.iterrows():
                # 对于每个值，如果是年份或月份列，确保使用整数格式，其他数值保持原有精度
                formatted_values = []
                for col, val in zip(df_prices.columns, row.values):
                    if col in ['年份', '月份']:
                        formatted_values.append(str(int(float(val))))
                    else:
                        formatted_values.append(str(val))
                f.write("# " + " ".join(formatted_values) + "\n")
            f.write("</电价数值>\n\n")
            
            # 写入电价区间表
            f.write("<电价区间>\n")
            f.write("// 尖峰 = 3 高峰 = 2 平段 = 1 低谷 = 0\n")
            f.write("@ " + " ".join(list(df_ranges.columns)) + "\n")
            # 每行数据前加#号，确保年份和月份为整数格式
            for _, row in df_ranges.iterrows():
                # 对于每个值，如果是年份或月列，确保使用整数格式，其他数值保持原有格式
                formatted_values = []
                for col, val in zip(df_ranges.columns, row.values):
                    if col in ['年份', '月', '月份']:
                        formatted_values.append(str(int(float(val))))
                    else:
                        formatted_values.append(str(val))
                f.write("# " + " ".join(formatted_values) + "\n")
            f.write("</电价区间>")
            
        print(f"文件已成功保存到: {output_path}")
    except Exception as e:
        print(f"保存文件时发生错误: {str(e)}")

def create_climate_units_df():
    """
    创建气候数据单位表
    
    返回:
        pd.DataFrame: 包含气候数据各列单位的DataFrame
    """
    # 定义单位数据（根据实际气候数据的列调整）
    units = ['年', '月', '℃', '℃', 'mm', '天', '%', '小时', '小时']
    columns = ['年份', '月份', '平均高温', '平均低温', '降雨量', '降雨天数', '湿度', '日照时长', '阳光时长']
    # 创建DataFrame
    df_units = pd.DataFrame([units], columns=columns)
    return df_units

def process_climate_data(file_path):
    """
    读取四川气候数据Excel文件并处理
    
    参数:
        file_path (str): 气候数据Excel文件的路径
    
    返回:
        tuple: (气候数据DataFrame, 气候数据单位DataFrame)
    """
    try:
        # 读取气候数据表
        df_climate = pd.read_excel(file_path)
        
        # 数据清理和规范化
        # 1. 确保年份和月份为整数
        df_climate['年份'] = df_climate['年份'].astype(int)
        df_climate['月份'] = df_climate['月份'].astype(int)
        
        # 2. 确保数值列为浮点数
        numeric_columns = ['平均高温', '平均低温', '降雨量', '降雨天数', '湿度', '日照时长', '阳光时长']
        for col in numeric_columns:
            df_climate[col] = pd.to_numeric(df_climate[col], errors='coerce')
        
        # 3. 删除包含空值的行
        df_climate = df_climate.dropna()
        
        # 4. 按年份和月份排序
        df_climate = df_climate.sort_values(['年份', '月份'])
        
        # 创建单位表
        df_units = create_climate_units_df()
        
        print("气候数据处理完成")
        print("\n气候数据表预览:")
        print(df_climate.head())
        
        return df_climate, df_units
        
    except Exception as e:
        print(f"处理气候数据时发生错误: {str(e)}")
        return None, None

def save_climate_data(df_climate, df_units, output_path):
    """
    将气候数据保存为Qs格式文件
    
    参数:
        df_climate (pd.DataFrame): 气候数据DataFrame
        df_units (pd.DataFrame): 单位DataFrame
        output_path (str): 输出文件路径
    """
    try:
        # 创建数据副本以进行格式转换
        df_climate = df_climate.copy()
        
        # 确保年份和月份为整数格式
        df_climate['年份'] = df_climate['年份'].astype(float).astype(int)
        df_climate['月份'] = df_climate['月份'].astype(float).astype(int)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            # 写入单位表
            f.write("<气候单位>\n")
            f.write("@ " + " ".join(list(df_units.columns)) + "\n")
            for _, row in df_units.iterrows():
                f.write("# " + " ".join(map(str, row.values)) + "\n")
            f.write("</气候单位>\n\n")
            
            # 写入气候数据表
            f.write("<气候数据>\n")
            f.write("@ " + " ".join(list(df_climate.columns)) + "\n")
            # 每行数据前加#号，确保年份和月份为整数格式
            for _, row in df_climate.iterrows():
                # 对于每个值，如果是年份或月份列，确保使用整数格式
                formatted_values = []
                for col, val in zip(df_climate.columns, row.values):
                    if col in ['年份', '月份']:
                        formatted_values.append(str(int(float(val))))
                    else:
                        formatted_values.append(str(val))
                f.write("# " + " ".join(formatted_values) + "\n")
            f.write("</气候数据>")
            
        print(f"气候数据文件已成功保存到: {output_path}")
    except Exception as e:
        print(f"保存气候数据文件时发生错误: {str(e)}")

def climate_xlsx_to_efile(input_file=None):
    """
    将四川气候数据Excel文件转换为Qs格式文件
    
    参数:
        input_file (str): 输入文件路径，如果为None则使用默认文件名
    """
    try:
        # 获取当前工作目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 如果没有指定输入文件，使用默认文件名
        if input_file is None:
            input_file = os.path.join(current_dir, "四川气候数据.xlsx")
        
        # 处理气候数据
        df_climate, df_units = process_climate_data(input_file)
        
        if df_climate is not None and df_units is not None:
            # 设置输出文件路径
            output_path = os.path.join(current_dir, "气候数据.Qs")
            # 保存为Qs格式
            save_climate_data(df_climate, df_units, output_path)
            
    except Exception as e:
        print(f"转换气候数据文件时发生错误: {str(e)}")

def main():
    # 1. 处理电价数据
    print("\n处理电价数据...")
    df_prices, df_ranges = xlsxToEfile()
    
    if df_prices is not None and df_ranges is not None:
        # 创建单位表（已包含所有列的单位）
        df_units = create_units_df()
        
        # 获取当前工作目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 将所有表保存到同一个文件中，使用Qs格式
        output_path = os.path.join(current_dir, "电价数据.Qs")
        save_all_tables(df_units, df_prices, df_ranges, output_path)
        print("电价数据处理完成")
    
    # 2. 处理气候数据
    print("\n处理气候数据...")
    climate_xlsx_to_efile()

if __name__ == "__main__":
    main()