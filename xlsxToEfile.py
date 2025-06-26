import pandas as pd
import os

def create_units_df():
    """
    创建电价单位表
    
    返回:
        pd.DataFrame: 包含电价各列单位的DataFrame
    """
    # 定义单位数据
    units = ['kV', '元/千瓦时', '元/千瓦时', '元/千瓦时', '元/千瓦时', '元/千瓦·月', '元/千伏安·月']
    # 创建DataFrame，使用与电价数值表相同的列名
    df_units = pd.DataFrame([units])
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
            # 每行数据前加#号
            for _, row in df_prices.iterrows():
                f.write("# " + " ".join(map(str, row.values)) + "\n")
            f.write("</电价数值>\n\n")
            
            # 写入电价区间表
            f.write("<电价区间>\n")
            f.write("// 尖峰 = 3 高峰 = 2 平段 = 1 低谷 = 0\n")
            f.write("@ " + " ".join(list(df_ranges.columns)) + "\n")
            # 每行数据前加#号
            for _, row in df_ranges.iterrows():
                f.write("# " + " ".join(map(str, row.values)) + "\n")
            f.write("</电价区间>")
            
        print(f"文件已成功保存到: {output_path}")
    except Exception as e:
        print(f"保存文件时发生错误: {str(e)}")

def main():
    # 读取表的数据
    df_prices, df_ranges = xlsxToEfile()
    
    if df_prices is not None and df_ranges is not None:
        # 创建单位表
        df_units = create_units_df()
        # 使用电价数值表的后几列作为单位表的列名（忽略时间相关的列）
        df_units.columns = df_prices.columns[2:]  # 跳过前两列（年份和月份）
        
        # 获取当前工作目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 将所有表保存到同一个文件中，使用Qs格式
        output_path = os.path.join(current_dir, "电价数据.Qs")
        save_all_tables(df_units, df_prices, df_ranges, output_path)

if __name__ == "__main__":
    main()