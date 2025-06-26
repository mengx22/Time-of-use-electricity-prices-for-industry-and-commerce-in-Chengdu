import pandas as pd
import os

def xlsxToEfile():
    """
    读取四川省省级代理购电电工商业用户电价.xlsx文件的两张表
    并返回电价数值和电价区间的DataFrame
    
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

def save_both_tables(df_prices, df_ranges, output_path):
    """
    将两个DataFrame保存到同一个文件中
    
    参数:
        df_prices (pd.DataFrame): 电价数值表
        df_ranges (pd.DataFrame): 电价区间表
        output_path (str): 输出文件的路径
    """
    try:
        # 准备输出内容
        with open(output_path, 'w', encoding='utf-8') as f:
            # 写入电价数值表
            f.write("<xx>\n")
            f.write("@ " + " ".join(list(df_prices.columns)) + "\n")
            f.write("# " + df_prices.to_string(index=False))
            f.write("\n</xx>\n")
            
            # 写入电价区间表
            f.write("<xx>\n")
            f.write("@ " + " ".join(list(df_ranges.columns)) + "\n")
            f.write("# " + df_ranges.to_string(index=False))
            f.write("\n</xx>")
            
        print(f"文件已成功保存到: {output_path}")
    except Exception as e:
        print(f"保存文件时发生错误: {str(e)}")

def main():
    # 读取两张表的数据
    df_prices, df_ranges = xlsxToEfile()
    
    if df_prices is not None and df_ranges is not None:
        # 获取当前工作目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 将两张表保存到同一个文件中
        output_path = os.path.join(current_dir, "电价数据.database")
        save_both_tables(df_prices, df_ranges, output_path)

if __name__ == "__main__":
    main()