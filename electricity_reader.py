import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
plt.rcParams['figure.max_open_warning'] = 0  # 禁用最大图形警告
from typing import Dict, Optional
from efile_parser import EFileParser

class ElectricityReader:
    """
    电价数据读取器，负责读取和处理电价数据
    """
    
    def __init__(self):
        """
        初始化电价数据读取器
        """
        # 设置中文显示
        plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
        plt.rcParams['axes.unicode_minus'] = False    # 用来正常显示负号
        
        # 获取当前工作目录
        self.current_dir = os.path.dirname(os.path.abspath(__file__))
        # 初始化数据
        self._initialize_data()
        
    def _initialize_data(self):
        """
        初始化电价数据，从文件中读取并解析
        """
        try:
            # 创建EFile解析器实例
            parser = EFileParser()
            
            # 读取电价数据文件
            file_path = os.path.join(self.current_dir, "电价数据.Qs")
            tables = parser.read_file(file_path)
            
            # 获取各个表的数据
            self.units_df = tables.get('电价单位', pd.DataFrame())
            self.prices_df = tables.get('电价数值', pd.DataFrame())
            self.ranges_df = tables.get('电价区间', pd.DataFrame())
            
            # 打印调试信息
            print("\n电价数值表的列名:", self.prices_df.columns.tolist())
            print("电价区间表的列名:", self.ranges_df.columns.tolist())
            
            # 验证数据完整性
            if self.units_df.empty or self.prices_df.empty or self.ranges_df.empty:
                raise ValueError("无法读取完整的电价数据")
            
            # 数据清理和类型转换
            # 确保电价数值表有正确的列
            if not self.prices_df.empty:
                # 检查列名是否包含年和月，如果是其他名称则重命名
                column_mapping = {}
                for col in self.prices_df.columns:
                    if '年' in col:
                        column_mapping[col] = '年份'
                    elif '月' in col:
                        column_mapping[col] = '月份'
                if column_mapping:
                    self.prices_df = self.prices_df.rename(columns=column_mapping)
                    
                # 转换类型并清理数据
                if '年份' in self.prices_df.columns and '月份' in self.prices_df.columns:
                    self.prices_df['年份'] = pd.to_numeric(self.prices_df['年份'], errors='coerce')
                    self.prices_df['月份'] = pd.to_numeric(self.prices_df['月份'], errors='coerce')
                    self.prices_df = self.prices_df.dropna(subset=['年份', '月份'])
                
            # 确保电价区间表有正确的列
            if not self.ranges_df.empty:
                # 检查列名是否包含年和月，如果是其他名称则重命名
                column_mapping = {}
                for col in self.ranges_df.columns:
                    if col == '月':  # 电价区间表中月份列名是'月'
                        column_mapping[col] = '月份'
                    elif '年' in col:
                        column_mapping[col] = '年份'
                if column_mapping:
                    self.ranges_df = self.ranges_df.rename(columns=column_mapping)
                    
                # 转换类型并清理数据
                if '年份' in self.ranges_df.columns and '月份' in self.ranges_df.columns:
                    self.ranges_df['年份'] = pd.to_numeric(self.ranges_df['年份'], errors='coerce')
                    self.ranges_df['月份'] = pd.to_numeric(self.ranges_df['月份'], errors='coerce')
                    self.ranges_df = self.ranges_df.dropna(subset=['年份', '月份'])
                
        except Exception as e:
            print(f"初始化电价数据时发生错误: {str(e)}")
            self.units_df = pd.DataFrame()
            self.prices_df = pd.DataFrame()
            self.ranges_df = pd.DataFrame()

    def get_monthly_electricity_prices(self, year: int, month: int, voltage_level: float) -> Dict[str, pd.DataFrame]:
        """
        获取指定年月和电压等级的电价数据
        
        参数:
            year (int): 年份
            month (int): 月份
            voltage_level (float): 电压等级（kV）
        
        返回:
            Dict[str, pd.DataFrame]: 包含电价单位、电价数值和时段区间的字典
        """
        try:
            # 筛选指定年月的电价数值
            monthly_prices = self.prices_df[
                (self.prices_df['年份'].fillna(-1).astype(int) == year) & 
                (self.prices_df['月份'].fillna(-1).astype(int) == month) &
                (self.prices_df['电压等级'] == voltage_level)
            ]
            
            # 筛选指定年月的时段区间
            monthly_ranges = self.ranges_df[
                (self.ranges_df['年份'].fillna(-1).astype(int) == year) & 
                (self.ranges_df['月份'].fillna(-1).astype(int) == month)
            ]
            
            if monthly_prices.empty or monthly_ranges.empty:
                print(f"未找到{year}年{month}月 {voltage_level}kV电压等级的电价数据")
                return {}
            
            return {
                '电价单位': self.units_df,
                '电价数值': monthly_prices,
                '电价区间': monthly_ranges
            }
            
        except Exception as e:
            print(f"获取{year}年{month}月 {voltage_level}kV电压等级的电价数据时发生错误: {str(e)}")
            return {}

    def plot_price_summary(self, year: int, month: int, voltage_level: float, save_path: Optional[str] = None):
        """
        绘制电价汇总图表
        
        参数:
            year (int): 年份
            month (int): 月份
            voltage_level (float): 电压等级（kV）
            save_path (Optional[str]): 保存图片的路径
        """
        data = self.get_monthly_electricity_prices(year, month, voltage_level)
        if not data:
            return

        # 创建图形
        plt.figure('电价汇总', figsize=(12, 6))
        
        # 创建主坐标系和次坐标系
        ax1 = plt.gca()
        ax2 = ax1.twinx()

        # 准备数据
        prices_df = data['电价数值']
        time_period_cols = ['尖峰', '高峰', '平段', '低谷']
        capacity_cols = ['需量电价', '容量电价']
        
        # 计算柱状图的位置
        x_pos = np.arange(len(time_period_cols))
        
        # 绘制时段电价（左轴）
        bars1 = ax1.bar(x_pos, prices_df.iloc[0][time_period_cols], color=['red', 'orange', 'blue', 'green'])
        
        # 绘制需量和容量电价（右轴）
        x_pos_capacity = np.arange(len(time_period_cols), len(time_period_cols) + len(capacity_cols))
        bars2 = ax2.bar(x_pos_capacity, prices_df.iloc[0][capacity_cols], color=['purple', 'brown'])
        
        # 设置左轴
        ax1.set_ylabel('分时电价（元/千瓦时）', rotation=90, labelpad=15)
        
        # 设置右轴
        ax2.set_ylabel('容量电价（元/千瓦·月 或 元/千伏安·月）', rotation=90, labelpad=15)
        
        # 添加数值标签
        def autolabel(bars, ax):
            for bar in bars:
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2., height,
                       f'{height:.4f}',
                       ha='center', va='bottom')
        
        autolabel(bars1, ax1)
        autolabel(bars2, ax2)
        
        # 设置X轴标签
        plt.xticks(np.arange(len(time_period_cols) + len(capacity_cols)),
                  [*time_period_cols, *capacity_cols],
                  rotation=45, ha='right')
        
        # 设置标题
        plt.title(f'{year}年{month}月 {voltage_level}kV电压等级的电价汇总')
        
        # 调整布局
        plt.tight_layout()
        
        # 保存图片
        if save_path:
            plt.savefig(save_path.replace('.png', '_summary.png'))
            plt.close('电价汇总')

    def plot_time_distribution(self, year: int, month: int, voltage_level: float, save_path: Optional[str] = None):
        """
        绘制分时电价时段分布图
        
        参数:
            year (int): 年份
            month (int): 月份
            voltage_level (float): 电压等级（kV）
            save_path (Optional[str]): 保存图片的路径
        """
        data = self.get_monthly_electricity_prices(year, month, voltage_level)
        if not data:
            return

        # 获取各时段电价
        ranges_df = data['电价区间']
        prices_df = data['电价数值']
        
        # 获取时段值（去掉年份和月份列）
        time_periods = ranges_df.iloc[0].drop(['年份', '月份'])
        
        # 获取各时段电价
        period_prices = {
            '尖峰': prices_df.iloc[0]['尖峰'],
            '高峰': prices_df.iloc[0]['高峰'],
            '平段': prices_df.iloc[0]['平段'],
            '低谷': prices_df.iloc[0]['低谷']
        }
        
        # 创建24小时时间标签
        hours = np.arange(24)
        
        # 创建颜色和时段映射
        colors = ['green', 'blue', 'orange', 'red']
        period_names = ['低谷', '平段', '高峰', '尖峰']
        period_map = {str(i): (name, color) for i, (name, color) in enumerate(zip(period_names, colors))}
        
        # 计算每个小时的电价
        prices = []
        colors_by_hour = []
        for hour in hours:
            period_value = time_periods.iloc[hour]
            period_name = period_map[str(int(period_value))][0]
            period_color = period_map[str(int(period_value))][1]
            prices.append(period_prices[period_name])
            colors_by_hour.append(period_color)
            
        # 创建图形
        plt.figure('分时电价分布', figsize=(12, 6))
        ax = plt.gca()
        
        # 绘制时段电价柱状图
        bars = ax.bar(hours, prices, color=colors_by_hour)
        
        # 添加数值标签
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{height:.4f}',
                   ha='center', va='bottom', fontsize=8)
        
        # 创建自定义图例
        legend_elements = [plt.Rectangle((0,0),1,1, facecolor=color, label=f'{name}({price:.4f}元/千瓦时)')
                         for name, color, price in zip(period_names, colors, 
                             [period_prices[name] for name in period_names])]
        ax.legend(handles=legend_elements, loc='upper right')
        
        # 设置坐标轴
        ax.set_xticks(hours)
        ax.set_xlabel('小时', labelpad=10)
        ax.set_ylabel('电价（元/千瓦时）', rotation=90, labelpad=15)
        ax.set_title(f'{year}年{month}月 {voltage_level}kV电压等级的分时电价分布')
        
        # 调整布局
        plt.tight_layout()
        
        # 保存图片
        if save_path:
            plt.savefig(save_path.replace('.png', '_distribution.png'))
            plt.close('分时电价分布')
        
    def plot_monthly_prices(self, year: int, month: int, voltage_level: float, save_path: Optional[str] = None):
        """
        在两个独立窗口中绘制月度电价数据的所有图表
        
        参数:
            year (int): 年份
            month (int): 月份
            voltage_level (float): 电压等级（kV）
            save_path (Optional[str]): 保存图片的路径，如果不指定则显示图片
        """
        data = self.get_monthly_electricity_prices(year, month, voltage_level)
        if not data:
            return
        
        # 创建第一个窗口（电价汇总）
        plt.figure('电价汇总')
        self.plot_price_summary(year, month, voltage_level, save_path)
        
        # 创建第二个窗口（分时电价分布）
        plt.figure('分时电价分布')
        self.plot_time_distribution(year, month, voltage_level, save_path)
        
        # 如果不保存图片，则显示所有窗口
        if not save_path:
            plt.show()
    
    def get_price_time_series(self, voltage_level: float, price_type: str) -> pd.DataFrame:
        """
        获取指定电压等级和电价类型的所有时间序列数据
        
        参数:
            voltage_level (float): 电压等级（kV）
            price_type (str): 电价类型，可选值：'尖峰', '高峰', '平段', '低谷', '需量电价', '容量电价'
        
        返回:
            pd.DataFrame: 包含时间序列电价数据的DataFrame
        """
        try:
            # 筛选指定电压等级的数据
            voltage_data = self.prices_df[self.prices_df['电压等级'] == voltage_level].copy()
            
            if voltage_data.empty:
                print(f"未找到{voltage_level}kV电压等级的数据")
                return pd.DataFrame()
            
            if price_type not in voltage_data.columns:
                print(f"未找到{price_type}类型的电价数据")
                return pd.DataFrame()
            
            # 对数据按时间排序
            time_series_data = voltage_data.sort_values(['年份', '月份'])
            
            # 将年份和月份先转换为字符串，去掉小数点
            time_series_data['年份'] = time_series_data['年份'].apply(lambda x: str(int(float(x))))
            time_series_data['月份'] = time_series_data['月份'].apply(lambda x: str(int(float(x))).zfill(2))
            
            # 创建日期列，使用格式化的字符串
            time_series_data['日期'] = pd.to_datetime(
                time_series_data['年份'] + '-' + time_series_data['月份'] + '-01',
                format='%Y-%m-%d'
            )
            
            # 选择需要的列并设置日期索引
            result = time_series_data[['日期', price_type]].copy()
            result.columns = ['日期', '电价']
            
            return result
            
        except Exception as e:
            print(f"获取时间序列数据时发生错误: {str(e)}")
            return pd.DataFrame()
    
    def plot_price_trend(self, voltage_levels: float | list, price_type: str, save_path: Optional[str] = None):
        """
        绘制电价趋势图，支持单个电压等级或多个电压等级的对比
        
        参数:
            voltage_levels (float | list): 电压等级（kV）或电压等级列表
            price_type (str): 电价类型，可选值：'尖峰', '高峰', '平段', '低谷', '需量电价', '容量电价'
            save_path (Optional[str]): 保存图片的路径
        """
        # 将单个电压等级转换为列表形式，统一处理
        if isinstance(voltage_levels, (int, float)):
            voltage_levels = [float(voltage_levels)]
        
        plt.figure('电价趋势', figsize=(12, 6))
        ax = plt.gca()
        
        # 用不同的颜色和标记绘制每个电压等级的趋势
        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']
        markers = ['o', 's', '^', 'D', 'v']
        
        # 记录所有数据的时间范围
        all_dates = []
        
        # 绘制每个电压等级的趋势线
        for i, voltage in enumerate(voltage_levels):
            data = self.get_price_time_series(voltage, price_type)
            if not data.empty:
                # 过滤掉电价为0的数据点
                data = data[data['电价'] > 0]
                if not data.empty:
                    color = colors[i % len(colors)]
                    marker = markers[i % len(markers)]
                    
                    # 绘制趋势线
                    ax.plot(data['日期'], data['电价'], 
                           color=color, marker=marker, 
                           linestyle='-', linewidth=2, markersize=6,
                           label=f'{voltage}kV')
                    
                    # 添加数据点标签
                    for x, y in zip(data['日期'], data['电价']):
                        ax.annotate(f'{y:.4f}', 
                                  (x, y), 
                                  textcoords="offset points",
                                  xytext=(0,10), 
                                  ha='center',
                                  color=color,
                                  fontsize=8)
                    
                    all_dates.extend(data['日期'])
        
        if not all_dates:
            print(f"未找到任何有效的{price_type}电价数据")
            plt.close('电价趋势')
            return
            
        # 设置x轴格式
        ax.xaxis.set_major_locator(mdates.YearLocator())
        ax.xaxis.set_minor_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y年'))
        
        # 生成次刻度标签
        minor_labels = []
        for label in ax.get_xticks(minor=True):
            date = mdates.num2date(label)
            minor_labels.append(f'{date.month}月' if date.month != 1 else '')
        ax.set_xticklabels(minor_labels, minor=True)
        
        # 调整标签显示
        plt.setp(ax.get_xticklabels(), rotation=0, ha='center')
        plt.setp(ax.get_xticklabels(minor=True), rotation=0, ha='center', fontsize=8)
        
        # 设置网格
        ax.grid(True, which='major', linestyle='-', alpha=0.5)
        ax.grid(True, which='minor', linestyle=':', alpha=0.2)
        
        # 获取y轴单位
        y_unit = '元/千瓦·月' if price_type == '需量电价' else '元/千伏安·月' if price_type == '容量电价' else '元/千瓦时'
        
        # 设置坐标轴标签
        ax.set_xlabel('时间', labelpad=10)
        ax.set_ylabel(f'电价（{y_unit}）', rotation=90, labelpad=15)
        
        # 设置标题
        start_year = min(d.year for d in all_dates)
        end_year = max(d.year for d in all_dates)
        if len(voltage_levels) == 1:
            title = f'{voltage_levels[0]}kV电压等级 {price_type}的变化趋势\n({start_year}年-{end_year}年)'
        else:
            title = f'不同电压等级的{price_type}变化趋势对比\n({start_year}年-{end_year}年)'
        plt.title(title)
        
        # 如果有多个电压等级，添加图例
        if len(voltage_levels) > 1:
            ax.legend(loc='best')
        
        # 调整布局
        plt.tight_layout()
        
        # 保存或显示图片
        if save_path:
            suffix = '_voltage_comparison' if len(voltage_levels) > 1 else '_trend'
            plt.savefig(save_path.replace('.png', f'{suffix}.png'))
            plt.close('电价趋势')
        else:
            plt.show()
    
def main():
    # 使用示例
    reader = ElectricityReader()
    
    # 设置要查询的参数
    year = 2023
    month = 1
    voltage_level = 10.0  # 10kV电压等级
    
    # 1. 获取某月电价数据并显示
    data = reader.get_monthly_electricity_prices(year, month, voltage_level)
    
    if data:
        print("\n电价单位:")
        print(data['电价单位'])
        print("\n电价数值:")
        print(data['电价数值'])
        print("\n电价区间:")
        print(data['电价区间'])
        
        # 绘制并保存图表
        reader.plot_monthly_prices(year, month, voltage_level)
    
    # 2. 分析电价趋势
    print("\n分析电价趋势:")
    # 获取尖峰电价的变化趋势
    trend_data = reader.get_price_time_series(voltage_level, '尖峰')
    if not trend_data.empty:
        print("\n尖峰电价变化趋势:")
        print(trend_data)
        
        # 绘制趋势图
        reader.plot_price_trend(voltage_level, '尖峰')
        
        # 可以连续查看其他电价类型的趋势
        reader.plot_price_trend(voltage_level, '平段')
        
        # 可以指定保存路径
        # reader.plot_price_trend(voltage_level, '尖峰', "电价趋势图_尖峰.png")
    
    # 多电压等级平段电价趋势对比
    print("\n对比不同电压等级的平段电价趋势...")
    reader.plot_price_trend([10.0, 35.0, 110.0, 220.0], "平段")
    
    print("\n分析完成！")

if __name__ == "__main__":
    main()
