import os
import pandas as pd
from typing import Dict, List, Optional

class EFileParser:
    """
    Efile格式文件解析器，用于读取特定格式的文件并转换为DataFrame
    支持通过配置文件定义文件格式
    """
    
    def __init__(self, format_file: str = 'Eformat.properties'):
        """
        初始化解析器
        
        参数:
            format_file (str): 格式配置文件的路径，默认为'Eformat.properties'
        """
        self.current_dir = os.path.dirname(os.path.abspath(__file__))
        self.format_config = self._load_format_config(format_file)
        
    def _load_format_config(self, format_file: str) -> Dict[str, str]:
        """
        加载格式配置文件
        
        参数:
            format_file (str): 配置文件路径
        
        返回:
            Dict[str, str]: 配置信息字典
        """
        config = {}
        config_path = os.path.join(self.current_dir, format_file)
        
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                for line in f:
                    # 移除行首尾的空白字符
                    line = line.strip()
                    
                    # 跳过空行
                    if not line:
                        continue
                        
                    # 处理注释
                    # 如果#不在行首，说明它可能是值的一部分
                    if line.startswith('#'):
                        continue
                    
                    # 查找第一个未转义的#号位置
                    comment_pos = -1
                    i = 0
                    while i < len(line):
                        if line[i] == '\\' and i + 1 < len(line):
                            i += 2  # 跳过转义字符和被转义的字符
                            continue
                        if line[i] == '#':
                            comment_pos = i
                            break
                        i += 1
                    
                    # 如果找到了未转义的注释符号，只取注释符号之前的部分
                    if comment_pos != -1:
                        line = line[:comment_pos].strip()
                    
                    # 解析配置项
                    if '=' in line:
                        key, value = line.split('=', 1)
                        key = key.strip()
                        value = value.strip()
                        
                        # 特殊处理反斜杠分隔符
                        if value == '\\':
                            value = ' '  # 如果值就是一个反斜杠，则使用空格作为分隔符
                        elif '\\' in value:
                            # 处理转义字符
                            i = 0
                            result = []
                            while i < len(value):
                                if value[i] == '\\' and i + 1 < len(value):
                                    next_char = value[i + 1]
                                    if next_char == '#':
                                        result.append('#')  # 保留转义后的#号
                                    elif next_char == 'n':
                                        result.append('\n')
                                    elif next_char == 't':
                                        result.append('\t')
                                    elif next_char == '\\':
                                        result.append('\\')
                                    else:
                                        result.append(next_char)
                                    i += 2
                                else:
                                    result.append(value[i])
                                    i += 1
                            value = ''.join(result)
                        
                        config[key] = value
            return config
        except Exception as e:
            print(f"加载配置文件时发生错误: {str(e)}")
            return {}
            
    def _find_sections(self, content: str) -> List[Dict[str, int]]:
        """
        在文件内容中查找所有数据段
        
        参数:
            content (str): 文件内容
        
        返回:
            List[Dict[str, int]]: 包含每个数据段名称和位置的列表
        """
        sections = []
        lines = content.split('\n')
        current_section = None
        
        for i, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue
                
            # 查找段落开始标记
            if line.startswith('<') and not line.startswith('</'):
                current_section = {
                    'name': line[1:-1],  # 去除 < >
                    'start': i
                }
            # 查找段落结束标记
            elif line.startswith('</') and current_section:
                if line[2:-1] == current_section['name']:
                    current_section['end'] = i
                    sections.append(current_section)
                    current_section = None
                    
        return sections

    def _parse_section(self, lines: List[str], start: int, end: int) -> pd.DataFrame:
        """
        解析数据段内容为DataFrame
        
        参数:
            lines (List[str]): 文件内容行列表
            start (int): 段落开始行
            end (int): 段落结束行
        
        返回:
            pd.DataFrame: 解析后的数据表
        """
        data = []
        columns = None
        
        for line in lines[start + 1:end]:
            line = line.strip()
            if not line or line.startswith('//'):
                continue
                
            if line.startswith(self.format_config['AttributeNameStarter']):
                # 处理列名行
                line = line[1:].strip()  # 移除@符号和前后空格
                columns = [col.strip() for col in line.split(self.format_config['AttributeBreaker'])]
            elif line.startswith(self.format_config['DataLineStarter']):
                # 处理数据行
                line = line[1:].strip()  # 移除#符号和前后空格
                values = [val.strip() for val in line.split(self.format_config['DataBreaker'])]
                data.append(values)
        
        if columns and data:
            df = pd.DataFrame(data, columns=columns)
            # 将数值列转换为浮点数
            for col in df.columns:
                try:
                    df[col] = pd.to_numeric(df[col])
                except:
                    pass
            return df
        return pd.DataFrame()

    def read_file(self, file_path: str) -> Dict[str, pd.DataFrame]:
        """
        读取Efile格式文件并解析为DataFrame字典
        
        参数:
            file_path (str): 要读取的文件路径
        
        返回:
            Dict[str, pd.DataFrame]: 表名到DataFrame的映射字典
        """
        try:
            print(f"正在读取文件: {file_path}")
            print(f"当前使用的配置: {self.format_config}")
            
            # 读取文件内容
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 查找所有数据段
            sections = self._find_sections(content)
            print(f"找到的数据段: {[section['name'] for section in sections]}")
            
            lines = content.split('\n')
            
            # 解析每个数据段
            result = {}
            for section in sections:
                print(f"\n正在解析段落: {section['name']}")
                df = self._parse_section(lines, section['start'], section['end'])
                if not df.empty:
                    result[section['name']] = df
                    print(f"成功解析段落: {section['name']}, 数据形状: {df.shape}")
                else:
                    print(f"警告: 段落 {section['name']} 解析为空DataFrame")
            
            return result
            
        except Exception as e:
            print(f"读取文件时发生错误: {str(e)}")
            return {}

def main():
    # 使用示例
    parser = EFileParser()
    
    # 指定要读取的文件路径
    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "电价数据.Qs")
    
    # 读取并解析文件
    tables = parser.read_file(file_path)
    
    # 打印每个表的信息
    for table_name, df in tables.items():
        print(f"\n{table_name}表的内容:")
        print(df)

if __name__ == "__main__":
    main()
