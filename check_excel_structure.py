import pandas as pd
import os

# 检查Excel文件结构
def check_excel_structure(file_path):
    print(f"检查文件: {file_path}")
    
    # 读取所有工作表
    excel_file = pd.ExcelFile(file_path)
    print(f"工作表数量: {len(excel_file.sheet_names)}")
    print(f"工作表名称: {excel_file.sheet_names}")
    
    for sheet_name in excel_file.sheet_names:
        print(f"\n检查工作表: {sheet_name}")
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"工作表形状: {df.shape}")
        print(f"列数: {df.shape[1]}")
        print(f"行数: {df.shape[0]}")
        
        # 查找所有包含'指标'的行
        print("\n包含'指标'的行:")
        for i in range(min(100, len(df))):  # 只检查前100行
            if pd.notna(df.iloc[i, 0]):
                cell_value = str(df.iloc[i, 0]).strip()
                if '指标' in cell_value:
                    print(f"第{i+1}行, A列值: '{cell_value}'")
                    # 打印该行的所有列值
                    print(f"该行所有值: {[str(df.iloc[i, j]).strip() if pd.notna(df.iloc[i, j]) else 'NaN' for j in range(min(10, len(df.columns)))]}")
        
        # 查看前20行数据结构
        print("\n前20行数据:")
        print(df.head(20).to_string(max_rows=20, max_cols=10))
        
        # 查看工作表末尾20行
        print("\n末尾20行数据:")
        print(df.tail(20).to_string(max_rows=20, max_cols=10))

if __name__ == "__main__":
    file_path = r"F:\桌面\海南省.xls"
    if os.path.exists(file_path):
        check_excel_structure(file_path)
    else:
        print(f"错误: 文件不存在 - {file_path}")