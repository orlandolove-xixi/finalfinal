import sys
import pandas as pd
import os

def process_hainan_data(input_file, output_file):
    """
    直接处理海南省Excel文件，添加详细调试日志
    """
    try:
        print(f"\n=== 开始处理文件: {input_file} ===")
        
        # 检查文件是否存在
        if not os.path.exists(input_file):
            print(f"错误: 文件不存在: {input_file}")
            return False
        
        # 尝试不同的引擎读取Excel文件
        try:
            # 首先尝试默认引擎
            excel_file = pd.ExcelFile(input_file)
        except Exception as e:
            print(f"默认引擎读取失败: {e}")
            # 尝试使用openpyxl引擎
            try:
                excel_file = pd.ExcelFile(input_file, engine='openpyxl')
                print("使用openpyxl引擎成功读取")
            except Exception as e2:
                print(f"openpyxl引擎读取失败: {e2}")
                return False
        
        print(f"找到 {len(excel_file.sheet_names)} 个工作表: {excel_file.sheet_names}")
        
        # 存储所有数据的字典
        all_county_data = {}
        all_metrics = set()
        # 保存县域名称的原始顺序
        original_county_order = []
        
        for sheet_name in excel_file.sheet_names:
            print(f"\n--- 处理工作表: {sheet_name} ---")
            # 读取当前工作表
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            print(f"工作表 '{sheet_name}' 形状: {df.shape}")
            
            # 打印前5行数据样例
            print("\n工作表前5行数据:")
            print(df.head())
            
            # 首先找到所有表格的起始行
            table_starts = []
            print("\n扫描表格起始行:")
            for i in range(len(df)):
                # 打印每行A列的值用于调试
                if i < 50:  # 只打印前50行以避免过多输出
                    a_cell = df.iloc[i, 0] if i < len(df) and 0 < len(df.columns) else "N/A"
                    b_cell = df.iloc[i, 1] if i < len(df) and 1 < len(df.columns) else "N/A"
                    print(f"  行{i+1}: A={a_cell}, B={b_cell}")
                
                if len(df.columns) > 0 and pd.notna(df.iloc[i, 0]):
                    cell_value = str(df.iloc[i, 0]).strip()
                    if cell_value == '指标':
                        # 检查B列是否为'单位'，确保是有效的表格头
                        if len(df.columns) > 1 and pd.notna(df.iloc[i, 1]) and str(df.iloc[i, 1]).strip() == '单位':
                            table_starts.append(i)
                            print(f"  ✅ 发现表格起始行: 第{i+1}行")
            
            print(f"\n工作表 '{sheet_name}' 中发现 {len(table_starts)} 个表格区域: {[s+1 for s in table_starts]}")
            
            # 处理每个表格区域
            for start_row in table_starts:
                print(f"\n  === 处理表格区域，起始行: {start_row + 1} ===")
                
                # 打印表格头信息
                header_row = df.iloc[start_row]
                print(f"  表格头信息: {header_row.tolist()}")
                
                # 提取县域名称（从C列开始的表头）
                county_names = []
                print(f"  提取县域名称 (从第{start_row+1}行的C列开始):")
                for j in range(2, len(df.columns)):
                    if pd.notna(df.iloc[start_row, j]) and str(df.iloc[start_row, j]).strip():
                        county_name = str(df.iloc[start_row, j]).strip()
                        county_names.append(county_name)
                        # 初始化该县域的数据字典
                        if county_name not in all_county_data:
                            all_county_data[county_name] = {}
                            # 保留原始顺序
                            if county_name not in original_county_order:
                                original_county_order.append(county_name)
                            print(f"    ✅ 添加新县域: '{county_name}'")
                    else:
                        col_value = df.iloc[start_row, j] if j < len(df.columns) else "N/A"
                        print(f"    列{j+1}: {col_value} (跳过)")
                
                print(f"  从该表格区域识别到 {len(county_names)} 个县域: {county_names}")
                
                # 处理当前表格的数据行
                i = start_row + 1  # 从表头的下一行开始
                # 找到下一个表格的起始位置作为当前表格的结束
                next_table_start = None
                for next_start in table_starts:
                    if next_start > start_row:
                        next_table_start = next_start
                        break
                
                # 确定当前表格的结束行
                end_row = next_table_start if next_table_start is not None else len(df)
                print(f"  当前表格处理范围: 行 {i+1} 到行 {end_row}")
                
                # 读取指标数据，直到遇到下一个表格开始或工作表结束
                processed_rows = 0
                while i < end_row:
                    # 确保当前行有数据再处理
                    if pd.notna(df.iloc[i, 0]):
                        cell_value = str(df.iloc[i, 0]).strip()
                        # 指标处理
                        metric = cell_value
                        # 打印处理的行信息用于调试
                        if processed_rows < 10:  # 只打印前10行处理的详情
                            print(f"    处理行 {i+1}: 指标='{metric}'")
                        # 跳过分类行（如"一、基本情况"等）
                        if metric and not any(metric.startswith(prefix) for prefix in ['一、', '二、', '三、', '四、', '五、', '六、']):
                            all_metrics.add(metric)
                            # 读取每个县域的数据
                            for j, county_name in enumerate(county_names):
                                data_col = j + 2  # 数据从C列开始
                                if data_col < len(df.columns) and pd.notna(df.iloc[i, data_col]):
                                    all_county_data[county_name][metric] = df.iloc[i, data_col]
                                    if processed_rows < 3:  # 只打印前3行的数据详情
                                        print(f"      {county_name}: {df.iloc[i, data_col]}")
                        processed_rows += 1
                    i += 1
                print(f"  表格区域处理完成，共处理 {processed_rows} 行数据")
        
        print(f"\n=== 总统计 ===")
        print(f"- 识别到 {len(all_county_data)} 个县域")
        print(f"- 识别到 {len(all_metrics)} 个指标")
        print(f"- 县域列表: {original_county_order}")
        
        # 如果有数据，创建输出文件
        if all_county_data and all_metrics:
            metrics_list = sorted(list(all_metrics))
            
            # 创建新的数据框
            result_data = []
            
            # 添加县域名称列
            result_data.append(['指标名称'] + original_county_order)
            
            # 添加每个指标的数据
            for metric in metrics_list:
                row = [metric]
                for county in original_county_order:
                    if metric in all_county_data[county]:
                        row.append(all_county_data[county][metric])
                    else:
                        row.append('')  # 空值处理
                result_data.append(row)
            
            # 创建DataFrame
            result_df = pd.DataFrame(result_data)
            
            # 保存到Excel
            try:
                result_df.to_excel(output_file, index=False, header=False)
                print(f"\n✅ 处理完成！结果已保存到: {output_file}")
                return True
            except Exception as e:
                print(f"保存文件失败: {e}")
                return False
        else:
            print("\n❌ 未识别到有效数据")
            return False
            
    except Exception as e:
        print(f"处理过程中出错: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    # 处理非加密版本的Excel文件
    input_file = "F:\\桌面\\工作簿1.xlsx"
    output_file = "c:\\Users\\L\\Documents\\trae_projects\\data2\\result.xlsx"
    
    print(f"=== 海南省表格处理测试 ===")
    print(f"输入文件: {input_file}")
    print(f"输出文件: {output_file}")
    
    success = process_hainan_data(input_file, output_file)
    
    if success:
        print("\n测试成功完成！")
    else:
        print("\n测试失败，请检查错误信息。")