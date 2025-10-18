import pandas as pd
import numpy as np

def create_mock_hainan_data():
    """
    根据用户提供的海南省.xls文件结构创建模拟数据
    """
    # 创建第一部分数据（第1-30行）
    data1 = [
        ['海南省', '', '', '', '', '', '', ''],  # 第1行
        ['指标', '单位', '西沙群岛', '南沙群岛', '中沙群岛办事处', '五指山市', '琼海市'],  # 第2行
        ['一、基本情况', '', '', '', '', '', ''],  # 第3行
        ['行政区划面积', '平方公里', '', '', '', 1143, 1710],  # 第4行
        ['乡', '个', '', '', '', '', ''],  # 第5行
        ['镇', '个', '', '', '', 4, 12],  # 第6行
        ['街道办事处', '个', '', '', '', '', ''],  # 第7行
        ['二、综合经济', '', '', '', '', '', ''],  # 第8行
        ['户籍人口', '万人', '', '', '', 10, 52],  # 第9行
        ['地区生产总值', '万元', '', '', '', 385102, 3569448],  # 第10行
        ['第一产业增加值', '万元', '', '', '', 63859, 1257239],  # 第11行
        ['第二产业增加值', '万元', '', '', '', 83452, 466050],  # 第12行
        ['第三产业增加值', '万元', '', '', '', 237791, 1846159],  # 第13行
        ['地方一般公共预算收入', '万元', '', '', '', 28098, 106301],  # 第14行
        ['地方一般公共预算支出', '万元', '', '', '', 660388, 542169],  # 第15行
        ['住户存款余额', '万元', '', '', '', 697600, 3016432],  # 第16行
        ['三、农业、工业和交通', '', '', '', '', '', ''],  # 第17行
        ['设施农业种植占地面积', '公顷', '', '', '', 98, 222],  # 第18行
        ['油料产量', '吨', '', '', '', 373, 246],  # 第19行
        ['热带水果产量', '吨', '', '', '', '', ''],  # 第20行
        ['规模以上工业企业', '个', '', '', '', 5, 13],  # 第21行
        ['固定电话用户', '户', '', '', '', 11017, 97321],  # 第22行
        ['四、教育、卫生和社会保障', '', '', '', '', '', ''],  # 第23行
        ['普通中学在校学生', '人', '', '', '', 9159, 34301],  # 第24行
        ['小学在校学生', '人', '', '', '', 8168, 23606],  # 第25行
        ['提供住宿的民政服务机构', '个', '', '', '', 4, 17],  # 第26行
        ['提供住宿的民政服务机构床位', '张', '', '', '', 490, 1976],  # 第27行
        ['提供住宿的民政服务机构床位', '张', '', '', '', 246, 739],  # 第28行
        ['', '', '', '', '', '', ''],  # 第29行
        ['', '', '', '', '', '', '']  # 第30行
    ]
    
    # 创建第二部分数据（第31-46行）
    data2 = [
        ['', '', '', '', '', '', '', ''],  # 第31行
        ['', '', '', '', '', '', '', ''],  # 第32行
        ['海南省', '', '', '', '', '', '', ''],  # 第33行
        ['指标', '单位', '文昌市', '万宁市', '东方市', '定安县', '屯昌县'],  # 第34行
        ['一、基本情况', '', '', '', '', '', ''],  # 第35行
        ['行政区划面积', '平方公里', 2459, 1905, 2213, 1197, 1224],  # 第36行
        ['乡', '个', 17, 12, 2, 10, 8],  # 第37行
        ['镇', '个', '', '', '', '', ''],  # 第38行
        ['户籍人口', '万人', 60, 63, 47, 33, 31],  # 第39行
        ['二、综合经济', '', '', '', '', '', ''],  # 第40行
        ['地区生产总值', '万元', 3439585, 2875001, 2213401, 1228646, 1013319],  # 第41行
        ['第一产业增加值', '万元', 1243751, 954200, 835870, 467931, 366761],  # 第42行
        ['第二产业增加值', '万元', 781455, 724900, 446239, 185801, 195551],  # 第43行
        ['第三产业增加值', '万元', 1414379, 1195901, 931292, 574914, 450987],  # 第44行
        ['地方一般公共预算收入', '万元', 159999, 122700, 93507, 46674, 29482],  # 第45行
        ['地方一般公共预算支出', '万元', 660933, 566589, 495246, 380000, 261203]  # 第46行
    ]
    
    # 合并数据
    all_data = data1 + data2
    
    # 创建DataFrame
    df = pd.DataFrame(all_data, columns=[f'列{i}' for i in range(1, len(all_data[0])+1)])
    
    print("\n=== 创建模拟海南省数据完成 ===")
    print(f"数据形状: {df.shape}")
    print("\n前10行数据:")
    print(df.head(10))
    print("\n33-38行数据 (第二个表格):")
    print(df.iloc[32:38])  # Python索引从0开始，所以第33行是索引32
    
    return df

def process_mock_data(df):
    """
    使用我们的算法处理模拟数据，测试表格识别和县域提取功能
    """
    print("\n=== 开始处理模拟数据 ===")
    
    # 存储所有数据的字典
    all_county_data = {}
    all_metrics = set()
    # 保存县域名称的原始顺序
    original_county_order = []
    
    # 首先找到所有表格的起始行
    table_starts = []
    print("\n扫描表格起始行:")
    for i in range(len(df)):
        # 打印每行第一列的值用于调试
        first_col = df.iloc[i, 0] if 0 < len(df.columns) else "N/A"
        second_col = df.iloc[i, 1] if 1 < len(df.columns) else "N/A"
        print(f"  行{i+1}: 第一列='{first_col}', 第二列='{second_col}'")
        
        # 检查是否为表格起始行
        if pd.notna(first_col) and str(first_col).strip() == '指标':
            # 检查第二列是否为'单位'
            if pd.notna(second_col) and str(second_col).strip() == '单位':
                table_starts.append(i)
                print(f"  ✅ 发现表格起始行: 第{i+1}行")
    
    print(f"\n总共发现 {len(table_starts)} 个表格区域: {[s+1 for s in table_starts]}")
    
    # 处理每个表格区域
    for start_row in table_starts:
        print(f"\n=== 处理表格区域，起始行: {start_row + 1} ===")
        
        # 打印表格头信息
        header_row = df.iloc[start_row]
        print(f"  表格头信息: {header_row.tolist()}")
        
        # 提取县域名称（从第三列开始的表头）
        county_names = []
        print(f"  提取县域名称 (从第{start_row+1}行的第三列开始):")
        for j in range(2, len(df.columns)):  # 从第三列开始
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
                # 跳过分类行（如"一、基本情况"等）
                if metric and not any(metric.startswith(prefix) for prefix in ['一、', '二、', '三、', '四、', '五、', '六、']):
                    all_metrics.add(metric)
                    # 读取每个县域的数据
                    for j, county_name in enumerate(county_names):
                        data_col = j + 2  # 数据从第三列开始
                        if data_col < len(df.columns) and pd.notna(df.iloc[i, data_col]):
                            all_county_data[county_name][metric] = df.iloc[i, data_col]
                processed_rows += 1
            i += 1
        print(f"  表格区域处理完成，共处理 {processed_rows} 行数据")
    
    print(f"\n=== 总统计 ===")
    print(f"- 识别到 {len(all_county_data)} 个县域")
    print(f"- 识别到 {len(all_metrics)} 个指标")
    print(f"- 县域列表: {original_county_order}")
    
    # 验证是否所有县域都被正确识别
    expected_counties = ['西沙群岛', '南沙群岛', '中沙群岛办事处', '五指山市', '琼海市', '文昌市', '万宁市', '东方市', '定安县', '屯昌县']
    print(f"\n=== 验证结果 ===")
    print(f"预期县域数量: {len(expected_counties)}")
    print(f"实际识别县域数量: {len(all_county_data)}")
    print(f"\n未识别的县域: {[county for county in expected_counties if county not in all_county_data]}")
    print(f"额外识别的县域: {[county for county in all_county_data if county not in expected_counties]}")
    
    # 检查算法是否正确识别了两个表格
    if len(table_starts) == 2:
        print("\n✅ 成功识别了两个表格区域！")
    else:
        print(f"\n❌ 未正确识别两个表格区域，实际识别: {len(table_starts)}个")
    
    # 检查是否所有县域都被识别
    if set(expected_counties).issubset(set(all_county_data.keys())):
        print("✅ 所有预期县域都被正确识别！")
    else:
        print("❌ 部分县域未被识别！")
    
    return all_county_data, all_metrics, original_county_order

def main():
    print("=== 海南省表格识别算法测试 ===")
    
    # 创建模拟数据
    df = create_mock_hainan_data()
    
    # 处理模拟数据
    all_county_data, all_metrics, original_county_order = process_mock_data(df)
    
    # 生成一个简单的结果文件来验证
    if all_county_data and all_metrics:
        metrics_list = sorted(list(all_metrics))
        
        # 创建结果数据
        result_data = []
        result_data.append(['指标名称'] + original_county_order)
        
        for metric in metrics_list:
            row = [metric]
            for county in original_county_order:
                if metric in all_county_data[county]:
                    row.append(all_county_data[county][metric])
                else:
                    row.append('')
            result_data.append(row)
        
        # 创建结果DataFrame
        result_df = pd.DataFrame(result_data)
        
        # 保存结果
        result_file = "c:\\Users\\L\\Documents\\trae_projects\\data2\\mock_hainan_result.xlsx"
        result_df.to_excel(result_file, index=False, header=False)
        print(f"\n✅ 模拟结果已保存到: {result_file}")
        print("\n结果预览:")
        print(result_df.head(10))
    
    print("\n=== 测试完成 ===")

if __name__ == "__main__":
    main()