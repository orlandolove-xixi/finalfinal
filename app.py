from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
import os
import pandas as pd
import numpy as np
from werkzeug.utils import secure_filename
import tempfile

app = Flask(__name__)
app.config['SECRET_KEY'] = 'supersecretkey'
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}

def allowed_file(filename):
    """检查文件是否为允许的类型"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def process_excel_data(input_file, output_file):
    """
    处理Excel数据，将多个表格整合为一个标准格式，确保县域名称在A列，指标在第一行
    支持识别工作表中的所有表格区域，包括后续表格中的县域名称
    
    参数:
    input_file: 输入Excel文件路径
    output_file: 输出Excel文件路径
    """
    try:
        print(f"开始处理文件: {input_file}")
        
        # 尝试读取Excel文件（支持xlsx和xls格式）
        try:
            excel_file = pd.ExcelFile(input_file)
            print(f"成功读取Excel文件，找到 {len(excel_file.sheet_names)} 个工作表")
        except Exception as e:
            print(f"读取Excel文件失败: {e}")
            raise
        
        # 存储所有数据的字典
        all_county_data = {}
        all_metrics = set()
        # 保存县域名称的原始顺序
        original_county_order = []
        total_tables = 0
        
        # 处理每个工作表
        for sheet_name in excel_file.sheet_names:
            print(f"\n处理工作表: {sheet_name}")
            # 读取当前工作表
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            print(f"工作表 '{sheet_name}' 形状: {df.shape}")
            
            # 首先找到所有表格的起始行
            table_starts = []
            for i in range(len(df)):
                # 安全检查列是否存在
                if len(df.columns) > 0 and pd.notna(df.iloc[i, 0]):
                    cell_value = str(df.iloc[i, 0]).strip()
                    # 检查是否为表格标题行
                    if cell_value == '指标':
                        # 检查B列是否为'单位'，确保是有效的表格头
                        if len(df.columns) > 1 and pd.notna(df.iloc[i, 1]) and str(df.iloc[i, 1]).strip() == '单位':
                            table_starts.append(i)
                            print(f"  发现表格起始行: 第{i+1}行")
            
            total_tables += len(table_starts)
            print(f"工作表 '{sheet_name}' 中发现 {len(table_starts)} 个表格区域")
            
            # 处理每个表格区域
            for start_row in table_starts:
                print(f"  处理表格区域，起始行: {start_row + 1}")
                
                # 提取县域名称（从C列开始的表头）
                county_names = []
                for j in range(2, len(df.columns)):  # 从第三列开始
                    if pd.notna(df.iloc[start_row, j]):
                        # 保留原始格式（包括可能的换行符）
                        county_name = str(df.iloc[start_row, j]).strip()
                        if county_name:  # 确保不为空字符串
                            county_names.append(county_name)
                            # 初始化该县域的数据字典
                            if county_name not in all_county_data:
                                all_county_data[county_name] = {}
                                # 保留原始顺序
                                if county_name not in original_county_order:
                                    original_county_order.append(county_name)
                                print(f"    添加新县域: '{county_name}'")
                
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
                    if len(df.columns) > 0 and pd.notna(df.iloc[i, 0]):
                        cell_value = str(df.iloc[i, 0]).strip()
                        # 指标处理
                        metric = cell_value
                        # 跳过分类行（如"一、基本情况"等）
                        if metric and not any(metric.startswith(prefix) for prefix in ['一、', '二、', '三、', '四、', '五、', '六、', '七、', '八、']):
                            all_metrics.add(metric)
                            # 读取每个县域的数据
                            for j, county_name in enumerate(county_names):
                                data_col = j + 2  # 数据从第三列开始
                                if data_col < len(df.columns) and pd.notna(df.iloc[i, data_col]):
                                    # 确保数据类型正确
                                    data_value = df.iloc[i, data_col]
                                    all_county_data[county_name][metric] = data_value
                        processed_rows += 1
                    i += 1
                print(f"  表格区域处理完成，共处理 {processed_rows} 行数据")
        
        # 输出总统计信息
        print(f"\n=== 总统计 ===")
        print(f"- 总共处理 {total_tables} 个表格区域")
        print(f"- 识别到 {len(all_county_data)} 个县域")
        print(f"- 识别到 {len(all_metrics)} 个指标")
        print(f"- 县域列表: {original_county_order}")
        
        # 创建新的数据框，注意转置结构
        metrics_list = sorted(list(all_metrics))
        
        # 创建数据，让县域作为行，指标作为列
        data = []
        for county in original_county_order:  # 使用原始顺序
            row_data = [county]  # 县域名称作为第一列数据
            for metric in metrics_list:
                row_data.append(all_county_data[county].get(metric, np.nan))
            data.append(row_data)
        
        # 创建列名：第一列为'县域'，其余为指标名
        columns = ['县域'] + metrics_list
        
        # 创建DataFrame
        result_df = pd.DataFrame(data, columns=columns)
        
        # 保存结果，不包含默认索引
        result_df.to_excel(output_file, index=False)
        
        return len(all_county_data), len(metrics_list)
    
    except Exception as e:
        raise Exception(f"处理过程中出现错误: {str(e)}")

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # 检查是否有文件上传
        if 'file' not in request.files:
            flash('请选择要上传的文件')
            return redirect(request.url)
        
        file = request.files['file']
        
        # 检查文件是否为空
        if file.filename == '':
            flash('请选择要上传的文件')
            return redirect(request.url)
        
        # 检查文件类型
        if file and allowed_file(file.filename):
            # 保存上传的文件
            filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(input_path)
            
            # 处理文件
            output_filename = f"处理结果_{filename}"
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
            
            try:
                county_count, metric_count = process_excel_data(input_path, output_path)
                flash(f'数据处理成功！共处理 {county_count} 个县域，{metric_count} 个指标')
                return redirect(url_for('download_file', filename=output_filename))
            except Exception as e:
                flash(str(e))
                return redirect(request.url)
        else:
            flash('只支持 .xlsx 和 .xls 格式的文件')
            return redirect(request.url)
    
    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

@app.route('/preview/<filename>')
def preview_file(filename):
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        df = pd.read_excel(file_path)
        
        # 获取前10行数据用于预览
        preview_data = df.head(10).to_html(classes='table table-striped table-hover', index=True)
        
        # 获取基本统计信息
        stats = {
            '行数': len(df),
            '列数': len(df.columns),
            '指标数量': len(df),
            '县域列表': list(df.columns)
        }
        
        return render_template('preview.html', preview_data=preview_data, stats=stats, filename=filename)
    except Exception as e:
        flash(f'预览失败: {str(e)}')
        return redirect(url_for('index'))

if __name__ == '__main__':
    # 确保上传文件夹存在
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    
    # 在生产环境中应该修改为更安全的配置
    app.run(debug=True, host='0.0.0.0', port=5000)