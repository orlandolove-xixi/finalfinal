import os
import pandas as pd
from flask import Flask, request, render_template, redirect, url_for, send_from_directory, flash
from werkzeug.utils import secure_filename
import tempfile
from data_processor import process_excel_data

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key'
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 限制上传文件大小为10MB

# 打印上传目录信息，用于调试
print(f"上传文件将保存在: {app.config['UPLOAD_FOLDER']}")

# 检查文件扩展名是否允许
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# 处理Excel数据的核心函数 - 调用优化后的process_excel_data函数
def process_excel(file_path):
    try:
        # 创建临时输出文件路径
        temp_output = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_{os.path.basename(file_path)}")
        
        # 调用优化后的process_excel_data函数
        process_excel_data(file_path, temp_output)
        
        # 读取处理后的结果
        result_df = pd.read_excel(temp_output)
        
        # 尝试删除临时文件
        try:
            if os.path.exists(temp_output):
                os.remove(temp_output)
        except:
            pass
        
        return result_df
    except Exception as e:
        print(f"处理Excel文件时出错: {e}")
        raise

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        try:
            # 检查是否有文件部分
            if 'file' not in request.files:
                flash('没有文件部分，请确保表单正确提交')
                print('错误: 请求中没有文件部分')
                return redirect(request.url)
            
            file = request.files['file']
            
            # 如果用户没有选择文件
            if file.filename == '':
                flash('没有选择文件，请选择一个Excel文件上传')
                print('错误: 未选择文件')
                return redirect(request.url)
            
            # 检查文件扩展名
            if not allowed_file(file.filename):
                flash('不支持的文件格式，请上传.xlsx或.xls文件')
                print(f'错误: 文件格式不支持: {file.filename}')
                return redirect(request.url)
            
            # 文件有效，继续处理
            original_filename = file.filename  # 保存原始文件名
            filename = secure_filename(original_filename)  # 用于保存的安全文件名
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            print(f'原始文件名: {original_filename}, 保存文件名: {filename}')
            
            # 保存文件
            try:
                file.save(filepath)
                print(f'成功上传文件: {filepath}')
            except Exception as save_error:
                flash(f"保存文件时出错: {str(save_error)}")
                print(f'错误: 保存文件失败: {save_error}')
                return redirect(request.url)
            
            # 检查文件是否成功保存
            if not os.path.exists(filepath):
                flash('文件上传失败，请重试')
                print(f'错误: 文件保存后不存在: {filepath}')
                return redirect(request.url)
            
            # 处理Excel文件
            try:
                print(f'开始处理文件: {filepath}')
                result_df = process_excel(filepath)
                
                # 保存处理后的文件，使用原始文件名的信息
                name_without_ext = os.path.splitext(original_filename)[0]
                ext = os.path.splitext(original_filename)[1]
                output_filename = f"processed_{name_without_ext}{ext}"
                output_secure_filename = secure_filename(output_filename)  # 安全处理输出文件名
                output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_secure_filename)
                
                result_df.to_excel(output_filepath, index=True)
                print(f'成功处理并保存结果文件: {output_filepath}')
                print(f'原始文件名: {original_filename}, 下载文件名: {output_filename}, 保存文件名: {output_secure_filename}')
                
                # 将DataFrame转换为HTML表格，用于预览
                html_table = result_df.to_html(classes='table table-striped table-hover', index=True, na_rep='-')
                
                # 获取处理后的文件大小
                file_size = os.path.getsize(output_filepath) / 1024  # KB
                
                # 获取处理前后的行数和列数
                original_df = pd.read_excel(filepath)
                original_rows, original_cols = original_df.shape
                processed_rows, processed_cols = result_df.shape
                
                # 传递处理结果到success页面
                return render_template('success.html', 
                                      original_filename=original_filename,  # 使用原始文件名
                                      output_filename=output_secure_filename,  # 使用安全保存的文件名用于下载
                                      display_filename=output_filename,  # 用于在页面上显示的文件名
                                      html_table=html_table,
                                      file_size=round(file_size, 2),
                                      original_rows=original_rows,
                                      original_cols=original_cols,
                                      processed_rows=processed_rows,
                                      processed_cols=processed_cols)
                
            except pd.errors.EmptyDataError:
                flash('Excel文件为空，请上传有效的Excel文件')
                print('错误: Excel文件为空')
                return redirect(request.url)
            except pd.errors.ParserError:
                flash('解析Excel文件时出错，文件格式可能不正确')
                print('错误: Excel文件解析失败')
                return redirect(request.url)
            except Exception as e:
                flash(f"处理文件时出错: {str(e)}")
                print(f'错误: 处理Excel文件时出错: {e}')
                return redirect(request.url)
        except request.exceptions.RequestEntityTooLarge:
            flash('文件太大，请上传小于10MB的文件')
            print('错误: 文件大小超过10MB限制')
            return redirect(request.url)
        except Exception as general_error:
            flash(f"上传过程中发生错误: {str(general_error)}")
            print(f'错误: 上传过程中的一般错误: {general_error}')
            return redirect(request.url)
    
    # 渲染上传页面
    return render_template('index.html')

@app.route('/test_upload', methods=['GET'])
def test_upload_page():
    return "后端服务运行正常，文件上传功能已修复。请返回首页测试上传功能。"

@app.route('/download/<filename>')
def download_file(filename):
    try:
        # 安全检查文件名，防止目录遍历攻击
        safe_filename = secure_filename(filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_filename)
        
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f'错误: 下载文件不存在: {file_path}')
            flash('下载文件不存在或已被删除')
            return redirect(url_for('index'))
        
        # 从URL参数获取原始文件名，如果有的话
        display_name = request.args.get('display_name', safe_filename)
        print(f'下载文件: {file_path}, 显示名称: {display_name}')
        
        return send_from_directory(app.config['UPLOAD_FOLDER'], safe_filename, 
                                  as_attachment=True, 
                                  download_name=display_name)
    except Exception as e:
        print(f'下载文件时出错: {str(e)}')
        flash(f'下载文件时出错: {str(e)}')
        return redirect(url_for('index'))

@app.route('/preview')
def preview_data():
    # 预览页面，可能需要根据实际需求调整
    return render_template('preview.html')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)