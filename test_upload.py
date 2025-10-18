import requests
import os

# 确保上传文件夹存在
def test_upload():
    print("开始测试文件上传功能...")
    
    # 创建一个简单的Excel文件用于测试
    import pandas as pd
    import tempfile
    
    # 创建一个测试用的Excel文件
    df = pd.DataFrame({
        '指标': ['地区生产总值', '人口数', '工业增加值'],
        '单位': ['亿元', '万人', '亿元'],
        '成都市': [15000, 1600, 4500],
        '绵阳市': [3500, 500, 1200],
        '德阳市': [2800, 400, 900]
    })
    
    # 保存测试文件
    temp_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    df.to_excel(temp_file.name, index=False)
    temp_file.close()
    
    print(f"创建测试文件: {temp_file.name}")
    
    # 测试上传
    try:
        url = 'http://127.0.0.1:5000/'
        with open(temp_file.name, 'rb') as f:
            files = {'file': (os.path.basename(temp_file.name), f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            print("正在发送上传请求...")
            response = requests.post(url, files=files, allow_redirects=False)
            
            print(f"\n响应状态码: {response.status_code}")
            print(f"响应头: {response.headers}")
            
            # 解码响应内容，处理可能的编码问题
            try:
                print(f"响应内容: {response.content.decode('utf-8', errors='replace')[:1000]}...")
            except:
                print("无法解码响应内容")
                print(f"原始响应内容前1000字节: {response.content[:1000]}")
                
            # 检查是否有重定向
            if response.status_code in [301, 302, 303, 307, 308]:
                print(f"重定向到: {response.headers.get('Location')}")
                
    except Exception as e:
        print(f"上传测试时出错: {e}")
    finally:
        # 清理临时文件
        try:
            os.unlink(temp_file.name)
            print(f"已删除测试文件: {temp_file.name}")
        except:
            pass

if __name__ == "__main__":
    test_upload()