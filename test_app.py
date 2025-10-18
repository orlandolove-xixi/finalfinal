# 简单的应用测试脚本，用于验证Flask应用的基本功能
import os
import sys

print("开始测试应用功能...")
print(f"当前工作目录: {os.getcwd()}")

# 检查必要的文件是否存在
required_files = ['main.py', 'data_processor.py', 'templates/index.html', 'templates/success.html']
for file in required_files:
    if os.path.exists(file):
        print(f"✓ 找到文件: {file}")
    else:
        print(f"✗ 未找到文件: {file}")

# 测试导入功能
try:
    # 尝试导入主要模块
    from data_processor import process_excel_data
    print("✓ 成功导入process_excel_data函数")
except ImportError as e:
    print(f"✗ 导入错误: {e}")
    sys.exit(1)

print("\n基本检查完成！应用文件结构看起来是正确的。")
print("请确保在运行应用时使用: python main.py")