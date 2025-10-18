@echo off

REM 设置Python环境变量（如果需要）
REM set PATH=C:\Python39\;C:\Python39\Scripts\;%PATH%

REM 创建虚拟环境（如果不存在）
if not exist "venv" (
    echo 创建虚拟环境...
    python -m venv venv
)

REM 激活虚拟环境
echo 激活虚拟环境...
call venv\Scripts\activate

REM 升级pip
echo 升级pip...
python -m pip install --upgrade pip

REM 安装依赖
echo 安装依赖包...
pip install -r requirements.txt

REM 启动应用
echo 启动Excel数据处理工具...
python main.py

pause