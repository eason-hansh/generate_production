#!/usr/bin/env python3
"""
生产任务单生成系统启动脚本
"""

import uvicorn
import os
import sys
from pathlib import Path

def main():
    """主函数"""
    # 确保必要的目录存在
    directories = ['uploads', 'outputs', 'templates', 'static']
    for dir_name in directories:
        Path(dir_name).mkdir(exist_ok=True)
    
    # 检查依赖
    try:
        import fastapi
        import openpyxl
        import jinja2
    except ImportError as e:
        print(f"缺少依赖: {e}")
        print("请运行: pip install -r requirements.txt")
        sys.exit(1)
    
    print("🚀 启动生产任务单生成系统...")
    print("📱 访问地址: http://localhost:8000")
    print("⏹️  按 Ctrl+C 停止服务")
    print("-" * 50)
    
    # 启动服务器
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="info"
    )

if __name__ == "__main__":
    main() 