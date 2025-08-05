from fastapi import FastAPI, File, UploadFile, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.requests import Request
import os
import json
import uuid
from datetime import datetime, timedelta
import shutil
from pathlib import Path
from openpyxl import load_workbook
import asyncio
from typing import Optional
import logging

# 导入配置和自定义模块
from config import *
from utils.excel_processor import ExcelProcessor
from utils.pdf_extractor import PDFExtractor

# 配置日志
logging.basicConfig(level=getattr(logging, LOG_LEVEL), format=LOG_FORMAT)
logger = logging.getLogger(__name__)

# 初始化 pdf 抽取类
pdf_extractor = PDFExtractor()
# 初始化 Excel 处理类
excel_processor = ExcelProcessor()

app = FastAPI(
    title="生产任务单生成系统", 
    version="1.0.0",
    debug=DEBUG
)

# 创建必要的目录
for dir_path in [UPLOAD_DIR, TEMPLATE_DIR, OUTPUT_DIR, STATIC_DIR]:
    dir_path.mkdir(exist_ok=True)

# 挂载静态文件
app.mount("/static", StaticFiles(directory="static"), name="static")

# 模板配置
templates = Jinja2Templates(directory="templates")

# 全局变量存储处理状态
processing_status = {}

@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    """主页面"""
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/upload")
async def upload_files(
    background_tasks: BackgroundTasks,
    pdf_file: UploadFile = File(...),
    excel_template: UploadFile = File(...),
    task_order_no: str = Form(...),
    order_date: str = Form(...),
    delivery_date: str = Form(...)
):
    """上传文件并处理"""
    try:
        # 生成唯一任务ID
        task_id = str(uuid.uuid4())
        
        # 为每个任务创建独立的文件夹
        task_upload_dir = UPLOAD_DIR / f"task_{task_id}"
        task_upload_dir.mkdir(exist_ok=True)
        
        # 保存上传的文件到任务文件夹
        pdf_path = task_upload_dir / "pdf_file.pdf"
        excel_path = task_upload_dir / "excel_template.xlsx"
        
        with open(pdf_path, "wb") as buffer:
            shutil.copyfileobj(pdf_file.file, buffer)
        
        with open(excel_path, "wb") as buffer:
            shutil.copyfileobj(excel_template.file, buffer)
        
        # 获取PDF文件名（不含扩展名）
        pdf_name = Path(pdf_file.filename).stem
        
        # 初始化任务状态
        processing_status[task_id] = {
            "status": "processing", 
            "message": "正在处理...",
            "upload_dir": str(task_upload_dir),
            "created_time": datetime.now().isoformat(),
            "can_cleanup": False
        }
        
        # 使用 BackgroundTasks 处理文件
        background_tasks.add_task(
            process_files,
            task_id, pdf_path, excel_path, 
            task_order_no, order_date, delivery_date, pdf_name
        )
        
        logger.info(f"任务 {task_id} 已添加到后台处理队列，上传目录: {task_upload_dir}")
        return {"task_id": task_id, "status": "processing"}
    
    except Exception as e:
        logger.error(f"文件上传失败: {str(e)}")
        raise HTTPException(status_code=500, detail=f"上传失败: {str(e)}")

async def process_files(task_id: str, pdf_path: Path, excel_path: Path, 
                       task_order_no: str, order_date: str, delivery_date: str, pdf_name: str):
    """后台处理文件"""
    # 记录开始时间
    start_time = datetime.now()
    
    try:
        logger.info(f"开始处理任务 {task_id}")
        processing_status[task_id]["message"] = "正在提取PDF信息..."
        
        # 1. 提取PDF信息
        pdf_info = pdf_extractor.process(str(pdf_path))
        logger.info(f"任务 {task_id} PDF信息提取完成")
        
        processing_status[task_id]["message"] = "正在生成生产任务单..."
        
        # 2. 处理Excel模板并保存到输出路径
        saved_file_path = excel_processor.process(pdf_info, excel_path, task_order_no, order_date, delivery_date, str(OUTPUT_DIR), pdf_name)
        
        # 计算处理时长
        end_time = datetime.now()
        processing_duration = (end_time - start_time).total_seconds()
        
        processing_status[task_id].update({
            "status": "completed", 
            "message": f"处理完成 (耗时: {processing_duration:.1f}秒)",
            "output_file": saved_file_path,
            "completed_time": datetime.now().isoformat(),
            "processing_duration": processing_duration
        })
        
        logger.info(f"任务 {task_id} 处理完成，文件保存至: {saved_file_path}，耗时: {processing_duration:.1f}秒")
        
    except Exception as e:
        # 计算处理时长（即使失败也要记录）
        end_time = datetime.now()
        processing_duration = (end_time - start_time).total_seconds()
        
        error_msg = f"处理失败: {str(e)}"
        logger.error(f"任务 {task_id} {error_msg}，耗时: {processing_duration:.1f}秒")
        processing_status[task_id].update({
            "status": "error", 
            "message": f"{error_msg} (耗时: {processing_duration:.1f}秒)",
            "completed_time": datetime.now().isoformat(),
            "processing_duration": processing_duration
        })

@app.get("/status/{task_id}")
async def get_status(task_id: str):
    """获取处理状态"""
    if task_id not in processing_status:
        raise HTTPException(status_code=404, detail="任务不存在")
    
    return processing_status[task_id]

@app.get("/download/{task_id}")
async def download_file(task_id: str):
    """下载生成的文件"""
    if task_id not in processing_status:
        raise HTTPException(status_code=404, detail="任务不存在")
    
    status = processing_status[task_id]
    if status["status"] != "completed":
        raise HTTPException(status_code=400, detail="文件尚未处理完成")
    
    file_path = Path(status["output_file"])
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    
    # 标记任务为可清理状态
    processing_status[task_id]["can_cleanup"] = True
    processing_status[task_id]["download_time"] = datetime.now().isoformat()
    
    logger.info(f"任务 {task_id} 文件已下载，标记为可清理")
    
    # 使用保存时的文件名作为下载文件名
    download_filename = file_path.name
    
    return FileResponse(
        path=file_path,
        filename=download_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

async def cleanup_expired_files():
    """清理过期的上传文件"""
    try:
        current_time = datetime.now()
        cleanup_count = 0
        
        for task_id, status in list(processing_status.items()):
            # 检查是否标记为可清理
            if not status.get("can_cleanup", False):
                continue
            
            # 检查下载时间是否超过1小时
            download_time_str = status.get("download_time")
            if not download_time_str:
                continue
                
            download_time = datetime.fromisoformat(download_time_str)
            if current_time - download_time < timedelta(hours=CLEANUP_DELAY_HOURS):
                continue
            
            # 清理上传文件夹
            upload_dir = status.get("upload_dir")
            if upload_dir and Path(upload_dir).exists():
                try:
                    shutil.rmtree(upload_dir)
                    logger.info(f"已清理任务 {task_id} 的上传文件夹: {upload_dir}")
                    cleanup_count += 1
                except Exception as e:
                    logger.warning(f"清理任务 {task_id} 文件夹失败: {e}")
            
            # 从状态中移除任务
            del processing_status[task_id]
            logger.info(f"已从状态中移除任务 {task_id}")
        
        if cleanup_count > 0:
            logger.info(f"本次清理完成，共清理 {cleanup_count} 个任务文件夹")
            
    except Exception as e:
        logger.error(f"文件清理过程中发生错误: {e}")

@app.on_event("startup")
async def startup_event():
    """应用启动时的初始化"""
    logger.info("应用启动，开始定时清理任务")
    
    # 启动定时清理任务
    asyncio.create_task(periodic_cleanup())

async def periodic_cleanup():
    """定期执行清理任务"""
    while True:
        try:
            await cleanup_expired_files()
            # 使用配置的清理间隔
            await asyncio.sleep(CLEANUP_CHECK_INTERVAL)
        except Exception as e:
            logger.error(f"定时清理任务出错: {e}")
            await asyncio.sleep(CLEANUP_ERROR_RETRY_INTERVAL)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000) 