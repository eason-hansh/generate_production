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
from typing import Optional, Union, List, Dict
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

def get_excel_template_path(customer_code: str) -> Dict[str, str]:
    """
    根据客户号获取Excel模板路径
    
    Args:
        customer_code: 客户号
        
    Returns:
        模板字典 {'GX': 'path1', 'GM': 'path2'} 或 {'GX': 'path1'}
        
    Raises:
        ValueError: 当客户号对应的模板目录不存在或没有找到有效模板时
    """
    template_dir = Path("company_templates") / customer_code
    
    if not template_dir.exists():
        raise ValueError(f"客户号 {customer_code} 对应的模板目录不存在")
    
    templates = {}
    for file_path in template_dir.glob("*.xlsx"):
        if "广美" in file_path.name:
            templates['GM'] = str(file_path)
        elif "广线" in file_path.name:
            templates['GX'] = str(file_path)
    
    if not templates:
        raise ValueError(f"客户号 {customer_code} 对应的模板目录中没有找到有效的模板文件")
    
    return templates

@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    """主页面"""
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/upload")
async def upload_files(
    background_tasks: BackgroundTasks,
    pdf_file: UploadFile = File(...),
    task_order_no: str = Form(...),
    order_date: str = Form(...),
    delivery_date: str = Form(...),
    customer_code: str = Form(...)
):
    """上传文件并处理"""
    try:
        # 生成唯一任务ID
        task_id = str(uuid.uuid4())
        
        # 为每个任务创建独立的文件夹
        task_upload_dir = UPLOAD_DIR / f"task_{task_id}"
        task_upload_dir.mkdir(exist_ok=True)
        
        # 保存上传的PDF文件
        pdf_path = task_upload_dir / "pdf_file.pdf"
        with open(pdf_path, "wb") as buffer:
            shutil.copyfileobj(pdf_file.file, buffer)
        
        # 以字典形式，根据客户号获取Excel模板路径
        templates_2_path = get_excel_template_path(customer_code)
        
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
        
        # 使用 BackgroundTasks 异步处理文件
        background_tasks.add_task(
            process_files,
            task_id, pdf_path, templates_2_path,
            task_order_no, order_date, delivery_date, pdf_name, customer_code
        )
        
        logger.info(f"任务 {task_id} 已添加到后台处理队列，上传目录: {task_upload_dir}")
        return {"task_id": task_id, "status": "processing"}
    
    except Exception as e:
        logger.error(f"文件上传失败: {str(e)}")
        raise HTTPException(status_code=500, detail=f"上传失败: {str(e)}")

async def process_files(task_id: str, pdf_path: Path, templates: Dict[str, str], 
                       task_order_no: str, order_date: str, delivery_date: str, pdf_name: str, customer_code: str):
    """ 后台处理文件 """
    # 记录开始时间
    start_time = datetime.now()
    
    try:
        logger.info(f"开始处理任务 {task_id}")
        processing_status[task_id]["message"] = "正在提取PDF信息..."
        
        # 1. 提取PDF信息
        pdf_info, template_2_info = pdf_extractor.process(str(pdf_path), templates)
        logger.info(f"任务 {task_id} PDF信息提取完成")
        
        processing_status[task_id]["message"] = "正在生成生产任务单..."
        
        # 2. 将抽取后的信息插入 Excel模板，并保存到输出路径
        saved_file_paths = excel_processor.process(pdf_info, template_2_info, task_order_no, order_date, delivery_date, str(OUTPUT_DIR), pdf_name, customer_code, task_id)
        
        # 计算处理时长
        end_time = datetime.now()
        processing_duration = (end_time - start_time).total_seconds()
        
        # 处理返回的文件路径（现在是字典格式）
        if isinstance(saved_file_paths, dict):
            # 字典格式：{'GX': 'path1', 'GM': 'path2'} 或 {'GX': 'path1'}
            # 由于文件已按客户号分文件夹保存，这里记录客户号用于下载
            saved_file_path = customer_code  # 使用客户号作为下载标识
        else:
            saved_file_path = saved_file_paths
        
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
async def download_task_files(task_id: str):
    """下载任务ID对应的所有Excel文件（打包为ZIP）"""
    import zipfile
    import tempfile
    
    # 根据任务ID查找对应的文件夹
    outputs_dir = Path("outputs")
    task_dirs = [d for d in outputs_dir.iterdir() 
                if d.is_dir() and d.name.endswith(f"_{task_id}")]
    
    if not task_dirs:
        raise HTTPException(status_code=404, detail=f"任务 {task_id} 的文件不存在")
    
    # 使用找到的任务文件夹
    task_dir = task_dirs[0]  # 应该只有一个匹配的文件夹
    
    # 检查是否有Excel文件
    excel_files = list(task_dir.glob("*.xlsx"))
    if not excel_files:
        raise HTTPException(status_code=404, detail=f"任务 {task_id} 没有找到Excel文件")
    
    # 创建临时ZIP文件
    with tempfile.NamedTemporaryFile(suffix='.zip', delete=False) as tmp_file:
        with zipfile.ZipFile(tmp_file.name, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for file_path in excel_files:
                # 将文件添加到ZIP中，保持原始文件名
                zip_file.write(file_path, file_path.name)
    
    # 从文件夹名中提取客户号
    customer_code = task_dir.name.split('_')[0]  # 提取客户号
    
    # 返回ZIP文件
    return FileResponse(
        tmp_file.name, 
        filename=f"{customer_code}_{task_id}_生产任务单.zip",
        media_type="application/zip"
    )

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