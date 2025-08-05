"""
系统配置文件
"""

import os
from pathlib import Path

# 基础路径配置
BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
TEMPLATE_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"

# 文件上传配置
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
ALLOWED_PDF_EXTENSIONS = {".pdf"}
ALLOWED_EXCEL_EXTENSIONS = {".xlsx", ".xls"}

# 处理配置
PROCESSING_TIMEOUT = 300  # 5分钟超时
STATUS_CHECK_INTERVAL = 2  # 状态检查间隔（秒）

# AI接口配置（可选）
AI_API_URL = os.getenv("AI_API_URL", None)
AI_API_KEY = os.getenv("AI_API_KEY", None)

# Excel模板配置
EXCEL_MAIN_SHEET = "主表"
EXCEL_PO_CELL = "B3"
EXCEL_ORDER_DATE_CELL = "E1"
EXCEL_DELIVERY_DATE_CELL = "B2"
EXCEL_QUANTITY_COLUMN = "D"
EXCEL_TASK_ORDER_COLUMN = "G"
EXCEL_BOX_SERIAL_COLUMN = "G"

# 产品信息配置
PRODUCT_CODE_ROW_START_MARKER = "客户货号"

# 开发配置
DEBUG = os.getenv("DEBUG", "True").lower() == "true"
HOST = os.getenv("HOST", "0.0.0.0")
PORT = int(os.getenv("PORT", "8000"))

# 日志配置
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"

# 安全配置
SECRET_KEY = os.getenv("SECRET_KEY", "your-secret-key-here")
ACCESS_TOKEN_EXPIRE_MINUTES = 30

# 数据库配置（如果将来需要）
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./production_tasks.db")

# 邮件配置（如果将来需要）
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USERNAME = os.getenv("SMTP_USERNAME", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")

# 缓存配置
CACHE_TTL = 3600  # 1小时

# 文件清理配置
CLEANUP_INTERVAL_HOURS = 1  # 清理间隔（小时）
CLEANUP_DELAY_HOURS = 1  # 下载后延迟清理时间（小时）
CLEANUP_CHECK_INTERVAL = 3600  # 清理检查间隔（秒）
CLEANUP_ERROR_RETRY_INTERVAL = 300  # 清理错误重试间隔（秒）

# 性能配置
MAX_CONCURRENT_TASKS = 10
TASK_QUEUE_SIZE = 100

# 监控配置
ENABLE_METRICS = os.getenv("ENABLE_METRICS", "False").lower() == "true"
METRICS_PORT = int(os.getenv("METRICS_PORT", "9090"))

# 国际化配置
DEFAULT_LANGUAGE = "zh-CN"
SUPPORTED_LANGUAGES = ["zh-CN", "en-US"]

# 主题配置
THEME_COLORS = {
    "primary": "#667eea",
    "secondary": "#764ba2",
    "success": "#28a745",
    "warning": "#ffc107",
    "danger": "#dc3545",
    "info": "#17a2b8"
}

# 功能开关
FEATURES = {
    "file_upload": True,
    "ai_extraction": True,
    "excel_processing": True,
    "download": True,
    "status_tracking": True,
    "file_cleanup": True,
    "user_authentication": False,  # 未来功能
    "batch_processing": False,     # 未来功能
    "template_management": False,  # 未来功能
} 