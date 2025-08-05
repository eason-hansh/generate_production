# 生产任务单生成系统

一个基于FastAPI和现代Web技术的生产任务单自动生成系统，支持PDF采购订单信息提取和Excel模板自动填充。

## 功能特性

- 📄 **PDF信息提取**: 通过AI技术自动提取采购订单中的关键信息
- 📊 **Excel模板处理**: 基于现有Excel模板自动生成生产任务单
- 🎨 **现代化界面**: 简约大方的Web界面，支持拖拽上传
- ⚡ **异步处理**: 支持大文件处理和实时状态反馈
- 📱 **响应式设计**: 适配各种设备屏幕

## 系统架构

```
modify_by_link/
├── main.py                 # FastAPI主应用
├── requirements.txt        # 项目依赖
├── script.py              # 原始测试脚本
├── utils/                 # 工具模块
│   ├── __init__.py
│   ├── excel_processor.py # Excel处理逻辑
│   └── pdf_extractor.py   # PDF信息提取
├── templates/             # HTML模板
│   └── index.html         # 主页面
├── static/                # 静态文件
├── uploads/               # 上传文件存储
├── outputs/               # 生成文件存储
└── data/                  # 示例数据
```

## 安装和运行

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 运行应用

```bash
python main.py
```

或者使用uvicorn：

```bash
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

### 3. 访问系统

打开浏览器访问：`http://localhost:8000`

## 使用说明

### 1. 上传文件
- 上传采购订单PDF文件
- 上传Excel模板文件
- 支持点击选择或拖拽上传

### 2. 填写参数
- **任务单号**: 生产任务单的唯一标识
- **制单日期**: 任务单制作日期
- **交货期**: 产品交货日期

### 3. 处理流程
1. 点击"开始处理"按钮
2. 系统自动提取PDF信息
3. 根据模板生成生产任务单
4. 下载生成的Excel文件

## 技术栈

### 后端
- **FastAPI**: 现代、快速的Web框架
- **openpyxl**: Excel文件处理
- **asyncio**: 异步处理支持

### 前端
- **HTML5**: 语义化标记
- **CSS3**: 现代化样式设计
- **JavaScript**: 交互逻辑
- **Font Awesome**: 图标库

## 配置说明

### AI接口集成

在 `utils/pdf_extractor.py` 中修改 `extract_info_with_ai_api` 方法：

```python
async def extract_info_with_ai_api(self, pdf_path: str, ai_api_url: str = None):
    # 集成你的AI API
    import requests
    
    with open(pdf_path, 'rb') as f:
        files = {'file': f}
        response = requests.post(ai_api_url, files=files)
        return response.json()
```

### Excel模板配置

确保Excel模板包含以下工作表：
- `Bảng tóm tắt总表`: 主表，包含产品信息
- 其他相关工作表: 根据产品代码自动显示/隐藏

## API接口

### 上传文件
```
POST /upload
Content-Type: multipart/form-data

参数:
- pdf_file: PDF文件
- excel_template: Excel模板文件
- task_order_no: 任务单号
- order_date: 制单日期
- delivery_date: 交货期
```

### 查询状态
```
GET /status/{task_id}
```

### 下载文件
```
GET /download/{task_id}
```

## 开发说明

### 项目结构
- `main.py`: 主应用入口，包含所有API路由
- `utils/excel_processor.py`: Excel处理核心逻辑，基于原有script.py
- `utils/pdf_extractor.py`: PDF信息提取模块，可集成AI接口
- `templates/index.html`: 前端界面

### 扩展功能
1. **用户认证**: 添加登录和权限控制
2. **文件管理**: 添加文件历史记录和版本管理
3. **批量处理**: 支持多个文件同时处理
4. **模板管理**: 支持多种Excel模板

## 注意事项

1. **文件格式**: 确保上传的PDF和Excel文件格式正确
2. **模板结构**: Excel模板必须包含指定的工作表名称
3. **AI接口**: 需要配置有效的AI信息提取接口
4. **存储空间**: 定期清理uploads和outputs目录

## 故障排除

### 常见问题

1. **文件上传失败**
   - 检查文件格式是否正确
   - 确认文件大小是否超限

2. **Excel处理错误**
   - 验证Excel模板结构
   - 检查工作表名称是否正确

3. **AI提取失败**
   - 确认AI接口配置
   - 检查PDF文件质量

## 许可证

MIT License

## 联系方式

如有问题或建议，请联系开发团队。 