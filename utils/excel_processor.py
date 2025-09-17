import json
import asyncio
from openpyxl import load_workbook
from pathlib import Path
from typing import Dict, Any

from openpyxl import load_workbook
from openpyxl.styles import PatternFill


class ExcelProcessor:
    """
    Excel处理类，负责生成生产任务单

    ai_output_template = {
            "po_no": "提取的采购订单号",
            "product_info": [
                {"cust_item_code": "08484", "quantity": "50000"},
                {"cust_item_code": "08485", "quantity": "6000"}
            ]
        }
    """
    
    def __init__(self):
        pass
    
    def extract_customer_code_from_excel(self, excel_path):
        """
        从Excel模板的B1单元格提取客户号
        
        Args:
            excel_path: Excel文件路径
            
        Returns:
            客户号字符串，如果提取失败返回空字符串
        """
        try:
            wb = load_workbook(filename=excel_path, data_only=True)
            main_sheet = wb['主表']
            customer_code = main_sheet['B1'].value
            return str(customer_code) if customer_code else ""
        except Exception as e:
            print(f"从Excel提取客户号失败: {e}")
            return ""
    
    def convert_json_2_dict(self, raw_json: str) -> Dict[str, Any]:
        """将JSON字符串转换为字典"""
        # 提取 {}里的所有内容，再转为字典
        start_index = raw_json.find('{')
        if start_index == -1:
            raise ValueError("未检测到左大括号 {")

        # 定位最后一个右大括号
        end_index = raw_json.rfind('}')
        if end_index == -1:
            raise ValueError("未检测到右大括号 }")

        # 检查括号顺序有效性
        if end_index <= start_index:
            raise ValueError("右大括号位置在左大括号之前或相同")

        # 提取目标子串
        json_str = raw_json[start_index:end_index + 1]
        # 将所有的单引号 --> 双引号
        new_json_str = json_str.replace("'", '"')

        json_2_dict = json.loads(new_json_str)
        return json_2_dict

    def generate_task_order_no(self, raw_pdf_info, task_order_no):
        # 批量生成 任务单号
        left = task_order_no.find("(1)")
        if left == -1:
            raise ValueError("输入字符串必须包含 '(1)'")

        # 拆分前缀和后缀
        prefix = task_order_no[:left]  # "TW25040782"
        suffix = task_order_no[left + 3:]  # "BC"（跳过 "(1)"）

        product_num = len(raw_pdf_info['product_info']) # 除去 'po_no'
        
        # 生成范围格式的任务单号
        if product_num == 1:
            task_order_range = f"{prefix}(1){suffix}"
        else:
            task_order_range = f"{prefix}(1-{product_num}){suffix}"
        
        # 生成单个任务单号列表（用于Excel填写）
        task_orders = [f"{prefix}({i}){suffix}" for i in range(1, product_num + 1)]

        pdf_info = {}

        for index, item in enumerate(raw_pdf_info['product_info']):
            pdf_info[item['cust_item_code']] = {
                'quantity': item['quantity'],
                'task_order_no': task_orders[index]
            }

        return pdf_info, task_order_range

    def process(self, raw_pdf_info, excel_path, task_order_no, order_date, delivery_date, output_dir=None, pdf_name=None):
        # 加载 Excel 文件，注意使用 data_only=False 以保留公式
        wb = load_workbook(filename=excel_path, data_only=False)

        # 将sheet的修改部分，高亮出来，创建高亮样式（黄色背景）
        formula_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # 获取主表
        main_sheet = wb['主表']
        modify_sheets = ['主表']

        # 填写基本信息填写
        main_sheet['E1'] = order_date   # 制单日期
        main_sheet['B3'] = raw_pdf_info['po_no']  # PO NO
        main_sheet['B2'] = delivery_date  # 交货期

        # 生成任务单号
        pdf_info, task_order_range = self.generate_task_order_no(raw_pdf_info, task_order_no)
        
        # 提取客户号
        customer_code = self.extract_customer_code_from_excel(excel_path)

        exit_flag = False
        for row in main_sheet.iter_rows():
            try:
                if exit_flag and row[0].value:
                    product_name = str(row[0].value) if type(row[0].value) is int else row[0].value.split()[0]
                    row_idx = row[0].row  # 获取当前行号
                    # 匹配到产品则进行修改，否则，则将改行进行隐藏
                    if product_name in pdf_info:
                        # 修改数量
                        main_sheet[f"D{row_idx}"] = int(pdf_info[product_name]['quantity'])
                        # 修改 任务单号
                        main_sheet[f"G{row_idx}"] = pdf_info[product_name]['task_order_no']
                        main_sheet.row_dimensions[row_idx].hidden = False  # 该行一定不隐藏
                        modify_sheets.append(str(row[0].value))
                    else:
                        # 修改数量
                        main_sheet[f"D{row_idx}"] = 0
                        # 修改生产任务单
                        main_sheet[f"G{row_idx}"] = ''
                        # # 外箱流水号转为空
                        # main_sheet[f"G{row_idx}"] = ''
                        # 将改行隐藏
                        main_sheet.row_dimensions[row_idx].hidden = True

                if row[0].value and "客户货号" in str(row[0].value):
                    exit_flag = True
                    continue
            except Exception as e:
                print(e)

        # 将修改后的sheet显示，其余的隐藏
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            sheet.sheet_state = 'visible' if sheet_name in modify_sheets else 'hidden'

            # 高亮公式单元格，只处理可见工作表
            if sheet.sheet_state == 'visible':
                for row in sheet.iter_rows():
                    for cell in row:
                        # 检查公式单元格（兼容不同格式）
                        if cell.data_type == 'f' or (cell.value and str(cell.value).startswith('=')):
                            cell.fill = formula_fill

        # 生成输出路径
        if output_dir is None:
            output_dir = Path('.')
        else:
            output_dir = Path(output_dir)
            output_dir.mkdir(exist_ok=True)

        # 生成文件名：任务单号_PO号_客户号
        po_no = raw_pdf_info.get('po_no', '')
        if po_no:
            po_prefix = f"PO{po_no}"
        else:
            po_prefix = "PO"
        
        if customer_code:
            filename = f"{task_order_range}_{po_prefix}_{customer_code}.xlsx"
        else:
            # 如果客户号提取失败，使用原来的命名方式
            filename = f"{task_order_range}_{po_prefix}.xlsx"

        output_path = output_dir / filename

        wb.save(output_path)
        return str(output_path)