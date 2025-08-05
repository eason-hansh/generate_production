#!/usr/bin/env python3
"""
PDF信息提取模块
使用marker和dashscope(Qwen API)提取PDF中的采购订单信息
"""

import json
import re
from pathlib import Path
from typing import Dict, Any, Optional

import dashscope
from http import HTTPStatus

from marker.converters.pdf import PdfConverter
from marker.models import create_model_dict
from marker.output import text_from_rendered


dashscope.api_key = 'sk-ecf4a4044d444e75a8c232cb39250066'

class PDFExtractor:
    """PDF信息提取器"""
    
    def __init__(self):
        """初始化PDF提取器"""

        self.converter = PdfConverter(
            artifact_dict=create_model_dict(),
        )
        # logger.info("pdf converter started")
        # 设计提示词提取 PO 中的 po_no 和 采购商品的货号、数量
        ai_output_template = {
            "po_no": "提取的采购订单号",
            "product_info": [
                {"cust_item_code": "08484", "quantity": "50000"},
                {"cust_item_code": "08485", "quantity": "6000"}
            ]
        }
        self.prompt = (f"你是一位专业的信息采集员，擅长从文本中精准提取指定字段并以JSON格式返回。请严格遵循以下要求：\n"
                       f"## 任务\n"
                       f"从输入的采购订单中提取：\n"
                       f"1. 采购订单号（po_no）：识别 'PURCHASE ORDER NUMBER'、'PO'等关键字后的字母数字组合；\n"
                       f"2. 所有采购商品信息：每个商品的'货号(custom item code)'和'数量(quantity)'。\n"
                       f"## 输出格式\n"
                       f"{ai_output_template}\n"
                       f"## 强制规则\n"
                       f"1. 只返回JSON格式的数据，不要包含任何额外的解释或文本；\n"
                       f"2. 若无法找到po_no，该字段值为空字符串；\n"
                       f"3. '货号(custom item code)' 仅考虑 'EPM Part'对应的信息；\n"
                       f"4. product_info必须是数组，即使只有1个商品；\n"
                       f"5. quantity字段必须为整数类型（无逗号、无小数点）；\n"
                       f"## 待处理采购订单：\n")

    def parse_pdf(self, pdf_path):
        rendered = self.converter(pdf_path)
        text, _, images = text_from_rendered(rendered)

        return text

    @staticmethod
    def call_api(prompt):
        messages = [
            {'role': 'user', 'content': prompt}
        ]
        response = dashscope.Generation.call(
            'qwen-long',
            messages=messages,
            temperature=0.1,
            result_format='message'  # set the result is message format.
        )
        if response.status_code == HTTPStatus.OK:
            output = response['output']['choices'][0]['message']['content']
            input_tokens = response['usage']['input_tokens']
            output_tokens = response['usage']['output_tokens']
            return output, input_tokens, output_tokens

        else:
            print('Request id: %s, Status code: %s, error code: %s, error message: %s' % (
                response.request_id, response.status_code,
                response.code, response.message
            ))
            return "{}", 0, 0

    def extract_info(self, parsed_pdf):
        # 基于 qwen api,以提示词的方式抽取信息
        prompt = (f"{self.prompt}\n"
                  f"{parsed_pdf}")
        raw_production_info, input_tokens, output_tokens = self.call_api(prompt)
        # logger.info(f"raw_ai_result: {raw_production_info}")
        # logger.info(f"input_token_nums: {input_tokens}")
        # logger.info(f"output_token_nums: {output_tokens}")

        return raw_production_info

    @staticmethod
    def convert_json_2_dict(raw_json):
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

        try:
            # 尝试解析JSON
            json_2_dict = json.loads(new_json_str)
            # logger.info(f"ai_output: {json_2_dict}")
            return json_2_dict

        except json.JSONDecodeError as e:
            # 增强错误信息提示
            error_msg = f"JSON解析失败: {e.msg}\n错误位置: {e.pos}\n错误片段: {new_json_str[max(0, e.pos - 20):e.pos + 20]}"
            raise ValueError(error_msg) from e

    def process(self, pdf_path):
        """
        1. 解析 pdf --> markdown
        2. ai extract raw_info
        3. raw info --> dict
        :return:
        """
        pdf_2_markdown = self.parse_pdf(pdf_path)
        product_info_str = self.extract_info(pdf_2_markdown)
        po_product_info_dict = self.convert_json_2_dict(product_info_str)

        return po_product_info_dict