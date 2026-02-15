# -*- coding: utf-8 -*-
"""
腾讯云OCR综合工具 - 双选项卡界面
功能1：表格识别V3 - 图片/PDF转Excel
功能2：图纸图号识别 - PDF图纸批量重命名
作者：基于腾讯云官方SDK开发
"""
import os
import sys
import base64
import json
import re
import shutil
import time
import threading
import openpyxl
from io import BytesIO
from pathlib import Path
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from typing import Union, List, Optional
import cv2
import numpy as np
from PIL import Image
import fitz  # PyMuPDF

# 腾讯云SDK
from tencentcloud.common import credential
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.ocr.v20181119 import ocr_client, models

# ==================== 配置区域 ====================
# 从环境变量读取密钥（推荐）
TENCENT_SECRET_ID = "AKID62ub6KoNnDWkz50ymMq58mQxTp0161mO"
TENCENT_SECRET_KEY = "Zw9C5ttobWK0a5zztdDk6TjnnsxnRt8A"
DEFAULT_REGION = "ap-shanghai"  # 图纸识别推荐上海，表格识别推荐广州

# ==================== 图纸图号识别模块 ====================
class DrawingNumberRecognizer:
    """图纸图号识别类（基于原代码优化）"""
    
    def __init__(self, secret_id=None, secret_key=None, region="ap-shanghai"):
        self.secret_id = secret_id or TENCENT_SECRET_ID
        self.secret_key = secret_key or TENCENT_SECRET_KEY
        self.region = region
        self.temp_dir = "temp_drawing"
        self.output_dir = "图纸_已命名"
        
    def setup_temp_dir(self):
        """创建临时目录"""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
        os.makedirs(self.temp_dir, exist_ok=True)
        os.makedirs(self.output_dir, exist_ok=True)
    
    def cleanup_temp(self):
        """清理临时目录"""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
    
    def cv_imread(self, file_path):
        """解决imread不能读取中文路径的问题"""
        cv_img = cv2.imdecode(np.fromfile(file_path, dtype=np.uint8), -1)
        return cv_img
    
    def change_image(self, img, angle):
        """旋转图像"""
        if angle == 90:
            img = cv2.transpose(img)
            img = cv2.flip(img, flipCode=1)
            return img
        # 任意角度旋转
        cx, cy = img.shape[1] / 2, img.shape[0] / 2
        M = cv2.getRotationMatrix2D((cx, cy), angle, 1)
        rotated_img = cv2.warpAffine(img, M, (img.shape[1], img.shape[0]))
        return rotated_img
    
    def pdf_to_image(self, pdf_path, zoom=3):
        """PDF转PNG图片"""
        pdf = fitz.open(pdf_path)
        images = []
        for pg in range(len(pdf)):
            page = pdf[pg]
            trans = fitz.Matrix(zoom, zoom)
            pm = page.get_pixmap(matrix=trans, alpha=False)
            img_path = os.path.join(self.temp_dir, f"{Path(pdf_path).stem}_p{pg+1}.png")
            pm.save(img_path)
            images.append(img_path)
        pdf.close()
        return images
    
    def recognize_drawing_number(self, image_path):
        """
        识别图纸中的图号
        返回: [图号, 版本号, 部件标识] 如 ["DRM-2023-001", "-1", "-A"]
        """
        try:
            # 初始化OCR客户端
            cred = credential.Credential(self.secret_id, self.secret_key)
            httpProfile = HttpProfile()
            httpProfile.endpoint = "ocr.tencentcloudapi.com"
            clientProfile = ClientProfile("TC3-HMAC-SHA256")
            clientProfile.httpProfile = httpProfile
            client = ocr_client.OcrClient(cred, self.region, clientProfile)
            
            # 读取图片
            with open(image_path, 'rb') as f:
                base64_data = base64.b64encode(f.read())
                s = base64_data.decode()
            
            # 调用通用OCR
            req = models.GeneralBasicOCRRequest()
            params = '{"ImageBase64":"%s"}' % s
            req.from_json_string(params)
            
            resp = client.GeneralBasicOCR(req)
            resp = json.loads(resp.to_json_string())
            
            # 解析识别结果
            resp_list = resp.get('TextDetections', [])
            
            # 提取图号特征
            str0, str1, str2 = '', '-A', ''
            
            # 先识别版本标识
            for resp in resp_list:
                text = resp.get('DetectedText', '')
                match = re.findall(r'[1-9]/[1-9]', text)
                if len(match) >= 1:
                    str2 = match[0].replace('/', '-')
                    str2 = '-' + str2.split('-')[0]
                    if str2 == '-1' and str2.split('-')[1][0] != '2':
                        str2 = ''
                if 'A' in text:
                    str1 = '-A'
                if 'B' in text:
                    str1 = '-B'
            
            # 识别图号主体
            for resp in resp_list:
                result = resp.get('DetectedText', '')
                
                # 图号特征匹配
                if (('RM' in result and '-' in result) or 
                    ('ME' in result and '-' in result) or 
                    ('TF' in result and '.' in result) or 
                    (result.count('.') == 4)):
                    
                    result = result.replace(')', '1')
                    result = result.replace('图', '')
                    result = result.replace('号', '')
                    result = result.replace('专', '')
                    result = result.replace('+', '')
                    result = result.replace(' ', '|')
                    
                    result_parts = result.split('|')
                    for item in result_parts:
                        if 'ME' in item and '-' in item:
                            str0 = item
                        if 'RM' in item and '-' in item:
                            str0 = item
                        if 'TF' in item and '.' in item:
                            str0 = item
                        if item.count('.') == 4:
                            str0 = item
                    
                    if 'TF' in str0 and str0.find('TF') != 0:
                        str0 = str0[str0.find('TF'):]
                    if 'R' in str0 and 'DR' not in str0:
                        str0 = str0.replace('R', 'DR')
            
            # 清理结果
            if 'A' in str0:
                str0 = str0.split('A')[0]
            if 'B' in str0:
                str0 = str0.split('B')[0]
            
            str0 = str0.replace('/', '').replace('.', ' ').replace(':', ' ')
            str1 = str1.replace('.', ' ').replace(':', ' ')
            str2 = str2.replace('.', ' ').replace(':', ' ')
            
            return [str0.strip(), str2.strip(), str1.strip()]
            
        except Exception as e:
            print(f"识别失败: {str(e)}")
            return [None, None, None]
    
    def process_pdf_drawing(self, pdf_path, log_callback=None):
        """
        处理单个PDF图纸文件
        """
        def log(msg):
            if log_callback:
                log_callback(msg)
            else:
                print(msg)
        
        try:
            log(f"📄 处理文件: {os.path.basename(pdf_path)}")
            
            # PDF转图片
            img_paths = self.pdf_to_image(pdf_path, zoom=3)
            if not img_paths:
                log("❌ PDF转图片失败")
                return None
            
            # 处理第一页（通常图号在第一页）
            img_path = img_paths[0]
            
            # 读取图片并调整方向
            img_big = self.cv_imread(img_path)
            if img_big is None:
                log("❌ 无法读取图片")
                return None
            
            height, width = img_big.shape[:2]
            
            # 如果宽度小于高度，先旋转90度
            if width < height:
                img_big = self.change_image(img_big, 90)
                height, width = img_big.shape[:2]
            
            # 裁剪右下角区域（图号通常在这里）
            crop_x = int(width * 0.45)
            crop_y = int(height * 0.7)
            img_crop = img_big[crop_y:height, crop_x:width]
            
            # 保存裁剪图片
            crop_path = os.path.join(self.temp_dir, f"crop_{Path(img_path).name}")
            cv2.imencode('.png', img_crop)[1].tofile(crop_path)
            
            # 尝试多次旋转识别
            angles_to_try = [0, 180, 90, 270]
            best_result = [None, None, None]
            
            for angle in angles_to_try:
                if angle > 0:
                    rotated_img = self.change_image(img_big.copy(), angle)
                    height, width = rotated_img.shape[:2]
                    crop_x = int(width * 0.45)
                    crop_y = int(height * 0.7)
                    img_crop = rotated_img[crop_y:height, crop_x:width]
                    cv2.imencode('.png', img_crop)[1].tofile(crop_path)
                
                result = self.recognize_drawing_number(crop_path)
                if result[0] and result[0] not in ['', None]:
                    best_result = result
                    log(f"✅ 识别到图号: {result[0]}{result[1]}{result[2]}")
                    break
            
            if best_result[0]:
                # 生成新文件名
                new_filename = f"{best_result[0]}{best_result[1]}{best_result[2]}.pdf"
                new_path = os.path.join(self.output_dir, new_filename)
                
                # 处理重名
                counter = 1
                while os.path.exists(new_path):
                    name_part = f"{best_result[0]}{best_result[1]}{best_result[2]}"
                    new_filename = f"{name_part}_{counter}.pdf"
                    new_path = os.path.join(self.output_dir, new_filename)
                    counter += 1
                
                # 复制并重命名文件
                shutil.copy2(pdf_path, new_path)
                log(f"💾 已保存: {new_filename}")
                return new_path
            else:
                log("❌ 未识别到图号")
                return None
                
        except Exception as e:
            log(f"❌ 处理失败: {str(e)}")
            return None
    
    def batch_process(self, pdf_files, log_callback=None):
        """批量处理PDF图纸"""
        self.setup_temp_dir()
        
        success_count = 0
        fail_count = 0
        results = []
        
        for i, pdf_path in enumerate(pdf_files):
            if log_callback:
                log_callback(f"\n📌 进度: {i+1}/{len(pdf_files)}")
            
            result = self.process_pdf_drawing(pdf_path, log_callback)
            if result:
                success_count += 1
                results.append(result)
            else:
                fail_count += 1
        
        self.cleanup_temp()
        return success_count, fail_count, results


# ==================== 表格识别模块 ====================
class TableOCRRecognizer:
    """表格识别V3封装类"""
    
    def __init__(self, secret_id=None, secret_key=None, region="ap-guangzhou"):
        self.secret_id = secret_id or TENCENT_SECRET_ID
        self.secret_key = secret_key or TENCENT_SECRET_KEY
        self.region = region
    
    def recognize_from_image(self, image_input):
        """表格识别V3核心方法"""
        # 实例化认证对象
        cred = credential.Credential(self.secret_id, self.secret_key)
        
        # HTTP配置
        http_profile = HttpProfile()
        http_profile.endpoint = "ocr.tencentcloudapi.com"
        http_profile.reqTimeout = 60
        
        # 客户端配置
        client_profile = ClientProfile()
        client_profile.httpProfile = http_profile
        client_profile.signMethod = "TC3-HMAC-SHA256"
        
        # 初始化客户端
        client = ocr_client.OcrClient(cred, self.region, client_profile)
        
        # 处理图片输入
        if isinstance(image_input, str):
            with open(image_input, 'rb') as f:
                img_data = f.read()
        else:
            img_data = image_input
        
        # 构造请求
        req = models.RecognizeTableAccurateOCRRequest()
        req.ImageBase64 = base64.b64encode(img_data).decode('utf-8')
        
        # PDF处理
        if isinstance(image_input, str) and image_input.lower().endswith('.pdf'):
            req.IsPdf = True
            req.PdfPageNumber = 1
        
        # 发起请求
        resp = client.RecognizeTableAccurateOCR(req)
        excel_data = base64.b64decode(resp.Data)
        return excel_data


    def replace_in_excel_file(self, excel_data, pattern_replacements):
        """
        使用openpyxl处理Excel文件，安全地替换单元格内容
        pattern_replacements: 列表，每个元素为 (正则表达式, 替换字符串或函数)
        例如: [(r'中(\d+)', r'Φ\1')]  # 将"中6"替换为"Φ6"
        """
        try:
            # 将二进制数据加载为Excel工作簿
            excel_bytes = BytesIO(excel_data)
            wb = openpyxl.load_workbook(excel_bytes)
            
            # 遍历所有工作表
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                
                # 遍历所有单元格
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            original = cell.value
                            
                            # 应用所有正则规则
                            for pattern, repl in pattern_replacements:
                                cell.value = re.sub(pattern, repl, cell.value)
                            
                            # 如果发生变化，打印日志
                            if cell.value != original:
                                print(f"替换: '{original}' → '{cell.value}'")
            
            # 保存到新的BytesIO对象
            output_bytes = BytesIO()
            wb.save(output_bytes)
            output_bytes.seek(0)
            return output_bytes.read()
            
        except Exception as e:
            print(f"Excel处理失败: {e}")
            return excel_data

    def save_as_excel(self, image_input, output_path=None):
        """识别并保存为Excel文件"""
        excel_data = self.recognize_from_image(image_input)
        
        # ===== 使用正则表达式替换"中+任意数字"为"Φ+相同数字" =====
        pattern_replacements = [
            (r'中(\d)', r'Φ\1'),  # 中6 → Φ6, 中123 → Φ123
            # 可以添加更多正则规则
            # (r'直径(\d+)', r'Φ\1'),  # 直径6 → Φ6
        ]
        
        excel_data = self.replace_in_excel_file(excel_data, pattern_replacements)
        
        # 后续代码保持不变...
        if output_path is None:
            if isinstance(image_input, str):
                base_name = Path(image_input).stem
                output_path = f"{base_name}_识别结果.xlsx"
            else:
                output_path = "表格识别结果.xlsx"
        elif not output_path.endswith(('.xlsx', '.xls')):
            output_path += '.xlsx'
        
        # 处理重名
        counter = 1
        original_path = output_path
        while os.path.exists(output_path):
            name_part = Path(original_path).stem
            ext = Path(original_path).suffix
            if name_part.endswith(f"_{counter-1}"):
                name_part = name_part[:-3]
            output_path = f"{name_part}_{counter}{ext}"
            counter += 1
        
        with open(output_path, 'wb') as f:
            f.write(excel_data)
        
        return output_path


# ==================== 主GUI应用 ====================
class OCRTabbedApp:
    """双选项卡OCR综合工具"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("腾讯云OCR综合工具")
        self.root.geometry("900x600")
        self.root.minsize(800, 600)
        
        # 共享变量
        self.secret_id = StringVar(value=TENCENT_SECRET_ID)
        self.secret_key = StringVar(value=TENCENT_SECRET_KEY)
        self.table_region = StringVar(value="ap-guangzhou")
        self.drawing_region = StringVar(value="ap-shanghai")
        
        # 设置UI
        self.setup_ui()
        

    
    def setup_ui(self):
        """初始化用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(N, W, E, S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # ========== 标题 ==========
        # title_label = ttk.Label(
        #     main_frame,
        #     text="腾讯云OCR综合工具",
        #     font=("微软雅黑", 18, "bold")
        # )
        # title_label.grid(row=0, column=0, pady=(0, 15))

        
        # ========== 选项卡 ==========
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=2, column=0, sticky=(N, S, E, W), pady=(10, 0))
        main_frame.rowconfigure(2, weight=1)
        
        # 创建两个选项卡
        self.setup_table_tab()    # 表格识别选项卡
        self.setup_drawing_tab()  # 图纸识别选项卡
   
    
    def setup_table_tab(self):
        """表格识别选项卡"""
        tab = ttk.Frame(self.notebook, padding="15")
        self.notebook.add(tab, text="📊 表格识别V3")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(3, weight=1)
        

        
        # ===== 文件选择 =====
        file_frame = ttk.LabelFrame(tab, text="文件选择", padding="10")
        file_frame.grid(row=1, column=0, sticky=(W, E), pady=(0, 15))
        file_frame.columnconfigure(1, weight=1)
        
        # 表格文件变量
        self.table_files = []
        
        ttk.Button(
            file_frame,
            text="📁 选择图片/PDF",
            command=self.select_table_files,
            width=15
        ).grid(row=0, column=0, padx=(0, 10))
        
        self.table_file_label = ttk.Label(file_frame, text="未选择文件")
        self.table_file_label.grid(row=0, column=1, sticky=W)
        
        ttk.Button(
            file_frame,
            text="清空",
            command=self.clear_table_files,
            width=8
        ).grid(row=0, column=2, padx=(10, 0))
        
        # 文件列表
        self.table_listbox = Listbox(
            file_frame,
            height=4,
            selectmode=EXTENDED,
            activestyle='none'
        )
        self.table_listbox.grid(row=1, column=0, columnspan=3, sticky=(W, E), pady=(10, 0))
        
        # ===== 导出设置 =====
        export_frame = ttk.LabelFrame(tab, text="导出设置", padding="10")
        export_frame.grid(row=2, column=0, sticky=(W, E), pady=(0, 15))
        export_frame.columnconfigure(1, weight=1)
        
        ttk.Label(export_frame, text="导出位置:").grid(row=0, column=0, sticky=W, padx=(0, 5))
        self.table_output_label = ttk.Label(export_frame, text="未选择", foreground="gray")
        self.table_output_label.grid(row=0, column=1, sticky=W, padx=(0, 10))
        
        ttk.Button(
            export_frame,
            text="📂 浏览",
            command=self.select_table_output,
            width=8
        ).grid(row=0, column=2)
        
        # ===== 日志区域 =====
        log_frame = ttk.LabelFrame(tab, text="处理日志", padding="10")
        log_frame.grid(row=3, column=0, sticky=(N, S, E, W))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.table_log = Text(log_frame, height=12, wrap=WORD)
        self.table_log.grid(row=0, column=0, sticky=(N, S, E, W))
        
        table_scrollbar = ttk.Scrollbar(log_frame, orient=VERTICAL, command=self.table_log.yview)
        table_scrollbar.grid(row=0, column=1, sticky=(N, S))
        self.table_log.configure(yscrollcommand=table_scrollbar.set)
        
        # ===== 操作按钮 =====
        btn_frame = ttk.Frame(tab)
        btn_frame.grid(row=4, column=0, pady=(15, 0))
        
        self.table_progress = ttk.Progressbar(btn_frame, mode='determinate', length=300)
        self.table_progress.grid(row=0, column=0, padx=(0, 20))
        
        self.table_btn = ttk.Button(
            btn_frame,
            text="🚀 开始识别",
            command=self.start_table_recognition,
            width=15
        )
        self.table_btn.grid(row=0, column=1, padx=5)
        
        ttk.Button(
            btn_frame,
            text="清除日志",
            command=lambda: self.table_log.delete(1.0, END),
            width=10
        ).grid(row=0, column=2, padx=5)
    
    def setup_drawing_tab(self):
        """图纸图号识别选项卡"""
        tab = ttk.Frame(self.notebook, padding="15")
        self.notebook.add(tab, text="📐 图纸图号识别")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(3, weight=1)


        
        # ===== 文件选择 =====
        file_frame = ttk.LabelFrame(tab, text="PDF图纸文件", padding="10")
        file_frame.grid(row=1, column=0, sticky=(W, E), pady=(0, 15))
        file_frame.columnconfigure(1, weight=1)
        
        # 图纸文件变量
        self.drawing_files = []
        
        ttk.Button(
            file_frame,
            text="📁 选择PDF图纸",
            command=self.select_drawing_files,
            width=15
        ).grid(row=0, column=0, padx=(0, 10))
        
        self.drawing_file_label = ttk.Label(file_frame, text="未选择文件")
        self.drawing_file_label.grid(row=0, column=1, sticky=W)
        
        ttk.Button(
            file_frame,
            text="清空",
            command=self.clear_drawing_files,
            width=8
        ).grid(row=0, column=2, padx=(10, 0))
        
        # 文件列表
        self.drawing_listbox = Listbox(
            file_frame,
            height=4,
            selectmode=EXTENDED,
            activestyle='none'
        )
        self.drawing_listbox.grid(row=1, column=0, columnspan=3, sticky=(W, E), pady=(10, 0))
        
        # ===== 输出设置 =====
        output_frame = ttk.LabelFrame(tab, text="输出设置", padding="10")
        output_frame.grid(row=2, column=0, sticky=(W, E), pady=(0, 15))
        output_frame.columnconfigure(1, weight=1)
        
        ttk.Label(output_frame, text="输出目录:").grid(row=0, column=0, sticky=W, padx=(0, 5))
        self.drawing_output_label = ttk.Label(output_frame, text="默认: ./图纸_已命名", foreground="gray")
        self.drawing_output_label.grid(row=0, column=1, sticky=W, padx=(0, 10))
        
        ttk.Button(
            output_frame,
            text="📂 浏览",
            command=self.select_drawing_output,
            width=8
        ).grid(row=0, column=2)
        
        # ===== 日志区域 =====
        log_frame = ttk.LabelFrame(tab, text="处理日志", padding="10")
        log_frame.grid(row=3, column=0, sticky=(N, S, E, W))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.drawing_log = Text(log_frame, height=12, wrap=WORD)
        self.drawing_log.grid(row=0, column=0, sticky=(N, S, E, W))
        
        drawing_scrollbar = ttk.Scrollbar(log_frame, orient=VERTICAL, command=self.drawing_log.yview)
        drawing_scrollbar.grid(row=0, column=1, sticky=(N, S))
        self.drawing_log.configure(yscrollcommand=drawing_scrollbar.set)
        
        # ===== 操作按钮 =====
        btn_frame = ttk.Frame(tab)
        btn_frame.grid(row=4, column=0, pady=(15, 0))
        
        self.drawing_progress = ttk.Progressbar(btn_frame, mode='determinate', length=300)
        self.drawing_progress.grid(row=0, column=0, padx=(0, 20))
        
        self.drawing_btn = ttk.Button(
            btn_frame,
            text="🔍 开始识别图号",
            command=self.start_drawing_recognition,
            width=15
        )
        self.drawing_btn.grid(row=0, column=1, padx=5)
        
        ttk.Button(
            btn_frame,
            text="清除日志",
            command=lambda: self.drawing_log.delete(1.0, END),
            width=10
        ).grid(row=0, column=2, padx=5)
    
    
    # ========== 日志方法 ==========
    def log_table(self, message):
        """表格选项卡日志"""
        self.table_log.insert(END, f"{message}\n")
        self.table_log.see(END)
        self.root.update_idletasks()
    
    def log_drawing(self, message):
        """图纸选项卡日志"""
        self.drawing_log.insert(END, f"{message}\n")
        self.drawing_log.see(END)
        self.root.update_idletasks()
    
    # ========== 表格识别方法 ==========
    def select_table_files(self):
        """选择表格文件"""
        files = filedialog.askopenfilenames(
            title="选择图片或PDF文件",
            filetypes=[
                ("图像文件", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff"),
                ("PDF文件", "*.pdf"),
                ("所有支持格式", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.pdf"),
                ("所有文件", "*.*")
            ]
        )
        
        if files:
            for file in files:
                if file not in self.table_files:
                    self.table_files.append(file)
                    self.table_listbox.insert(END, os.path.basename(file))
            
            self.table_file_label.config(text=f"已选择 {len(self.table_files)} 个文件")
            self.log_table(f"📎 已添加 {len(files)} 个文件，当前共 {len(self.table_files)} 个文件")
    
    def clear_table_files(self):
        """清空表格文件列表"""
        self.table_files.clear()
        self.table_listbox.delete(0, END)
        self.table_file_label.config(text="未选择文件")
        self.log_table("🗑️ 已清空文件列表")
    
    def select_table_output(self):
        """选择表格输出目录"""
        directory = filedialog.askdirectory(title="选择Excel导出目录")
        if directory:
            self.table_output_dir = directory
            self.table_output_label.config(text=directory, foreground="black")
            self.log_table(f"📂 导出目录: {directory}")
    
    def start_table_recognition(self):
        """开始表格识别"""
        # 验证输入
        if not hasattr(self, 'table_output_dir') or not self.table_output_dir:
            messagebox.showwarning("提示", "请选择Excel导出位置")
            return
        
        if not self.table_files:
            messagebox.showwarning("提示", "请选择要识别的文件")
            return
        
        # 禁用按钮
        self.table_btn.config(state=DISABLED)
        
        # 在新线程中执行
        thread = threading.Thread(target=self.process_table_files, daemon=True)
        thread.start()
    
    def process_table_files(self):
        """处理表格文件"""
        try:
            recognizer = TableOCRRecognizer(
                self.secret_id.get(),
                self.secret_key.get(),
                self.table_region.get()
            )
            
            total = len(self.table_files)
            success = 0
            fail = 0
            
            self.log_table(f"\n{'='*50}")
            self.log_table(f"开始表格识别，共 {total} 个文件")
            self.log_table(f"{'='*50}")
            
            self.table_progress['maximum'] = total
            self.table_progress['value'] = 0
            
            for i, file_path in enumerate(self.table_files):
                file_name = os.path.basename(file_path)
                self.log_table(f"\n[{i+1}/{total}] 处理: {file_name}")
                
                try:
                    output_path = os.path.join(
                        self.table_output_dir,
                        f"{Path(file_name).stem}_识别结果.xlsx"
                    )
                    
                    saved_path = recognizer.save_as_excel(file_path, output_path)
                    self.log_table(f"✅ 成功: {os.path.basename(saved_path)}")
                    success += 1
                    
                except Exception as e:
                    self.log_table(f"❌ 失败: {str(e)}")
                    fail += 1
                
                self.table_progress['value'] = i + 1
                self.root.update_idletasks()
            
            self.log_table(f"\n{'='*50}")
            self.log_table(f"处理完成！成功: {success} 个，失败: {fail} 个")
            
            if success > 0:
                self.root.after(100, lambda: messagebox.showinfo(
                    "完成", 
                    f"表格识别完成！\n成功: {success} 个\n失败: {fail} 个\n保存位置: {self.table_output_dir}"
                ))
            
        except Exception as e:
            self.log_table(f"❌ 程序错误: {str(e)}")
        finally:
            self.table_btn.config(state=NORMAL)
            self.table_progress['value'] = 0
    
    # ========== 图纸识别方法 ==========
    def select_drawing_files(self):
        """选择图纸PDF文件"""
        files = filedialog.askopenfilenames(
            title="选择PDF图纸文件",
            filetypes=[
                ("PDF文件", "*.pdf"),
                ("所有文件", "*.*")
            ]
        )
        
        if files:
            for file in files:
                if file not in self.drawing_files:
                    self.drawing_files.append(file)
                    self.drawing_listbox.insert(END, os.path.basename(file))
            
            self.drawing_file_label.config(text=f"已选择 {len(self.drawing_files)} 个文件")
            self.log_drawing(f"📎 已添加 {len(files)} 个PDF图纸")
    
    def clear_drawing_files(self):
        """清空图纸文件列表"""
        self.drawing_files.clear()
        self.drawing_listbox.delete(0, END)
        self.drawing_file_label.config(text="未选择文件")
        self.log_drawing("🗑️ 已清空文件列表")
    
    def select_drawing_output(self):
        """选择图纸输出目录"""
        directory = filedialog.askdirectory(title="选择重命名后图纸保存目录")
        if directory:
            self.drawing_output_dir = directory
            self.drawing_output_label.config(text=directory, foreground="black")
            self.log_drawing(f"📂 输出目录: {directory}")
    
    def start_drawing_recognition(self):
        """开始图纸图号识别"""
        # 验证输入
        if not self.drawing_files:
            messagebox.showwarning("提示", "请选择PDF图纸文件")
            return
        
        # 禁用按钮
        self.drawing_btn.config(state=DISABLED)
        
        # 在新线程中执行
        thread = threading.Thread(target=self.process_drawing_files, daemon=True)
        thread.start()
    
    def process_drawing_files(self):
        """处理图纸文件"""
        try:
            # 设置输出目录
            if hasattr(self, 'drawing_output_dir'):
                DrawingNumberRecognizer.output_dir = self.drawing_output_dir
            else:
                DrawingNumberRecognizer.output_dir = "图纸_已命名"
            
            # 确保输出目录存在
            os.makedirs(DrawingNumberRecognizer.output_dir, exist_ok=True)
            
            recognizer = DrawingNumberRecognizer(
                self.secret_id.get(),
                self.secret_key.get(),
                self.drawing_region.get()
            )
            
            total = len(self.drawing_files)
            self.log_drawing(f"\n{'='*50}")
            self.log_drawing(f"开始图纸图号识别，共 {total} 个文件")
            self.log_drawing(f"{'='*50}")
            
            self.drawing_progress['maximum'] = total
            self.drawing_progress['value'] = 0
            
            success = 0
            fail = 0
            
            for i, pdf_path in enumerate(self.drawing_files):
                self.log_drawing(f"\n📌 进度: {i+1}/{total}")
                
                def log_callback(msg):
                    self.log_drawing(msg)
                    self.root.update_idletasks()
                
                try:
                    result = recognizer.process_pdf_drawing(pdf_path, log_callback)
                    if result:
                        success += 1
                    else:
                        fail += 1
                except Exception as e:
                    self.log_drawing(f"❌ 处理异常: {str(e)}")
                    fail += 1
                
                self.drawing_progress['value'] = i + 1
                self.root.update_idletasks()
            
            self.log_drawing(f"\n{'='*50}")
            self.log_drawing(f"批量处理完成！")
            self.log_drawing(f"✅ 成功: {success} 个")
            self.log_drawing(f"❌ 失败: {fail} 个")
            self.log_drawing(f"📁 保存目录: {recognizer.output_dir}")
            
            if success > 0:
                self.root.after(100, lambda: messagebox.showinfo(
                    "完成", 
                    f"图纸识别完成！\n✅ 成功命名: {success} 个\n❌ 识别失败: {fail} 个\n\n保存位置:\n{os.path.abspath(recognizer.output_dir)}"
                ))
            
        except Exception as e:
            self.log_drawing(f"❌ 程序错误: {str(e)}")
        finally:
            self.drawing_btn.config(state=NORMAL)
            self.drawing_progress['value'] = 0


# ==================== 程序入口 ====================
def main():
    """主函数"""
    root = Tk()
    app = OCRTabbedApp(root)
    
    # 窗口居中
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()