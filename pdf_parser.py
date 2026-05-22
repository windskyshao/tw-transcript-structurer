# enhanced_second_class_parser.py - 一二類謄本解析器（修正Excel欄位順序）
# ---------------------------------------------------------------------------------
# 專門處理土地登記第一二類謄本和建物登記第一二類謄本
# 重點修正：
# 1. 支援第一二類謄本的格式差異
# 2. 處理更詳細的他項權利記錄
# 3. 支援跨頁續次頁處理
# 4. 優化共同擔保地號/建號的長列表解析
# 5. 處理多筆他項權利記錄
# 6. 修正Excel輸出的欄位順序（建築完成日期在層次之後）
# ---------------------------------------------------------------------------------
from __future__ import annotations
import sys
from datetime import datetime
import re
import fitz  # PyMuPDF
import cv2
import numpy as np
import os
import time
import warnings
import logging
import subprocess
from typing import Dict, List, Tuple, Optional, Any, Union
from PIL import Image
import json
import pandas as pd
from datetime import datetime
import glob
from collections import OrderedDict

class DebugLogger:
    def __init__(self):
        self.logs = []
        self.log_file = None
        self.log_filepath = None
        # 🔧 在初始化時就保存真正的 print 函數
        import builtins
        import sys
        self.original_print = builtins.print
        self.original_stdout = sys.stdout

    def start_logging(self, log_filepath=None):
        """開始收集並即時寫入所有print輸出"""
        logger = self  # 捕獲 self 到閉包中

        # 🔥 如果提供了檔案路徑，開啟檔案準備即時寫入
        if log_filepath:
            try:
                logger.log_filepath = log_filepath
                os.makedirs(os.path.dirname(log_filepath), exist_ok=True)
                logger.log_file = open(log_filepath, 'w', encoding='utf-8')
                # 寫入檔案標題
                logger.log_file.write("="*80 + "\n")
                logger.log_file.write("🔍 謄本解析即時除錯日誌\n")
                logger.log_file.write("="*80 + "\n")
                logger.log_file.write(f"開始時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                logger.log_file.flush()
            except Exception as e:
                logger.original_print(f"[DebugLogger警告] 無法開啟日誌檔案: {e}")
                logger.log_file = None

        def debug_print(*args, **kwargs):
            # 保留原本的print功能
            logger.original_print(*args, **kwargs)
            # 同時記錄到logs列表（向後相容）
            try:
                message = ' '.join(str(arg) for arg in args)
                timestamp = datetime.now().strftime('%H:%M:%S')
                log_entry = f"[{timestamp}] {message}"
                logger.logs.append(log_entry)

                # 🔥 即時寫入檔案
                if logger.log_file:
                    logger.log_file.write(log_entry + "\n")
                    logger.log_file.flush()  # 立即刷新到磁碟
            except Exception as e:
                logger.original_print(f"[DebugLogger錯誤] 無法記錄日誌: {e}")

        # 替換全域的print函數
        import builtins
        builtins.print = debug_print
        self.original_print("[DebugLogger] 日誌記錄已啟動，即時寫入模式")

    def close_log_file(self):
        """關閉日誌檔案"""
        if self.log_file:
            try:
                self.log_file.write("\n" + "="*80 + "\n")
                self.log_file.write(f"結束時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                self.log_file.write(f"記錄總數: {len(self.logs)} 筆\n")
                self.log_file.write("="*80 + "\n")
                self.log_file.close()
                self.original_print(f"📄 即時除錯日誌已保存到: {self.log_filepath}")
            except Exception as e:
                self.original_print(f"[DebugLogger錯誤] 關閉日誌檔案失敗: {e}")
            finally:
                self.log_file = None
    
    def save_logs(self, filename="debug_output.txt"):
        """保存所有日誌到檔案"""
        try:
            if not self.logs:
                self.original_print(f"⚠️ 警告：日誌列表為空，沒有內容可保存")
                return None
                
            with open(filename, 'w', encoding='utf-8') as f:
                f.write("="*80 + "\n")
                f.write("🔍 謄本解析完整除錯日誌\n")
                f.write("="*80 + "\n")
                f.write(f"生成時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"記錄總數: {len(self.logs)} 筆\n\n")
                
                for log in self.logs:
                    f.write(log + "\n")
                
                f.write("\n" + "="*80 + "\n")
                f.write("🏁 日誌結束\n")
                f.write("="*80 + "\n")
            
            self.original_print(f"📄 完整除錯日誌已保存到: {filename} (共 {len(self.logs)} 筆記錄)")
            return filename
        except Exception as e:
            self.original_print(f"❌ 日誌保存失敗: {e}")
            import traceback
            traceback.print_exc()
            return None

# 全域日誌收集器
debug_logger = DebugLogger()
# 🔧 先不啟動日誌，等到有輸出目錄時再啟動（讓日誌即時寫入）
# debug_logger.start_logging()

# 🔧 新增：全域變數記錄最後使用的輸出目錄
_last_output_dir = None
_logging_started = False

def _set_output_dir(output_dir: str):
    """記錄輸出目錄並啟動即時日誌（由 exporter 呼叫）"""
    global _last_output_dir, _logging_started
    _last_output_dir = output_dir

    # 🔥 如果日誌還沒啟動，現在啟動並指定檔案路徑
    if not _logging_started:
        log_path = os.path.join(output_dir, "transcript_debug.txt")
        debug_logger.start_logging(log_path)
        _logging_started = True

# 🔥 使用 atexit 確保程式結束時關閉日誌檔案
import atexit
def _close_debug_log():
    """程式結束時關閉日誌檔案"""
    try:
        debug_logger.close_log_file()
    except Exception as e:
        print(f"⚠️ 關閉日誌檔案失敗: {e}")

atexit.register(_close_debug_log)

# ---- OCR住址佇列（每份PDF處理時重置；按小圖出現順序存放）----
ocr_address_queue = []

# --- 內嵌小圖片字：通用工具（自動補缺字到任意欄位） ---

def is_inline_glyph(s: str) -> bool:
    '''是否為 1~2 個 CJK 字（去空白後）。'''
    if not s:
        return False
    s = re.sub(r"\s+", "", s)
    return bool(re.fullmatch(r"[\u4e00-\u9fff]{1,2}", s))

def normalize_cjk_simple_to_trad(s: str) -> str:
    '''簡→繁（小表，可再擴充）'''
    mapping = {
        "馆":"館","门":"門","区":"區","台":"臺","华":"華","国":"國","东":"東","业":"業",
        "镇":"鎮","县":"縣","厦":"廈","广":"廣","阳":"陽","阴":"陰","万":"萬","岗":"崗",
        "鳳山品":"鳳山區","燕巢品":"燕巢區","苓稚":"苓雅","术":"術","左警":"左營",
        " 後段":"廍後段",
        "廊":"廍",
    }
    return "".join(mapping.get(ch, ch) for ch in s)

def patch_inline_glyphs_in_text(text: str, glyphs: list) -> str:
    '''智慧補字：只在合理的地方補字'''
    if not text or not glyphs:
        return text
    
    t = text
    
    # 🎯 最高優先級：主要建材補字
    if '鋼' in t and glyphs:
        for i, glyph in enumerate(glyphs):
            if glyph in ['鉄']:
                # 各種主要建材模式
                patterns = [
                    ('鋼 ＲＣ造', f'鋼{glyph}ＲＣ造')
                ]
                
                for old_pattern, new_pattern in patterns:
                    if old_pattern in t:
                        result = t.replace(old_pattern, new_pattern)
                        if result != t:
                            glyphs.pop(i)
                            print(f"    ✅ 主要建材補字: {t} → {result}")
                            # 🔧 在return前先做OCR修正
                            result = apply_ocr_corrections(result)
                            return result

    # 🔧 新增：建築用途補字（在禁止檢查之前）
    if '爐房' in t and glyphs:
        # 檢查「爐房」前面是否缺字
        glyph_used = False
        for i, glyph in enumerate(glyphs):
            if glyph in ['大']:  # 如果有「大」字可用
                # 在「爐房」前面補字
                if '爐房' in t and '大爐房' not in t:  # 確保不是已經有「大爐房」
                    result = t.replace('爐房', f'{glyph}爐房')
                    if result != t:
                        glyphs.pop(i)
                        print(f"    ✅ 建築用途補字: {t} → {result}")
                        # 🔧 在return前先做OCR修正
                        result = apply_ocr_corrections(result)
                        return result
                        
    # 🔥 新增：所有權人姓氏補字（第二類謄本中姓氏可能是圖片）
    # 檢測是否只有遮罩符號（如「＊＊」或「**」）
    mask_only_pattern = re.fullmatch(r'[＊\*]+', t.strip())
    if mask_only_pattern and glyphs:
        # 所有權人只有遮罩符號，在前面補上姓氏
        glyph = glyphs.pop(0)
        result = glyph + t.strip()
        print(f"    ✅ 所有權人姓氏補字: {t} → {result}")
        return apply_ocr_corrections(result)

    # 🚫 禁止補字的場景
    forbidden_contexts = [
        '謄本',          # 不要補到謄本標題
        '第二類',        # 不要補到謄本類型
        '建號',          # 不要隨便補到建號
        '地號',          # 不要隨便補到地號
        '00053-000',     # 不要補到具體的號碼
        '0681-0000',     # 不要補到具體的號碼
        '民國',          # 不要補到日期
        '年',           # 不要補到年份
        '月',           # 不要補到月份
        '日',           # 不要補到日期
    ]
    
    # 檢查是否在禁止補字的上下文中
    for forbidden in forbidden_contexts:
        if forbidden in t:
            # print(f"    🚫 禁止補字：文本包含 '{forbidden}'")
            # 🔧 即使禁止補字，也要做OCR修正
            return apply_ocr_corrections(t)
    
    # 🎯 合理的補字場景
    # 1) 住址中的路街補字（原有邏輯保留）
    m = re.search(r'(\S)\s+((?:[\u4e00-\u9fff]{1,4})(?:路|街|大道|巷|弄|段))', t)
    if m:
        g = normalize_cjk_simple_to_trad(glyphs.pop(0))
        result = t.replace(m.group(0), m.group(1) + g + m.group(2), 1)
        # 🔧 在return前先做OCR修正
        return apply_ocr_corrections(result)
    
    # 2) 地段名稱補字（但要更嚴格）
    if '段' in t and not any(num in t for num in ['00', '11', '22', '33']):  # 排除編號
        # 只在真正的地段名稱中補字
        segment_patterns = [
            r'([^0-9])\s+段',  # X段
            r'([^0-9])\s+里',  # X里
        ]
        
        for pattern in segment_patterns:
            m = re.search(pattern, t)
            if m and glyphs:
                g = normalize_cjk_simple_to_trad(glyphs.pop(0))
                result = re.sub(pattern, m.group(1) + g + '段', t, count=1)
                # 🔧 在return前先做OCR修正
                return apply_ocr_corrections(result)
    
    # 🔧 新增：最終的OCR誤識別修正（即使沒有補字也要修正）
    final_result = apply_ocr_corrections(t)
    if final_result != t:
        return final_result
    
    # 如果都不符合，不補字
    # print(f"    ⚪ 無合適補字場景，保持原文: {t}")
    return t


def apply_ocr_corrections(text: str) -> str:
    """簡化版OCR修正（補字階段使用）"""
    result = text
    
    # OCR誤識別修正規則
    ocr_corrections = [
        ('大爐房', '煱爐房'),    # 大爐房 → 煱爐房
        ('大爐', '煱爐'),        # 大爐 → 煱爐
    ]
    
    for wrong_text, correct_text in ocr_corrections:
        if wrong_text in result:
            result = result.replace(wrong_text, correct_text)
            print(f"    🔧 外部OCR修正: {wrong_text} → {correct_text}")
            break
    
    return result

def patch_mapping_all_strings(d: dict, glyphs: list):
    '''
    對 dict 中**所有字串欄位**進行「多回合」補字：
    只要還有 glyphs 且本輪有任何欄位被改動，就再掃一輪。
    '''
    if not d or not glyphs:
        return
    changed = True
    while glyphs and changed:
        changed = False
        for k, v in list(d.items()):
            if isinstance(v, str) and v.strip():
                new_v = patch_inline_glyphs_in_text(v, glyphs)
                if new_v != v:
                    d[k] = new_v
                    changed = True

def deep_patch_glyphs(obj, glyphs):
    """遞迴補字：對任意巢狀的 list/dict 中的所有字串值進行補字。"""
    if not glyphs:
        return obj
    if isinstance(obj, str):
        return patch_inline_glyphs_in_text(obj, glyphs)
    if isinstance(obj, list):
        for i, v in enumerate(obj):
            obj[i] = deep_patch_glyphs(v, glyphs)
        return obj
    if isinstance(obj, dict):
        # 先針對目前層級的所有字串欄位做一次
        patch_mapping_all_strings(obj, glyphs)
        # 再遞迴處理子結構
        for k, v in list(obj.items()):
            obj[k] = deep_patch_glyphs(v, glyphs)
        return obj
    return obj

# ==================== Pillow 庫相容性修補 ====================
def fix_pillow_compatibility():
    """修補 Pillow 庫的相容性問題"""
    try:
        constants_to_fix = {
            'ANTIALIAS': getattr(Image, 'LANCZOS', 1),
            'NEAREST': 0,
            'BOX': 4,
            'BILINEAR': 2,
            'HAMMING': 5,
            'BICUBIC': 3,
            'LANCZOS': 1
        }
        
        patched = 0
        for name, value in constants_to_fix.items():
            if not hasattr(Image, name):
                setattr(Image, name, value)
                patched += 1
        
        if not hasattr(Image, 'Resampling'):
            class Resampling:
                NEAREST = 0
                BOX = 4
                BILINEAR = 2
                HAMMING = 5
                BICUBIC = 3
                LANCZOS = 1
                ANTIALIAS = 1
            
            Image.Resampling = Resampling
            patched += 1
            
        if patched > 0:
            print(f"🔧 Pillow 相容性修補完成 ({patched} 項)")
        return True
        
    except Exception as e:
        print(f"⚠️ Pillow 修補失敗: {e}")
        return False

fix_pillow_compatibility()

warnings.filterwarnings("ignore")
logging.getLogger("rapidocr_onnxruntime").setLevel(logging.ERROR)

# ==================== 全域設定 ====================
AUTO_INSTALL = True        
TARGET_HEIGHT = 60         # 目標圖像高度（像素）
PREFER_TRADITIONAL = True  

print(f"🔧 全域設定:")
print(f"  目標圖片高度: {TARGET_HEIGHT}")
print(f"  自動安裝套件: {AUTO_INSTALL}")
print(f"  偏好繁體中文: {PREFER_TRADITIONAL}")

# ==================== 繁簡字轉換功能 ====================
def traditional_to_simplified_map():
    """建立繁體字到簡體字的對照表"""
    return {
        '黄': '黃', '统': '統', '编': '編', '豊': '豐', '粼': '鄰',
        '権': '權', '学': '學', '毫': '臺', '崋': '華', '楼': '樓',
        '証': '證', '険': '險', '杨': '楊', '鞋': '華',
        '左管': '左營', '同魔': '同慶', '債权': '債權', '金璧': '金鑾',
        '左警': '左營',
        '5公': '5鄰公',
        '鳳山品': '鳳山區',
        '燕巢品': '燕巢區',
        '高雄巿': '高雄市',  # 巿 vs 市
        '臺南巿': '臺南市',
        '屏東巿': '屏東市',
        ' 後段': '廍後段',
        '廊': '廍',
        '苓稚': '苓雅',
        '镇': '鎮',
        '岗': '崗',
        '东': '東',
        '区': '區',
        '县': '縣',
        '门': '門',
        '华': '華',
        '国': '國',
        '业': '業',
        '术': '術',
        '号': '號',  # 簡繁轉換：号碼 → 號碼
        '挥': '揮',  # 簡繁轉換：挥 → 揮
        '台北': '臺北',  # 台北 → 臺北
        '台中': '臺中',  # 台中 → 臺中
        '台南': '臺南',  # 台南 → 臺南
        '台东': '臺東',  # 台东 → 臺東
        '经': '經',  # 经貿 → 經貿
        '贸': '貿',  # 经贸 → 經貿
        '昌林': '員林',  # OCR 誤識：昌林 → 員林
        '彰化縣林市': '彰化縣員林市',  # OCR 誤識：彰化縣林市 → 彰化縣員林市
        '彰化縣昌林市': '彰化縣員林市',  # OCR 誤識：彰化縣昌林市 → 彰化縣員林市
    }

def convert_to_traditional(text: str) -> str:
    """將文字轉換為繁體中文"""
    conv_map = traditional_to_simplified_map()
    result = text
    for simplified, traditional in conv_map.items():
        result = result.replace(simplified, traditional)
    return result

# ==================== 自動安裝工具 ====================
def auto_install(pkg: str, mod: str = None, extra: str = None) -> bool:
    """自動安裝 Python 套件"""
    name = (mod or pkg.split("==")[0]).replace("-", "_")
    
    try:
        __import__(name)
        return True
    except ImportError:
        if not AUTO_INSTALL:
            return False
        
        cmd = [sys.executable, "-m", "pip", "install"] + pkg.split()
        if extra:
            cmd += extra.split()
            
        print(f"    🔄 pip install {pkg}")
        
        try:
            subprocess.check_call(cmd)
            __import__(name)
            print(f"    ✅ {pkg} 安裝完成")
            return True
        except Exception as e:
            print(f"    ❌ {pkg} 安裝失敗: {e}")
            return False
def extract_pdf_text_only(pdf_path: str) -> Optional[str]:
    """提取PDF文字"""
    try:
        doc = fitz.open(pdf_path)
        all_text = []
        
        for page_num, page in enumerate(doc):
            text_dict = page.get_text("dict")
            for block in text_dict["blocks"]:
                if block.get("type") == 0:
                    for line in block.get("lines", []):
                        line_text = ""
                        for span in line.get("spans", []):
                            if span.get("text"):
                                line_text += span["text"]
                        if line_text.strip():
                            all_text.append(line_text.strip())
        
        doc.close()
        return "\n".join(all_text)
        
    except Exception as e:
        print(f"[錯誤] PDF 文字提取失敗: {e}")
        return None
        
# def extract_pdf_text_only(pdf_path: str) -> Optional[str]:
#     """提取PDF文字"""
#     try:
#         doc = fitz.open(pdf_path)
#         all_text = []
        
#         print(f"📄 開始提取 {os.path.basename(pdf_path)} 的文字...")
        
#         for page_num, page in enumerate(doc):
#             # 先取得整頁文字來檢查
#             page_text = page.get_text()
            
#             # 🔍 調試：檢查每頁是否包含年月地價
#             if "055年06月" in page_text:
#                 print(f"  ✅ 第{page_num+1}頁包含 055年06月")
#             if "101年05月" in page_text:
#                 print(f"  ✅ 第{page_num+1}頁包含 101年05月")
#             if re.search(r'\d{3}年\d{2}月.*?\d+\.\d+元', page_text):
#                 print(f"  📍 第{page_num+1}頁包含年月和地價資料")
            
#             # 繼續原本的處理
#             text_dict = page.get_text("dict")
#             for block in text_dict["blocks"]:
#                 if block.get("type") == 0:
#                     for line in block.get("lines", []):
#                         line_text = ""
#                         for span in line.get("spans", []):
#                             if span.get("text"):
#                                 line_text += span["text"]
#                         if line_text.strip():
#                             all_text.append(line_text.strip())
        
#         doc.close()
        
#         result = "\n".join(all_text)
        
#         # 🔍 調試：最終檢查
#         print(f"📊 文字提取完成統計：")
#         if "055年06月" in result:
#             print("  ✅ 最終文本包含 055年06月")
#         else:
#             print("  ❌ 最終文本不包含 055年06月")
        
#         if "101年05月" in result:
#             print("  ✅ 最終文本包含 101年05月")  
#         else:
#             print("  ❌ 最終文本不包含 101年05月")
        
#         # 計算年月地價配對
#         year_month_pattern = r'(\d{3}年\d{2}月)'
#         year_months = re.findall(year_month_pattern, result)
#         print(f"  📊 找到 {len(year_months)} 個年月資料")
#         if year_months:
#             print(f"     前5個: {year_months[:5]}")
        
#         return result
        
#     except Exception as e:
#         print(f"[錯誤] PDF 文字提取失敗: {e}")
#         import traceback
#         traceback.print_exc()
#         return None

# ==================== OCR 管理器 ====================
class TranscriptOCRManager:
    """謄本專用 OCR 管理器"""
    
    def __init__(self):
        print("🔧 開始初始化OCR管理器（通用）...")
        self.engines: Dict[str, object] = {}
        self._init_engines()
        
        # 🔧 OCR引擎狀態檢查
        print(f"\n🔍 OCR引擎狀態檢查:")
        print(f"  引擎數量: {len(self.engines)}")
        
        if "rapidocr" in self.engines:
            print(f"  ✅ RapidOCR引擎: 已載入")
            
            # 測試OCR引擎是否可用
            try:
                import numpy as np
                test_img = np.ones((60, 200, 3), dtype=np.uint8) * 255
                test_result = self.ocr_image(test_img)
                print(f"  🧪 OCR引擎測試: {'✅ 正常' if isinstance(test_result, str) else '❌ 異常'}")
            except Exception as e:
                print(f"  ❌ OCR引擎測試失敗: {e}")
        else:
            print(f"  ❌ RapidOCR引擎: 未載入")
        
        print(f"🎯 OCR管理器初始化完成\n")

    def _init_engines(self):
        print("\n🔄 初始化 OCR 引擎 (RapidOCR)...")

        if auto_install("rapidocr-onnxruntime", mod="rapidocr_onnxruntime"):
            from rapidocr_onnxruntime import RapidOCR
            self.engines["rapidocr"] = RapidOCR()
            print("  ✅ RapidOCR 啟用")

        print(f"  共 {len(self.engines)} 套 OCR 引擎載入完成\n")

    def _preprocess_image(self, img: np.ndarray, is_small: bool = False) -> np.ndarray:
        """圖片預處理

        Args:
            img: 輸入圖片
            is_small: 是否為小圖片（如單字），使用特殊處理
        """
        if len(img.shape) == 3:
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        else:
            gray = img

        if is_small:
            # 🔥 小圖片（單字）使用 OTSU 二值化 + 放大2倍
            _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            h, w = binary.shape[:2]
            result = cv2.resize(binary, (w * 2, h * 2), interpolation=cv2.INTER_CUBIC)
            print(f"    🔍 小圖片特殊處理: OTSU二值化 + 放大2倍 ({w}x{h} → {w*2}x{h*2})")
            return result
        else:
            # 一般圖片使用原有處理
            alpha = 1.3
            beta = 10
            enhanced = cv2.convertScaleAbs(gray, alpha=alpha, beta=beta)
            kernel = np.array([[-1, -1, -1], [-1, 9, -1], [-1, -1, -1]])
            sharpened = cv2.filter2D(enhanced, -1, kernel)
            return sharpened

    def ocr_image(self, img: np.ndarray) -> str:
        try:
            # 🔥 檢測是否為小圖片（如單字圖片）
            h, w = img.shape[:2]
            is_small = (w < 100 and h < 100)

            processed_img = self._preprocess_image(img, is_small=is_small)

            engine = self.engines["rapidocr"]
            res, _ = engine(processed_img)

            result = " ".join(t[1] for t in res) if res else ""
            return convert_to_traditional(result)

        except Exception as e:
            print(f"  ⚠️ OCR 失敗: {e}")
            return ""


# ==================== 通用謄本解析器 ====================
class UnifiedTranscriptParser:
    """通用謄本解析器"""
    
    def __init__(self):
        self.patterns = self._build_second_class_patterns()

    def _build_second_class_patterns(self):
        """建立謄本專用的正則表達式模式"""
        return {
            # === 基本資訊（完全複製一類程式） ===
            "謄本類型": r"([土地建物]+登記第[一二]類謄本.*?)(?:\n|（)",
            "地號建號": r"([\u4e00-\u9faf]+(?:鄉|鎮|市|巿|區)[\u4e00-\u9faf]+段(?:[\u4e00-\u9faf]*小段)?\s+[\d\-]+(?:地號|建號))",
            # "列印時間": r"列印時間[:：]\s*([^\n]+?)(?:\s+頁次[:：].*?)?(?:\n|$)",
            "列印時間": r"列印時間[:：]\s*([^\s]+(?:\s+[^\s]+)*?)(?:\s+頁次[:：]|\s*$)",
            "申請人": r"由([^\s]+)自行列印",
            "謄本種類碼": r"謄本種類碼[:：]\s*([^\s,，]+)",
            "機關名稱": r"([\u4e00-\u9faf]+地政事務所)",
            
            # === 標示部通用 ===
            "標示_登記日期": r"登記日[期日][:：]\s*([^\s]+(?:\s+[^\s]+)*?)(?:\s+(?:登記原因)|$)",
            "標示_登記原因": r"登記原因[:：]\s*([^\s\n]+)",
            
            # === 土地專用欄位 ===
            "面積": r"面\s*積[:：]\s*(\**[\d,]+\.?\d*\s*平方公尺)",
            "使用分區": r"使用分區[:：]\s*([^使\n]+?)(?=\s*使用地?類別|$)",
            "使用地類別": r"使用地?類別[:：]\s*([^\n]+)",
            "公告土地現值": r"(民國\d+年\d+月)\s+公告土地現值[:：]\s*(\**[\d,]+\.?\d*\s*元[／/]平方公尺)",
            "地上建物建號": r"地上建物建號[:：]\s*([^\n]+)",
            
            # === 建物專用欄位 ===
            "建物門牌": r"建物門牌[:：]\s*([^\n]+)",
            "建物坐落地號": r"建物坐落地號[:：]\s*([^\n]+)",
            "主要用途": r"主要用途[:：]\s*([^\n]+)",
            "主要建材": r"主要建材[:：]\s*(.*?)(?=層\s*數[:：])",
            "層數": r"層\s*數[:：]\s*(\d+層)",
            "總面積": r"總面積[:：]\s*(\**[\d,]+\.?\d*\s*平方公尺)",
            "建築完成日期": r"建築完成日[期日][:：]\s*([^\n]+)",
            
            # === 所有權部 ===
            "所有權_登記日期": r"登記日[期日][:：]\s*([^\s]+(?:\s+[^\s]+)*?)(?:\s+(?:登記原因|權利範圍)|$)",
            "所有權_登記原因": r"登記原因[:：]\s*([^\s\n]+)",
            "所有權_原因發生日期": r"原因發生日期[:：]\s*([^\s\n]+)",
            "所有權人": r"所有權人[:：]\s*(.+?)(?=\s*統一編號)",  # 🔥 支援跨行匹配到「統一編號」之前
            "統一編號": r"統一編號[:：]\s*([A-Z0-9\*]+)",
            "出生日期": r"出生日[期日][:：]\s*([^\s\n]+)",
            "住址": r"住\s*址[:：]\s*([^\n]+)",
            "權利範圍": r"權利範圍[:：]\s*([^\n]+)",
            "權狀字號": r"權狀字號[:：]\s*([^\s]+)",
            "當期申報地價": r"當期申報地價[:：]\s*([^\n]+)",
            "前次移轉現值": r"前次移轉現值或原規定地價[:：]\s*([^\n]+)",
            "歷次取得權利範圍": r"歷次取得權利範圍[:：]\s*([^\n]+)",
            "相關他項權利登記次序": r"相關他項權利登記次序[:：]\s*([^\n]+)",
            
            # === 他項權利部 ===
            "權利種類": r"權利種類[:：]\s*([^\s\n]+)",
            "收件年期": r"收件年期[:：]\s*([^\s]+)",
            "字號": r"字號[:：]\s*([^\s\n]+)",
            "他項_登記日期": r"登記日[期日][:：]\s*([^\s]+(?:\s+[^\s]+)*?)(?:\s+登記原因|$)",
            "他項_登記原因": r"登記原因[:：]\s*([^\s\n]+)",
            "權利人": r"權\s*利\s*人[:：]\s*([^\n]+)",
            "統一編號_他項": r"統一編號[:：]\s*([A-Z0-9\*\s]+)",
            "住址_他項": r"(?:住|居)\s*址[:：]\s*([^\n]+(?:\n(?!債權額比例|擔保債權總金額)[^\n]+)*)",
            "債權額比例": r"債權額比例[:：]\s*([^\n]+)",
            "擔保債權總金額": r'擔保債權總金額[:：]\s*(.*?(?:元正|元整|元))',
            "存續期間": r"存續期間[:：]\s*([^\n]+)",
            "擔保債權種類及範圍": r'擔保債權種類及範圍[:：]\s*(.*?)(?=\s*擔保債權確定期[日期][:：])',
            "擔保債權確定期日": r"擔保債權確定期[日期][:：]\s*([^\n。]*?)(?=\s*[。\n]|清償日[期日][:：])",
            "清償日期": r"清償日[期日][:：]\s*(.*?)(?=\s*利息\s*[（(]?率[）)]?[:：])",
            "利息(率)": r"利息\s*[（(]?率[）)]?[:：]\s*(.*?)(?=\s*遲延利息\s*[（(]?率[）)]?[:：])",
            "遲延利息(率)": r"遲延利息\s*[（(]?率[）)]?[:：]\s*(.*?)(?=\s*違\s*約\s*金[:：])",
            "違約金": r"違\s*約\s*金[:：]\s*([^\n]+?)(?=\s*其他擔保範圍約定[:：]|\s*債務人及債務額比例[:：]|\s*權利標的[:：]|\n|$)",
            "其他擔保範圍約定": r"其他擔保範圍約定[:：]\s*(.*?)(?=\s*債務人及債務額比例[:：]|\s*權利標的[:：]|\s*標的登記次序[:：])",
            "債務人及債務額比例": r"債務人及債務額比例[:：]\s*(.*?)(?=\s*權利標的[:：]|$)",
            "權利標的": r"權利標的[:：]\s*([^\s\n]+)",
            "標的登記次序": r"標的登記次序[:：]\s*([^\n]+)",
            "設定權利範圍": r"設定權利範圍[:：]\s*([^\n]+)",
            "證明書字號": r"證明書字號[:：]\s*([^\s\n]+)",
            "設定義務人": r"設定義務人[:：]\s*([^\n]+)",
            
            # === 共同擔保（多行支援） ===
            "共同擔保地號": r"共同擔保地號[:：]\s*(.*?)(?=\s*共同擔保建號|其他登記事項|$)",
            "共同擔保建號": r"共同擔保建號[:：]\s*(.*?)(?=\s*其他登記事項|$)",
        }

    def _check_if_record_needs_ocr_unified(self, record: Dict[str, Any]) -> bool:
        """統一邏輯：檢查所有權人記錄是否需要OCR補充（支援一類+二類）"""
        
        owner = record.get('所有權人', '').strip()
        id_number = record.get('統一編號', '').strip()
        birth_date = record.get('出生日期', '').strip()
        address = record.get('住址', '').strip()
        
        print(f"      🔍 統一OCR需求檢查:")
        print(f"         所有權人: {'✅' if owner else '❌'} {owner}")
        print(f"         統一編號: {'✅' if id_number else '❌'} {id_number}")
        print(f"         出生日期: {'✅' if birth_date else '❌'} {birth_date}")
        print(f"         住址: {'✅' if address else '❌'} {address[:30]}...")
        
        # 🔧 統一判斷邏輯：任何欄位缺失都可能需要OCR
        
        # 1. 完全沒有基本資料，肯定需要OCR
        if not owner and not id_number:
            print(f"      📝 缺少基本資料，需要OCR")
            return True
        
        # 2. 有基本資料，檢查完整性
        if owner and id_number:
            # 2a. 法人檢查（8位數字統一編號）
            if len(id_number) == 8 and id_number.isdigit():
                print(f"      🏢 法人資料檢查")
                if address:
                    print(f"      ✅ 法人資料完整，不需要OCR")
                    return False
                else:
                    print(f"      📝 法人缺少住址，需要OCR")
                    return True
            
            # 2b. 二類個人檢查（有*號）
            elif "*" in owner or "*" in id_number:
                print(f"      👤 二類個人資料檢查")
                if address:
                    print(f"      ✅ 二類個人住址已有，不需要OCR")
                    return False
                else:
                    print(f"      📝 二類個人缺少住址，需要OCR")
                    return True
            
            # 2c. 一類個人檢查（完整顯示）
            else:
                print(f"      👤 一類個人資料檢查")
                if owner and id_number and birth_date and address:
                    print(f"      ✅ 一類個人資料完整，不需要OCR")
                    return False
                else:
                    missing = []
                    if not birth_date: missing.append("出生日期")
                    if not address: missing.append("住址")
                    print(f"      📝 一類個人缺少: {missing}，需要OCR")
                    return True
        
        # 3. 其他情況都需要OCR
        print(f"      📝 資料不完整，需要OCR")
        return True

    def _supplement_record_with_ocr_unified(self, record: Dict[str, Any], ocr_data: Dict[str, str]) -> Dict[str, Any]:
        """統一OCR資料補充邏輯（支援一類+二類）"""
        
        print(f"      🔄 統一OCR資料補充:")
        
        # 建立補充後的記錄
        supplemented = record.copy()
        
        # 補充邏輯：缺什麼補什麼
        ocr_fields = ["所有權人", "統一編號", "出生日期", "住址"]
        
        for field in ocr_fields:
            current_value = supplemented.get(field, "").strip()
            ocr_value = ocr_data.get(field, "").strip()
            
            # 如果當前欄位為空且OCR有值，進行補充
            if not current_value and ocr_value:
                supplemented[field] = ocr_value
                print(f"         ✅ 補充 {field}: {ocr_value}")
            
            # 特殊情況：住址優化（OCR的住址可能更完整）
            elif field == "住址" and current_value and ocr_value:
                if len(ocr_value) > len(current_value) or "*" in current_value:
                    supplemented[field] = ocr_value
                    print(f"         🔄 住址優化: {current_value} → {ocr_value}")
                else:
                    print(f"         ⏭️ 保留原住址: {current_value}")
            
            else:
                print(f"         ⏭️ 保留 {field}: {current_value}")
        
        return supplemented
        
    def detect_transcript_class(self, text: str) -> str:
        """檢測謄本類別（第一類或第二類）"""
        if re.search(r'登記第二類謄本', text):
            return "第二類"
        elif re.search(r'登記第一類謄本', text):
            return "第一類"
        else:
            return "未知類別"

    def detect_document_type(self, text: str) -> str:
        """檢測謄本類型（土地或建物）"""
        if re.search(r'建物登記第[一二]類謄本', text):
            return "建物謄本"
        elif re.search(r'土地登記第[一二]類謄本', text):
            return "土地謄本"
        else:
            building_keywords = ['建物門牌', '主要用途', '主要建材', '層次', '總面積']
            land_keywords = ['使用分區', '使用地類別', '公告土地現值']
            
            building_count = sum(1 for keyword in building_keywords if keyword in text)
            land_count = sum(1 for keyword in land_keywords if keyword in text)
            
            return "建物謄本" if building_count > land_count else "土地謄本"

    def enhanced_clean_pagination(self, text: str) -> str:
        """修正：跨頁清理 - 保護地段建號行"""
        
        print("🧹 開始跨頁清理（保護重要行）...")
        
        lines = text.split('\n')
        protected_lines = []
        
        for i, line in enumerate(lines):
            original_line = line
            stripped_line = line.strip()
            
            # 🛡️ 保護重要行：包含地段建號的行
            if (('建號' in stripped_line or '地號' in stripped_line) and 
                ('段' in stripped_line) and 
                ('區' in stripped_line or '鄉' in stripped_line or '鎮' in stripped_line or '市' in stripped_line)):
                print(f"  🛡️ 保護地段建號行(第{i+1}行): {stripped_line}")
                protected_lines.append(original_line)
                continue
            
            # 🛡️ 保護謄本標題行
            if ('登記第' in stripped_line and '類謄本' in stripped_line):
                print(f"  🛡️ 保護謄本標題行(第{i+1}行): {stripped_line}")
                protected_lines.append(original_line)
                continue
            
            # 正常的跨頁清理邏輯
            should_keep = True
            
            # 刪除續次頁標記
            if '續次頁' in stripped_line:
                should_keep = False
            
            # 刪除單獨的地政事務所行
            elif stripped_line.endswith('地政事務所') and len(stripped_line) < 20:
                should_keep = False
            
            # 刪除頁次行
            elif '頁次：' in stripped_line and '列印時間' not in stripped_line:
                should_keep = False
            
            if should_keep:
                protected_lines.append(original_line)
            else:
                print(f"  🗑️ 清理行(第{i+1}行): {stripped_line}")
        
        result = '\n'.join(protected_lines)
        print(f"✅ 跨頁清理完成，保護了重要行")
        
        return result

    def enhanced_clean_pagination_for_single_transcript(self, text: str) -> str:
        """專門處理單份謄本的跨頁清理 - 改進版：處理所有位置的跨頁"""
        
        print("🧹 開始單份謄本跨頁清理（改進版）...")
        
        lines = text.split('\n')
        protected_lines = []
        i = 0
        
        while i < len(lines):
            original_line = lines[i]
            stripped_line = original_line.strip()
            
            # 🛡️ 保護前2行（謄本標題和地段建號）
            if i < 2:
                protected_lines.append(original_line)
                print(f"  🛡️ 保護前2行(第{i+1}行): {stripped_line[:60]}...")
                i += 1
                continue
            
            # 🔍 新增：檢測任何位置的續次頁標記
            if '續次頁' in stripped_line or '(續次頁)' in stripped_line or '（續次頁）' in stripped_line:
                print(f"  🔍 發現續次頁標記(第{i+1}行): {stripped_line[:60]}...")
                
                # 🔧 關鍵改進：智慧處理跨頁內容
                # 策略：續次頁前後的內容可能重複，需要比對並去重
                
                # 收集續次頁前的內容（往前看最多10行）
                before_content = []
                for j in range(max(0, i-10), i):
                    if lines[j].strip() and '續次頁' not in lines[j]:
                        before_content.append(lines[j].strip())
                
                # 收集續次頁後的內容（往後看最多10行）
                after_content = []
                j = i + 1
                while j < min(len(lines), i + 11):
                    if lines[j].strip() and '續次頁' not in lines[j]:
                        # 檢查是否為頁首重複（包含地號/建號 + 列印時間）
                        if re.search(r'[區鄉鎮市].*?段.*?\d+-\d+[地建]號.*?列印時間', lines[j]):
                            print(f"    🗑️ 跳過頁首重複(第{j+1}行)")
                            j += 1
                            continue
                        after_content.append(lines[j].strip())
                    j += 1
                
                # 🔧 比對策略：如果續次頁後的內容包含續次頁前的內容，說明是重複
                is_duplicate = False
                
                if before_content and after_content:
                    # 將前後內容簡化以便比對（移除空格、星號等）
                    before_text = ''.join(before_content)
                    before_simplified = re.sub(r'[\s\*]+', '', before_text)
                    
                    after_text = ''.join(after_content[:len(before_content)])  # 只比對相同數量的行
                    after_simplified = re.sub(r'[\s\*]+', '', after_text)
                    
                    # 如果內容相似度很高，視為重複
                    if before_simplified and after_simplified:
                        # 計算相似度（簡單方法：看是否包含）
                        if before_simplified in after_simplified or after_simplified in before_simplified:
                            is_duplicate = True
                            print(f"    🔄 偵測到跨頁重複內容")
                        
                        # 或者檢查是否有80%以上的內容相同
                        elif len(before_simplified) > 20 and len(after_simplified) > 20:
                            common_length = 0
                            for k in range(min(len(before_simplified), len(after_simplified))):
                                if before_simplified[k] == after_simplified[k]:
                                    common_length += 1
                            
                            similarity = common_length / min(len(before_simplified), len(after_simplified))
                            if similarity > 0.8:
                                is_duplicate = True
                                print(f"    🔄 偵測到高相似度({similarity:.1%})跨頁內容")
                
                # 🔧 處理跨頁
                if is_duplicate:
                    # 如果是重複，跳過續次頁標記行，並跳過重複的內容
                    print(f"    🗑️ 移除續次頁標記及重複內容")
                    
                    # 移除續次頁行本身
                    # 不加入 protected_lines
                    
                    # 跳過後續重複的行（數量與前面相同）
                    skip_lines = min(len(before_content), len(after_content))
                    i += skip_lines + 1  # +1 是續次頁行本身
                    print(f"    ⏭️ 跳過 {skip_lines} 行重複內容")
                else:
                    # 不是重複，只移除續次頁標記，保留內容
                    print(f"    ✅ 只移除續次頁標記，保留內容")
                    # 移除續次頁標記但保留該行其他內容
                    cleaned_line = re.sub(r'[（(]續次頁[）)]', '', original_line)
                    if cleaned_line.strip():
                        protected_lines.append(cleaned_line)
                    i += 1
                
                continue
            
            # 🔧 一般的跨頁清理邏輯（保持原有的清理規則）
            should_keep = True
            
            # 清理單獨的地政事務所行
            if stripped_line.endswith('地政事務所') and len(stripped_line) < 20:
                should_keep = False
                print(f"  🗑️ 清理地政事務所(第{i+1}行): {stripped_line}")
            
            # 清理重複的頁次行
            elif '頁次：' in stripped_line and '列印時間' not in stripped_line:
                should_keep = False
                print(f"  🗑️ 清理頁次行(第{i+1}行): {stripped_line}")
            
            # 清理跨頁後的頁首重複
            elif re.search(r'^\s*[0-9A-F]{1,4}\s+[0-9A-F]{1,4}\s+.*?段.*?[地建]號.*?列印時間', stripped_line):
                should_keep = False
                print(f"  🗑️ 清理跨頁頁首(第{i+1}行): {stripped_line[:60]}...")
            
            if should_keep:
                protected_lines.append(original_line)
            
            i += 1
        
        result = '\n'.join(protected_lines)
        
        # 🔍 最終檢查
        remaining_cross_page = result.count('續次頁')
        if remaining_cross_page > 0:
            print(f"  ⚠️ 還有 {remaining_cross_page} 個續次頁標記未處理")
        else:
            print(f"  ✅ 所有續次頁標記已清理")
        
        print(f"✅ 單份謄本跨頁清理完成")
        
        return result

    def split_sections_enhanced(self, text: str) -> Dict[str, str]:
        """謄本章節分割 - 處理更複雜的結構"""
        sections = {}
        
        print("    🔍 謄本章節分割...")
        
        # 基本資訊（從開始到第一個標示部）
        basic_pattern = r"^(.*?)(?=\*{10,}.*?標示部)"
        basic_match = re.search(basic_pattern, text, re.DOTALL)
        if basic_match:
            sections["基本資訊"] = basic_match.group(1).strip()
            print(f"      ✅ 基本資訊: {len(sections['基本資訊'])} 字元")
        
        # 標示部
        indicator_pattern = r"(\*{10,}\s*(?:[土地建物]*)?標示部\s*\*{10,})(.*?)(?=\*{10,}.*?所有權部|$)"
        indicator_match = re.search(indicator_pattern, text, re.DOTALL)
        if indicator_match:
            sections["標示部"] = indicator_match.group(2).strip()
            print(f"      ✅ 標示部: {len(sections['標示部'])} 字元")
        
        # 所有權部
        ownership_pattern = r"(\*{10,}\s*(?:[土地建物]*)?所有權部\s*\*{10,})(.*?)(?=\*{10,}.*?他項權利部|$)"
        ownership_match = re.search(ownership_pattern, text, re.DOTALL)
        if ownership_match:
            sections["所有權部"] = ownership_match.group(2).strip()
            print(f"      ✅ 所有權部: {len(sections['所有權部'])} 字元")
        
        # 他項權利部（包含更多內容）
        # 🔧 修正：簡化結束條件，只匹配「〈 本謄本列印完畢 〉」，避免提前截斷
        rights_pattern = r"(\*{10,}\s*(?:[土地建物]*)?他項權利部\s*\*{10,})(.*?)(?=〈\s*本謄本列印完畢\s*〉|$)"
        rights_match = re.search(rights_pattern, text, re.DOTALL)
        if rights_match:
            rights_content = rights_match.group(2).strip()
            sections["他項權利部"] = rights_content
            print(f"      ✅ 他項權利部: {len(sections['他項權利部'])} 字元")
            
            # 檢查關鍵內容
            key_terms = ["擔保債權總金額", "共同擔保地號", "共同擔保建號", "擔保債權種類及範圍"]
            for term in key_terms:
                if term in rights_content:
                    print(f"      ✅ 包含關鍵詞: {term}")
                else:
                    print(f"      ❌ 缺少關鍵詞: {term}")
        else:
            print("      ❌ 沒有找到他項權利部")

        # 🔍 檢查各章節是否還有跨頁殘留
        print("    🔍 檢查章節分割後的跨頁殘留:")
        for section_name, section_content in sections.items():
            續次頁_count = section_content.count('續次頁')
            if 續次頁_count > 0:
                print(f"      ⚠️ {section_name} 還有 {續次頁_count} 個「續次頁」")
            else:
                print(f"      ✅ {section_name} 已清理乾淨")
        
        return sections

    def extract_enhanced_rights_section(self, section_text: str, full_text: str) -> List[OrderedDict]:
        """他項權利部提取 - 修正欄位順序"""
        
        # 🔧 新增：預清理章節文本
        cleaned_section_text = self.clean_section_text_before_parsing(section_text, "他項權利部")

        rights_list = []
        
        if not cleaned_section_text or len(cleaned_section_text.strip()) < 10:
            print("  ⚠️ 他項權利部內容為空或太短")
            return rights_list
        
        print(f"🔍 謄本他項權利部解析（長度: {len(cleaned_section_text)} 字元）...")
        
        # 使用清理後的文本
        cleaned_text = cleaned_section_text
        print(f"  📝 清理後長度: {len(cleaned_text)} 字元")
        
        # 分割記錄 - 謄本可能有多筆記錄
        record_pattern = r'（(\d+)）登記次序[:：]\s*([\d\-]+)'
        records = re.findall(record_pattern, cleaned_text)
        
        if not records:
            print("  ⚠️ 沒有找到記錄標記，使用整個章節")
            records = [('0001', '0001')]
            record_contents = [cleaned_text]
        else:
            print(f"  📋 找到 {len(records)} 個他項權利記錄")
            record_contents = []
            
            for i, (num, seq) in enumerate(records):
                # 找到當前記錄的內容
                start_pattern = f'（{num}）登記次序[:：]\\s*{re.escape(seq)}'
                start_match = re.search(start_pattern, cleaned_text)
                
                if start_match:
                    start_pos = start_match.start()
                    
                    # 找結束位置
                    if i + 1 < len(records):
                        next_num, next_seq = records[i + 1]
                        end_pattern = f'（{next_num}）登記次序[:：]\\s*{re.escape(next_seq)}'
                        end_match = re.search(end_pattern, cleaned_text)
                        end_pos = end_match.start() if end_match else len(cleaned_text)
                    else:
                        end_pos = len(cleaned_text)
                    
                    record_content = cleaned_text[start_pos:end_pos].strip()
                    record_contents.append(record_content)
                    print(f"    📄 記錄 {i+1} 內容長度: {len(record_content)} 字元")
                else:
                    record_contents.append(cleaned_text)
        
        # 處理每個記錄
        for i, ((num, seq), content) in enumerate(zip(records, record_contents)):
            print(f"\n  🔄 處理記錄 {i+1}: 編號={num}, 次序={seq}")
            
            info = OrderedDict()
            info["記錄編號"] = num
            info["登記次序"] = seq
            
            # 🔧 修正：按正確順序提取欄位
            fields_order = [
                "權利種類", "收件年期", "字號", "他項_登記日期", "他項_登記原因",
                "權利人", "統一編號_他項", "住址_他項",  # 🔧 統一編號在權利人和住址之間
                "債權額比例", "擔保債權總金額",
                "存續期間", "擔保債權種類及範圍", "擔保債權確定期日", "清償日期",
                "利息(率)", "遲延利息(率)", "違約金", "其他擔保範圍約定",
                "債務人及債務額比例", "權利標的", "標的登記次序", "設定權利範圍",
                "證明書字號", "設定義務人"
            ]
            
            # 🔧 特殊處理：縮排多行欄位和需要精確邊界的欄位
            special_multiline_fields = ["擔保債權種類及範圍"]
            special_boundary_fields = ["其他擔保範圍約定"]
            
            for field in fields_order:
                if field in self.patterns:
                    # 🔧 對特殊的多行縮排欄位使用專門的提取方法
                    if field in special_multiline_fields:
                        if field == "擔保債權種類及範圍":
                            extracted_value = self.extract_multiline_indented_field(content, "擔保債權種類及範圍", "擔保債權確定期日")
                        else:
                            extracted_value = self.extract_multiline_indented_field(content, field)
                        
                        if extracted_value:
                            info[field] = extracted_value
                            print(f"    ✅ {field} (多行縮排): {extracted_value[:100]}...")
                        else:
                            print(f"    ❌ 未找到 {field} (多行縮排)")
                        continue
                    
                    # 🔧 對需要精確邊界的欄位使用專門方法
                    elif field in special_boundary_fields:
                        if field == "其他擔保範圍約定":
                            extracted_value = self.extract_precise_collateral_terms(content)
                        else:
                            extracted_value = ""
                        
                        if extracted_value:
                            info[field] = extracted_value
                            print(f"    ✅ {field} (精確邊界): {extracted_value[:100]}...")
                        else:
                            print(f"    ❌ 未找到 {field} (精確邊界)")
                        continue
                    
                    # 一般欄位的正則匹配
                    match = re.search(self.patterns[field], content, re.MULTILINE | re.DOTALL)
                    if match:
                        clean_name = field.replace("他項_", "")
                        if clean_name == "統一編號_他項":
                            clean_name = "統一編號"
                        elif clean_name == "住址_他項":
                            clean_name = "住址"
                        
                        extracted_value = self.clean_value(match.group(1))
                        info[clean_name] = extracted_value
                        
                        if field in ["擔保債權確定期日", "清償日期", "利息(率)"]:
                            print(f"    ✅ {clean_name}: {extracted_value[:100]}...")
                        else:
                            print(f"    ✅ {clean_name}: {extracted_value[:50]}...")
                    else:
                        clean_name = field.replace("他項_", "")
                        if clean_name == "統一編號_他項":
                            clean_name = "統一編號"
                        elif clean_name == "住址_他項":
                            clean_name = "住址"
                        
                        # 🔧 調試：對重要欄位進行額外檢查
                        if field in ["擔保債權確定期日", "清償日期", "利息(率)"]:
                            # 檢查是否存在欄位名稱但正則匹配失敗
                            field_name_check = field.replace("他項_", "").replace("_他項", "")
                            if field_name_check in content:
                                print(f"    ⚠️ {clean_name}: 欄位名稱存在但正則匹配失敗")
                                # 顯示找到的原始文字
                                simple_pattern = f"{field_name_check}[:：][^\\n]*"
                                simple_match = re.search(simple_pattern, content)
                                if simple_match:
                                    print(f"       原始文字: {simple_match.group()[:100]}...")
                            else:
                                print(f"    ❌ 未找到 {clean_name}")
                        else:
                            print(f"    ❌ 未找到 {clean_name}")
            
            # 提取共同擔保資訊
            self.extract_collateral_info(content, info)
            
            # 其他登記事項（使用改進版提取方法）
            other_items = self.extract_other_items_comprehensive(content, "他項權利部")
            if other_items:
                if isinstance(other_items, dict):
                    info.update(other_items)
                else:
                    info["其他登記事項"] = other_items
            
            # 🧩 住址補字（權利人住址）：優先在此精準補字，避免最終補字錯位
            try:
                _glyphs_local = globals().get('inline_glyphs', [])
                if _glyphs_local and isinstance(info.get("住址", ""), str) and info.get("住址", "").strip():
                    _before_addr = info["住址"]
                    info["住址"] = patch_inline_glyphs_in_text(info["住址"], _glyphs_local)
                    if info["住址"] != _before_addr:
                        print(f"    🧩 住址補字: {_before_addr} → {info['住址']}")
            except Exception as _e:
                print(f"    ⚠️ 住址補字失敗: {_e}")

            rights_list.append(info)
            print(f"    📊 記錄 {i+1} 完成，共 {len(info)} 個欄位")
            
        return rights_list

    def extract_other_items(self, text: str, section_name: str) -> Union[str, Dict[str, str]]:
        """提取其他登記事項 - 保持向後相容性"""
        return self.extract_other_items_comprehensive(text, section_name)

    def extract_multiline_indented_field(self, content: str, field_name: str, next_field_name: str = None) -> str:
        """專門提取包含縮排多行的欄位內容 - 加入跨頁去重"""

        # # # 🔍 診斷字元編碼問題
        # if field_name == "擔保債權種類及範圍":
        #     print(f"🔍 診斷 {field_name} 的字元編碼...")
            
        #     # 找出包含"前鎮區"的部分
        #     if "前鎮區" in content:
        #         idx = content.find("前鎮區")
        #         # 取出前10個字元來分析
        #         sample = content[max(0, idx-10):idx+5]
        #         print(f"   樣本文字: '{sample}'")
                
        #         # 顯示每個字元的Unicode編碼
        #         for i, char in enumerate(sample):
        #             print(f"   位置{i}: '{char}' = U+{ord(char):04X} (ASCII: {ord(char)})")
                    
        #         # 特別檢查"前"字前面的字元
        #         if idx > 0:
        #             prev_char = content[idx-1]
        #             print(f"   ⚠️ '前'字前面的字元: '{prev_char}' = U+{ord(prev_char):04X}")
                    
        #             if idx > 1:
        #                 prev_prev_char = content[idx-2]
        #                 print(f"   ⚠️ '前'字前兩個字元: '{prev_prev_char}' = U+{ord(prev_prev_char):04X}")
            
        #     # 嘗試各種可能的清理
        #     original = content
            
        #     # 清理各種不可見字元
        #     content = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', content)  # 控制字元
        #     content = re.sub(r'[\u200b-\u200f]', '', content)  # 零寬度字元
        #     content = re.sub(r'[\ufeff]', '', content)  # BOM
        #     content = re.sub(r'[\xa0]', ' ', content)  # 不間斷空格
            
        #     # 清理 G 和各種空格組合
        #     content = re.sub(r'[GＧ][\s\u3000\xa0\t]+', '', content)  # G + 各種空格
        #     content = re.sub(r'[GＧ]\s*前鎮區', '前鎮區', content)
            
        #     if content != original:
        #         print(f"   ✅ 已清理特殊字元")

        # 尋找欄位開始位置
        start_pattern = f"{field_name}[:：]\\s*"
        start_match = re.search(start_pattern, content)
        
        if not start_match:
            return ""
        
        # 從欄位開始位置後取內容
        start_pos = start_match.end()
        
        # 找結束位置（下一個欄位或內容結束）
        if next_field_name:
            end_pattern = f"\\s*{next_field_name}[:：]"
            end_match = re.search(end_pattern, content[start_pos:])
            end_pos = start_pos + end_match.start() if end_match else len(content)
        else:
            end_pos = len(content)
        
        # 提取範圍內的文字
        raw_content = content[start_pos:end_pos]
        
        # 🔧 關鍵改進：處理跨頁重複
        lines = raw_content.split('\n')
        content_lines = []
        seen_content = set()  # 用於去重
        
        for line in lines:
            line = line.strip()
            if line:
                # 移除可能的星號和續次頁標記
                line = re.sub(r'\*+', '', line)
                line = re.sub(r'[（(]續次頁[）)]', '', line)
                
                if line:
                    # 🔧 檢查是否為重複內容
                    # 簡化內容用於比對（移除單字母前綴、空格等）
                    simplified = re.sub(r'^[A-Z]\s+', '', line)
                    simplified = re.sub(r'\s+', '', simplified)  # 移除所有空格
                    
                    # 檢查是否已經出現過類似內容
                    is_duplicate = False
                    for seen in seen_content:
                        # 如果新內容是舊內容的子字串，或舊內容是新內容的子字串，視為重複
                        if len(simplified) > 10 and len(seen) > 10:  # 只對較長的內容做比對
                            if simplified in seen or seen in simplified:
                                is_duplicate = True
                                print(f"      ⏭️ 跳過跨頁重複: {line}")
                                break
                    
                    if not is_duplicate and line:
                        content_lines.append(line)
                        seen_content.add(simplified)
            
            # 如果遇到新的欄位（包含冒號），停止
            elif re.search(r'[\u4e00-\u9faf]+[:：]', line):
                break
        
        # 合併結果
        result = ' '.join(content_lines) if content_lines else ""
        
        

        # 🔧 額外的去重處理：針對「G 前鎮區愛群段」這種特定重複
        if result:
            # 分割成詞組並去重
            parts = result.split()
            unique_parts = []
            seen_parts = []
            
            for part in parts:
                # 簡化比對（移除單字母前綴如 G, F 等）
                simplified_part = re.sub(r'^[A-Z]\s*', '', part)
                
                # 特別處理地號格式（如：前鎮區愛群段 2904-0000地號）
                if re.search(r'[區鄉鎮市].*?段.*?\d+-\d+[地建]號', simplified_part):
                    # 提取地號核心部分
                    core_match = re.search(r'(.*?段.*?\d+-\d+[地建]號)', simplified_part)
                    if core_match:
                        core_part = core_match.group(1)
                        if core_part not in seen_parts:
                            unique_parts.append(part)
                            seen_parts.append(core_part)
                        else:
                            print(f"      🗑️ 移除重複地號: {part}")
                    else:
                        unique_parts.append(part)
                        seen_parts.append(simplified_part)
                else:
                    # 一般內容的去重
                    if simplified_part not in seen_parts:
                        unique_parts.append(part)
                        seen_parts.append(simplified_part)
                    else:
                        print(f"      🗑️ 移除重複片段: {part}")
            
            result = ' '.join(unique_parts)
        
        return result

    def extract_precise_collateral_terms(self, content: str) -> str:
        """精確提取其他擔保範圍約定，避免內容污染"""
        
        # 尋找"其他擔保範圍約定："的開始位置
        start_pattern = r"其他擔保範圍約定[:：]\s*"
        start_match = re.search(start_pattern, content)
        
        if not start_match:
            return ""
        
        start_pos = start_match.end()
        
        # 定義明確的停止條件 - 遇到這些欄位就停止
        stop_patterns = [
            r"\s*債務人及債務額比例[:：]",
            r"\s*權利標的[:：]", 
            r"\s*標的登記次序[:：]",
            r"\s*設定權利範圍[:：]",
            r"\s*證明書字號[:：]",
            r"\s*設定義務人[:：]",
            r"\s*共同擔保地號[:：]",
            r"\s*共同擔保建號[:：]",
            r"\s*其他登記事項[:：]"
        ]
        
        # 在開始位置後面的文本中尋找最近的停止點
        remaining_text = content[start_pos:]
        min_end_pos = len(remaining_text)
        
        for stop_pattern in stop_patterns:
            stop_match = re.search(stop_pattern, remaining_text)
            if stop_match:
                min_end_pos = min(min_end_pos, stop_match.start())
        
        # 提取範圍內的內容
        extracted_content = remaining_text[:min_end_pos].strip()
        
        # 清理內容
        lines = extracted_content.split('\n')
        clean_lines = []
        
        for line in lines:
            line = line.strip()
            if line:
                # 移除星號
                line = re.sub(r'\*+', '', line)
                if line:
                    clean_lines.append(line)
        
        result = ' '.join(clean_lines) if clean_lines else ""
        
        if result:
            print(f"      ✅ 精確提取其他擔保範圍約定: {result[:100]}...")
        
        return result

    def extract_collateral_info(self, content: str, info: OrderedDict):
        """提取共同擔保地號和建號資訊 """
        print("      🏢 提取共同擔保資訊...")
        
        # 共同擔保地號（可能跨多行）
        land_pattern = r"共同擔保地號[:：]\s*(.*?)(?=\s*共同擔保建號|其他登記事項|（\d+）|$)"
        land_match = re.search(land_pattern, content, re.DOTALL)
        if land_match:
            land_content = land_match.group(1).strip()
            # 清理並合併多行
            land_content = re.sub(r'\s+', ' ', land_content)
            land_content = re.sub(r'\*+', '', land_content)  # 移除星號
            info["共同擔保地號"] = land_content
            print(f"        ✅ 共同擔保地號: {land_content[:100]}...")
        
        # 共同擔保建號（可能跨多行）
        building_pattern = r"共同擔保建號[:：]\s*(.*?)(?=\s*其他登記事項|（\d+）登記次序|$)"
        building_match = re.search(building_pattern, content, re.DOTALL)
        if building_match:
            building_content = building_match.group(1).strip()
            # 清理並合併多行
            building_content = re.sub(r'\s+', ' ', building_content)
            building_content = re.sub(r'\*+', '', building_content)  # 移除星號
            info["共同擔保建號"] = building_content
            print(f"        ✅ 共同擔保建號: {building_content[:100]}...")

    def extract_other_items_comprehensive(self, text: str, section_name: str) -> Union[str, Dict[str, str]]:
        """全面提取各部的其他登記事項 - 修正版：支援複雜多行縮排格式"""
        
        if not text or not text.strip():
            print(f"        ⚠️ {section_name}文本為空")
            return ""
        
        # 🔧 更精確的排除模式，排除謄本說明文字
        exclusion_patterns = [
            # 謄本結尾說明
            r'（空白）\s*本謄本僅係.*?所有權.*?節本.*?詳細權利狀態請參閱全部謄本',
            r'本謄本僅係.*?節本.*?詳細權利狀態請參閱全部謄本',
            r'本謄本係.*?依.*?申請提供',
            r'本謄本列印完畢',
            r'〈.*?本謄本.*?列印完畢.*?〉',
            r'※注意.*本謄本',
            r'詳細權利狀態請參閱全部謄本',
            r'本謄本依據.*?提供',
            r'申請人.*?自行列印',
            # 🔧 新增：特定的謄本說明模式
            r'本謄本未申請列印.*?詳細.*?以登記機關登記為主',
            r'本謄本未申請列印',
            r'以登記機關登記為主',
        ]
        
        other_items_list = []
        
        # 🔧 改進：使用更靈活的方式找到"其他登記事項"
        lines = text.split('\n')
        
        for i, line in enumerate(lines):
            # 檢查是否包含"其他登記事項："
            if re.search(r'其他登記事項[:：]', line):
                
                # 提取冒號後的內容（第一行）
                match = re.search(r'其他登記事項[:：]\s*(.*)', line)
                first_line_content = match.group(1).strip() if match else ""
                
                # 收集所有相關內容行
                content_lines = []
                if first_line_content:
                    content_lines.append(first_line_content)
                
                # 🔧 關鍵改進：收集後續的縮排行和相關行
                j = i + 1
                while j < len(lines):
                    next_line = lines[j]
                    stripped_next = next_line.strip()
                    
                    # 如果是空行，跳過但繼續檢查
                    if not stripped_next:
                        j += 1
                        continue
                    
                    # 🔧 關鍵修正：檢查是否為謄本結束標記
                    is_transcript_end = False
                    for exclusion_pattern in exclusion_patterns:
                        if re.search(exclusion_pattern, stripped_next, re.IGNORECASE):
                            is_transcript_end = True
                            print(f"          🚫 遇到謄本結尾說明，停止: {stripped_next}")
                            break
                    
                    if is_transcript_end:
                        break
                    
                    # 檢查是否為謄本結束的其他標記
                    if (re.search(r'本謄本.*?列印完畢', stripped_next) or 
                        re.search(r'〈.*?本謄本.*?〉', stripped_next) or
                        stripped_next.startswith('※注意') or
                        stripped_next.startswith('※')):
                        print(f"          🚫 遇到謄本結束標記，停止: {stripped_next}")
                        break
                    
                    # 🔧 關鍵改進：更靈活的縮排和續行判斷
                    is_continuation = False
                    
                    # 🔧 修正：定義哪些是真正的新欄位（會導致停止）
                    # 只有這些明確的欄位標題才會停止收集
                    real_field_patterns = [
                        r'^標示部',
                        r'^所有權部', 
                        r'^他項權利部',
                        r'^\*{5,}',  # 五個以上星號開頭
                        r'^登記日期[:：]',
                        r'^登記原因[:：]',
                        r'^面積[:：]',
                        r'^使用分區[:：]',
                        r'^使用地類別[:：]',
                        r'^地上建物建號[:：]',
                        r'^所有權人[:：]',
                        r'^權利範圍[:：]',
                        r'^建物門牌[:：]',
                        r'^主要用途[:：]',
                    ]
                    
                    # 檢查是否為真正的新欄位
                    is_new_field = False
                    for field_pattern in real_field_patterns:
                        if re.search(field_pattern, stripped_next):
                            is_new_field = True
                            print(f"          🛑 遇到新欄位，停止: {stripped_next}")
                            break
                    
                    if is_new_field:
                        break
                    
                    # 🔧 關鍵修正：以下這些都應該是其他登記事項的內容
                    # 包括「合併自：」、「因分割增加地號：」等
                    continuation_keywords = [
                        '合併自：',
                        '合併自',
                        '因分割增加地號：',
                        '因分割增加地號',
                        '分割自：',
                        '分割自',
                        '重測前：',
                        '重測前',
                        '原地號：',
                        '原地號',
                        # 🔧 新增：信託相關關鍵字
                        '委託人：',
                        '受託人：',
                        '信託目的：',
                        '信託期間：',
                        '信託財產：',
                        '信託專簿：',
                    ]
                    
                    # 如果包含這些關鍵字，視為續行內容
                    for keyword in continuation_keywords:
                        if keyword in stripped_next:
                            is_continuation = True
                            print(f"          ✅ 關鍵字行({keyword}): {stripped_next}")
                            break
                    
                    if not is_continuation:
                        # 1. 明顯的縮排行（開頭有空格或Tab）
                        if (next_line.startswith('                  ') or  # 18個空格縮排
                            next_line.startswith('        ') or            # 8個空格縮排
                            next_line.startswith('    ') or                # 4個空格縮排
                            next_line.startswith('\t')):                   # Tab縮排
                            is_continuation = True
                            print(f"          ✅ 縮排行: {stripped_next}")
                        
                        # 2. 括號開頭的行（如"（一般註記事項）"）
                        elif (stripped_next.startswith('（') or stripped_next.startswith('(')):
                            is_continuation = True
                            print(f"          ✅ 括號行: {stripped_next}")
                        
                        # 3. 數字地號的續行（如：１６、－１７、－１９...）
                        elif re.match(r'^[－\-]?\d+[、，]', stripped_next) or re.match(r'^[０-９－、]+', stripped_next):
                            is_continuation = True
                            print(f"          ✅ 數字續行: {stripped_next}")
                        
                        # 4. 不包含冒號的行（一般續行內容）
                        # 🔧 修正：排除跨頁雜訊（如 "T", "4T", "2E", "12", "FF" 以及地號建號行）
                        elif '：' not in stripped_next and ':' not in stripped_next:
                            # 排除跨頁雜訊
                            pagination_noise = ['T', '4T', '2E', '12', 'FF', 'EG', 'F0', 'RE', 'CB', '76', '40', '58', '31', '4D', '64', '2B', '14', '48', '11', '7C', '1R', '2F', '3A', '4B', '5C', '6D', '87']
                            if stripped_next in pagination_noise:
                                print(f"          🚫 跨頁雜訊，跳過: {stripped_next}")
                                j += 1
                                continue
                            # 排除地號/建號行（跨頁標題）
                            if re.search(r'[區鄉鎮市].*?段\s+\d{4,5}-\d{3,4}[地建]號', stripped_next):
                                print(f"          🚫 跨頁地號建號行，停止: {stripped_next}")
                                break
                            is_continuation = True
                            print(f"          ✅ 續行內容: {stripped_next}")
                        
                        # 5. 🔧 修正：檢查是否為謄本說明行（需要排除）
                        elif any(keyword in stripped_next for keyword in ['本謄本未申請列印', '本謄本', '詳細', '以登記機關登記為主']):
                            # 這些是謄本說明文字，不是其他登記事項的內容
                            print(f"          🚫 謄本說明行，停止: {stripped_next}")
                            break
                    
                    if is_continuation:
                        content_lines.append(stripped_next)
                        j += 1
                    else:
                        # 其他情況，停止收集
                        print(f"          🛑 遇到不明內容，停止: {stripped_next}")
                        break
                
                # 處理收集到的內容
                if content_lines:
                    # 🔧 改進：保持適當的格式，支援多行結構
                    full_content = ""
                    
                    # 檢查是否有多行結構
                    if len(content_lines) > 1:
                        # 多行內容：用空格連接成一行
                        full_content = ' '.join(content_lines)
                        print(f"          📋 多行內容（已合併）:")
                        print(f"            {full_content}")
                    else:
                        # 單行內容
                        full_content = content_lines[0]
                        print(f"          📋 單行內容: {full_content}")
                    
                    # 🔧 最終檢查：確保不是謄本說明
                    is_final_exclusion = False
                    for exclusion_pattern in exclusion_patterns:
                        if re.search(exclusion_pattern, full_content, re.IGNORECASE):
                            is_final_exclusion = True
                            print(f"          🚫 最終排除: {exclusion_pattern}")
                            break
                    
                    if not is_final_exclusion:
                        # 移除多餘的星號
                        full_content = re.sub(r'\*+', '', full_content).strip()
                        
                        if full_content and len(full_content.strip()) > 0:
                            # 🔧 最後檢查：排除包含多個謄本關鍵詞的完整句子
                            transcript_keywords = ['本謄本', '節本', '詳細權利狀態', '參閱全部謄本']
                            keyword_count = sum(1 for keyword in transcript_keywords if keyword in full_content)
                            
                            if keyword_count >= 2:
                                print(f"          🚫 包含多個謄本關鍵詞，可能是說明文字: {full_content}")
                            else:
                                other_items_list.append(full_content)
                                print(f"          ✅ 有效其他登記事項: {full_content}")
                        else:
                            print(f"          ⚠️ 內容為空，跳過")
                    else:
                        print(f"          🚫 排除謄本說明內容: {full_content[:50]}...")
        
        all_matches = other_items_list
        
        if not all_matches:
            print(f"        ❌ {section_name}沒有找到有效的其他登記事項")
            return ""
        
        print(f"        ✅ {section_name}找到 {len(all_matches)} 個有效其他登記事項")
        
        # 清理和格式化
        valid_items = []
        for match in all_matches:
            # 移除多餘的星號
            cleaned = re.sub(r'\*+', '', match)
            cleaned = cleaned.strip()
            
            if cleaned:
                valid_items.append(cleaned)
                print(f"          ✅ 最終有效項目: {cleaned}")
        
        if not valid_items:
            print(f"        ❌ 所有項目都被排除，沒有有效的其他登記事項")
            return ""
        elif len(valid_items) == 1:
            return valid_items[0]
        else:
            # 多個其他登記事項
            result = {}
            for i, item in enumerate(valid_items):
                if i == 0:
                    result["其他登記事項"] = item
                else:
                    result[f"其他登記事項{i}"] = item
            return result

    def clean_section_text_before_parsing(self, section_text: str, section_name: str) -> str:
        """在解析前清理章節文本，移除跨頁重複內容"""
        
        print(f"    🧹 預清理{section_name}章節文本...")
        
        # 🔧 新增：先處理明顯的跨頁重複模式
        lines = section_text.split('\n')
        cleaned_lines = []
        
        for i, line in enumerate(lines):
            line_stripped = line.strip()
            
            # 檢查是否為續次頁標記
            if '續次頁' in line_stripped:
                print(f"      🔍 發現續次頁標記（第{i+1}行）")
                # 檢查前後內容是否重複
                if i > 0 and i < len(lines) - 1:
                    prev_line = lines[i-1].strip()
                    next_line = lines[i+1].strip() if i+1 < len(lines) else ""
                    
                    # 如果前後內容相似，跳過重複部分
                    if prev_line and next_line:
                        prev_simplified = re.sub(r'\s+', '', prev_line)
                        next_simplified = re.sub(r'\s+', '', next_line)
                        
                        if prev_simplified in next_simplified or next_simplified in prev_simplified:
                            print(f"      🗑️ 跳過跨頁重複內容")
                            continue
            
            # 移除續次頁標記
            line_cleaned = re.sub(r'[（(]續次頁[）)]', '', line_stripped)
            
            # 如果不是空行，加入結果
            if line_cleaned:
                cleaned_lines.append(line_cleaned)
        
        cleaned_text = '\n'.join(cleaned_lines)
        
        # 定義謄本結尾的切割點
        end_markers = [
            r'本謄本僅係.*?節本.*?詳細權利狀態請參閱全部謄本',
            r'本謄本係.*?依.*?申請提供',
            r'本謄本列印完畢',
            r'〈.*?本謄本.*?列印完畢.*?〉',
            r'※注意事項',
            r'（空白）\s*本謄本僅係\s*所有權個人全部\s*節本',
        ]
        
        for marker in end_markers:
            # 找到第一個匹配的結束標記，從該處截斷
            match = re.search(marker, cleaned_text, re.IGNORECASE | re.DOTALL)
            if match:
                # 從匹配位置截斷
                cleaned_text = cleaned_text[:match.start()].strip()
                print(f"      🔧 在「{marker[:20]}...」處截斷文本")
                break
        
        if len(cleaned_text) != len(section_text):
            print(f"      ✅ 預清理完成，文本長度: {len(section_text)} → {len(cleaned_text)}")
        else:
            print(f"      ✅ 預清理完成，無需截斷")
        
        return cleaned_text

    def fullwidth_to_halfwidth(self, text: str) -> str:
        """
        將全形數字和標點符號轉換為半形
        專門用於處理如「３９．９２」→「39.92」的轉換
        """
        if not text:
            return text

        # 全形→半形對應表
        fullwidth_map = {
            # 全形數字 (U+FF10 ~ U+FF19)
            '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
            '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
            # 全形標點
            '．': '.', '，': ',', '：': ':', '（': '(', '）': ')',
        }

        result = text
        for fullwidth, halfwidth in fullwidth_map.items():
            result = result.replace(fullwidth, halfwidth)

        return result

    # 修正問題 1：clean_value 函數
    def clean_value(self, value: str) -> str:
        """改進版清理值 - 智慧保留統一編號的*號"""
        if not value:
            return ""
        
        # 🔧 關鍵修正：檢查是否為統一編號格式
        stripped_value = value.strip()
        
        # 檢查統一編號模式（修正正則表達式）
        is_unified_number = False
        if (re.match(r'^[A-Z]\d\*+\d+', stripped_value) or      # 如：A1*******9
            re.match(r'^\d{8}', stripped_value) or               # 如：12345678
            re.match(r'^[A-Z0-9]\d[\*\d]{6,8}\d?', stripped_value)):  # 各種統一編號格式
            is_unified_number = True
            print(f"      🔍 檢測到統一編號格式，保留*號: {stripped_value}")
        
        # 第1步：移除*號（但統一編號例外）
        if is_unified_number:
            # 統一編號：保留*號，只清理其他
            cleaned = value
            print(f"      🛡️ 統一編號保護: {value}")
        else:
            # 其他欄位：移除*號
            cleaned = re.sub(r'\*+', '', value)
        
        # 第2步：清理跨頁標記
        cleaned = re.sub(r'[（(]續次頁[）)]\s*', '', cleaned)
        
        # 第3步：處理多行內容
        lines = cleaned.split('\n')
        cleaned_lines = []
        
        for line in lines:
            line = line.strip()
            if line:
                # 再次檢查跨頁標記
                line = re.sub(r'[（(]續次頁[）)]\s*', '', line)
                if line:
                    cleaned_lines.append(line)
        
        # 第4步：重新組合內容
        if len(cleaned_lines) > 1:
            # 對於包含編號列表的內容，保留換行
            if any(re.match(r'^\d+[．.]', line) for line in cleaned_lines):
                cleaned = '\n'.join(cleaned_lines)
            else:
                # 其他多行內容用空格連接
                cleaned = ' '.join(cleaned_lines)
        else:
            cleaned = ' '.join(cleaned_lines) if cleaned_lines else ""
        
        # 第5步：特殊處理
        cleaned = cleaned.replace('元／平方公尺', '元/平方公尺')

        result = cleaned.strip()
        
        # 🔧 新增：OCR誤識別修正（在return前）
        original_result = result
        # print(f"      🔍 OCR修正前: '{original_result}'")  # 新增這行
        result = apply_ocr_corrections(result)
        # print(f"      🔍 OCR修正後: '{result}'")  # 新增這行

        # 🔧 調試：顯示清理過程
        if value != original_result:
            if is_unified_number:
                print(f"      🛡️ 統一編號清理: '{value}' → '{original_result}'（保留*號）")
            else:
                pass
        
        # 🔧 調試：顯示OCR修正過程
        if original_result != result:
            print(f"      🔧 OCR修正: '{original_result}' → '{result}'")
        
        return result

    def apply_ocr_corrections(self, text: str) -> str:
        """OCR常見誤識別修正"""
        if not text:
            return text
        
        result = text
        
        # 🔧 特殊處理：單獨的"爐房"
        if result.strip() == '爐房':
            return '煱爐房'
        
        # OCR誤識別修正規則
        ocr_corrections = [
            ('大爐房', '煱爐房'),
            ('大爐', '煱爐'),
            ('鳳山品', '鳳山區'),
            ('左警', '左營'),
            (' 後段','廍後段'),
            ('廊','廍'),
            ('燕巢品', '燕巢區'),
            ('苓稚', '苓雅'),
            ('术', '術'),
        ]
        
        # 🔧 關鍵：找到一個匹配就停止
        for wrong_text, correct_text in ocr_corrections:
            if wrong_text in result:
                result = result.replace(wrong_text, correct_text)
                print(f"      🔧 OCR修正: {wrong_text} → {correct_text}")
                break  # 重要：停止處理
        
        return result
        
    # -------- 通用：抓取「label + 縮排續行」欄位 --------
    def get_multiline_indented(self, text: str, field: str,
                               stop_labels: List[str] = None) -> str:
        
        """
        取像『地上建物建號』這類：第一行有 label，後續行用縮排續行的欄位。
        ‑ 會自動濾掉跨頁雜訊（十六進位或頁眉行）
        ‑ 最後把 3 個以上連續空白壓成 2 個空白
        """
        if stop_labels is None:
            stop_labels = ["其他登記事項", "建物門牌", "主要用途", "主要建材",
                           "層數", "總面積", "層次", "附屬建物用途"]

        m = re.search(rf"{field}[:：]\s*(.*)", text)
        if not m:
            return ""

        lines = text.split("\n")
        start_idx = text[:m.start()].count("\n")
        collected = []

        first_part = m.group(1).strip()
        if first_part and first_part not in ("（空白）", "(空白)"):
            collected.append(first_part)

        i = start_idx + 1
        while i < len(lines):
            raw = lines[i]
            txt = raw.strip()
            if not txt:
                i += 1
                continue

            # ---------- 停止條件 ----------
            if (any(lbl in txt for lbl in stop_labels) or
                    re.match(r'^[\u4e00-\u9fff]{2,}[:：]', txt) or
                    txt.startswith('*****') or
                    '本謄本' in txt):
                break

            # ---------- 濾掉跨頁雜訊 ----------
            # ① 十六進位/頁眉行 (如 1R C0 70…)
            if re.match(r'^[0-9A-F]{1,2}[A-F0-9]?\s', txt):
                i += 1
                continue
            # ② 「段  xxx 地號」頁腳行
            if re.search(r'段\s+\d{3,5}[-之]\d{2,4}地號', txt) and '-' not in txt.strip().split()[0]:
                i += 1
                continue

            collected.append(txt)
            i += 1

        # ------- 合併、去星號、壓縮多餘空白 -------
        result = re.sub(r"\*+", "", " ".join(collected)).strip()
        # 把 3 個以上空白壓成 2 個
        result = re.sub(r'\s{3,}', '  ', result)

        return result


    def extract_basic_info(self, text: str) -> OrderedDict:
        """提取基本資訊"""
        basic_info = OrderedDict()
        
        # 欄位順序
        ordered_fields = ["謄本類型", "地號建號", "列印時間"]
        
        for field in ordered_fields:
            if field in self.patterns:
                match = re.search(self.patterns[field], text, re.MULTILINE)
                if match:
                    basic_info[field] = self.clean_value(match.group(1))
                else:
                    print(f"    ❌ 未找到 {field}")
        
        # 檢測謄本類別
        transcript_class = self.detect_transcript_class(text)
        basic_info["謄本類別"] = transcript_class
        
        return basic_info

    def extract_structured_data(self, transcript_text: str, ocr_data: Dict[str, str]) -> Dict[str, Any]:
        """結構化資料提取主函數"""
        
        try:
            # 步驟1：預處理（使用方法1：簡單字符串替換）
            processed_text = transcript_text  # 直接使用，因為已經清理過了

            # 🔍 最終跨頁檢查
            final_check = processed_text.count('續次頁')
            if final_check > 0:
                print(f"    ⚠️ 警告：預處理後仍有 {final_check} 個「續次頁」未清理")
            else:
                print(f"    ✅ 預處理成功：所有跨頁內容已清理")
            
            # 步驟2：章節分割
            sections = self.split_sections_enhanced(processed_text)
            
            # 步驟3：類型檢測
            doc_type = self.detect_document_type(processed_text)
            transcript_class = self.detect_transcript_class(processed_text)
            
            # 步驟4：基本資訊
            basic_info = self.extract_basic_info(processed_text)
            print(f"    📅 列印時間: {basic_info.get('列印時間', '未找到')}")
            
            # 步驟5：標示部
            indicator_data = self.extract_indicator_section(sections.get("標示部", ""), doc_type)
            
            # 步驟6：所有權部
            ownership_data = self.extract_ownership_section(sections.get("所有權部", ""), ocr_data, doc_type)
            
            # 步驟7：他項權利部
            rights_data = self.extract_enhanced_rights_section(sections.get("他項權利部", ""), processed_text)
            
            # 組合結果
            structured = OrderedDict([
                ("謄本類型", f"{doc_type} - {transcript_class}"),
                ("基本資訊", basic_info),
                ("標示部", indicator_data),
                ("所有權部", ownership_data),
                ("他項權利部", rights_data)
            ])
            
            return structured
            
        except Exception as e:
            print(f"❌ 謄本解析失敗: {e}")
            raise e

    def clean_single_transcript_cross_page(self, transcript_text: str) -> str:
        """清理單份謄本的跨頁重複資料"""
        
        # 🔍 檢查原始跨頁內容
        cross_page_lines = []
        lines = transcript_text.split('\n')
        
        for i, line in enumerate(lines):
            line_stripped = line.strip()
            # 檢查跨頁模式（建物謄本 + 土地謄本 + 孤立行）
            if ('續次頁' in line_stripped or 
                ('地政事務所' in line_stripped and '列印時間' in line_stripped) or
                # 建物謄本跨頁
                (re.search(r'^[0-9A-F]{1,4}\s+[0-9A-F]{1,4}\s+.*?段.*?建號.*?列印時間', line_stripped)) or
                (re.search(r'^.*?段.*?建號\s+列印時間：.*?頁次：\d+', line_stripped)) or
                # 土地謄本跨頁
                (re.search(r'^[0-9A-F]{1,4}\s+[0-9A-F]{1,4}\s+.*?段.*?地號.*?列印時間', line_stripped)) or
                (re.search(r'^.*?段.*?地號\s+列印時間：.*?頁次：\d+', line_stripped)) or
                # 單獨的十六進制字母
                (re.search(r'^\s*[A-F]{1,4}\s*', line_stripped)) or
                # 十六進制代碼開頭的行
                (re.search(r'^\s*[0-9A-F]{1,4}\s+', line_stripped)) or
                # 🔧 新增：孤立的地段建號行
                (re.search(r'^\s*.*?[區鄉鎮市].*?段.*?小段\s+\d{5}-\d{3}[建地]號\s*', line_stripped)) or
                (re.search(r'^\s*.*?[區鄉鎮市].*?段.*?小段\s+\d{4}-\d{4}[建地]號\s*', line_stripped))):
                cross_page_lines.append((i+1, line_stripped))

        # 🔧 保護共同擔保相關內容，防止被跨頁清理誤刪
        protected_content = {}
        cleaned_text = transcript_text
        
        if '共同擔保地號' in cleaned_text or '共同擔保建號' in cleaned_text:
            lines = cleaned_text.split('\n')
            i = 0
            while i < len(lines):
                if re.search(r'^.*共同擔保[地建]號[:：]', lines[i]):
                    # 這是真正的共同擔保標題行，需要保護
                    # 保護當前行和後續相關行的邏輯...
                    protected_lines = [lines[i]]
                    j = i + 1
                    while j < len(lines) and j < i + 3:
                        next_line = lines[j].strip()
                        if re.search(r'\d{4}-\d{4}', next_line) and not any(keyword in next_line for keyword in ['共同擔保建號', '其他登記事項']):
                            protected_lines.append(lines[j])
                            j += 1
                        else:
                            break
                    
                    # 保護所有相關行
                    for idx, pline in enumerate(protected_lines):
                        placeholder = f"__PROTECTED_COLLATERAL_LAND_{i}_{idx}__"
                        protected_content[placeholder] = pline
                        lines[i + idx] = placeholder
                        print(f"         🛡️ 保護共同擔保相關行: {pline}")
                    
                    i = j
                else:
                    i += 1
            cleaned_text = '\n'.join(lines)

        # 🔧 完整版清理模式（建物謄本 + 土地謄本）
        patterns_to_remove = [
            # === 通用模式 ===
            # 模式1: 續次頁開頭的行
            r'^.*?[（\(]續次頁[）\)].*?',
            
            # 模式2: 任何包含「列印時間」和「頁次」的行
            r'^.*?列印時間：.*?頁次：\d+.*?',
            
            # 模式3: 只有地政事務所的行
            r'^.*?地政事務所\s*',
            
            # 模式4: 頁次行
            r'^\s*頁次[:：]\s*\d+\s*',
            
            # === 建物謄本專用 ===
            # 模式5: 十六進制代碼 + 建號 + 列印時間
            r'^.*?[0-9A-F]{1,4}\s+[0-9A-F]{1,4}\s+.*?段.*?建號.*?列印時間.*?',
            
            # 模式6: 十六進制代碼 + 建號（沒有列印時間）
            r'^.*?[0-9A-F]{1,4}\s+[0-9A-F]{1,4}\s+.*?段.*?建號.*?',
            
            # 模式7: 直接建號 + 列印時間（沒有十六進制）
            r'^.*?段.*?小段\s+\d{5}-\d{3}建號\s+列印時間：.*?頁次：\d+.*?',
            
            # === 土地謄本專用 ===
            # 模式8: 十六進制代碼 + 地號 + 列印時間
            r'^.*?[0-9A-F]{1,4}\s+[0-9A-F]{1,4}\s+.*?段.*?地號.*?列印時間.*?',
            
            # 模式9: 十六進制代碼 + 地號（沒有列印時間）
            r'^.*?[0-9A-F]{1,4}\s+[0-9A-F]{1,4}\s+.*?段.*?地號.*?',
            
            # 模式10: 直接地號 + 列印時間（沒有十六進制）
            r'^.*?段.*?小段\s+\d{4}-\d{4}地號\s+列印時間：.*?頁次：\d+.*?',
            
            # 模式11: 只有十六進制字母的行（如單獨的 "RE", "EG"）
            r'^\s*[A-F]{1,4}\s*',
            
            # === 最寬鬆模式 ===
            # 模式12: 單獨的十六進制代碼行
            r'^\s*[0-9A-F]{1,4}\s+[0-9A-F]{1,4}\s*',
            
            # 模式13: 折行的地段地號（跨兩行）
            r'.*?[區鄉鎮市].*?\n\s*段\s+\d{4}-\d{4}地號',

            # 模式14: 折行的地段建號（跨兩行）  
            r'.*?[區鄉鎮市].*?\n\s*段\s+\d{5}-\d{3}建號',

            # 模式15: 更通用的折行地段號碼（支援地號和建號）
            r'.*?[區鄉鎮市].*?\n\s*段\s+\d{4,5}-\d{3,4}[地建]號',
        ]
        
        total_removed = 0

        # 測試跨頁內容是否會被清理
        test_content = "大寮區磚子\n段 3800-0000地號"
        print(f"🔍 測試跨頁內容是否會被清理:")
        print(f"📋 總共有 {len(patterns_to_remove)} 個清理模式")

        for i, pattern in enumerate(patterns_to_remove, 1):
            if re.search(pattern, test_content, re.MULTILINE):
                print(f"   ✅ 模式{i}匹配成功")
            elif i in [13, 14, 15, 16]:
                print(f"   ❌ 模式{i}未匹配: {pattern}")

        
        for i, pattern in enumerate(patterns_to_remove, 1):
            matches = re.findall(pattern, cleaned_text, re.MULTILINE)
            if matches:
                # 檢查是否誤傷共同擔保內容
                for match in matches:
                    if "共同擔保" in match or "磚子磘段" in match:
                        print(f"         ⚠️ 警告：模式{i}誤傷共同擔保: {match[:50]}...")
                        print(f"         📋 問題模式: {pattern}")
                
                # 顯示被清理的內容樣本
                for j, match in enumerate(matches[:2]):  # 顯示前2個
                    print(f"         - {match}")
                if len(matches) > 2:
                    print(f"         ... 還有 {len(matches) - 2} 項")
                
                cleaned_text = re.sub(pattern, "", cleaned_text, flags=re.MULTILINE)
                total_removed += len(matches)

        # 清理多餘的空行
        cleaned_text = re.sub(r'\n\s*\n', '\n', cleaned_text)
        
        # 🔧 修正版：智慧跨行清理（保護第一組資料）
        # 將文本按行分割
        lines = cleaned_text.split('\n')
        filtered_lines = []
        i = 0

        # 記錄已經出現的重要資料（用於判斷是否為重複）
        seen_office = None  # 記錄第一次出現的地政事務所
        seen_location = None  # 記錄第一次出現的地段建號
        first_office_processed = False

        while i < len(lines):
            line = lines[i].strip()
            
            # 🛡️ 檢查並保護第一組基本資料
            
            # 檢查地政事務所
            if re.search(r'.*?地政事務所\s*', line) and not first_office_processed:
                if seen_office is None:
                    seen_office = line
                    print(f"         🛡️ 保護第一組地政事務所：{line}")
                    filtered_lines.append(lines[i])
                    first_office_processed = True
                    i += 1
                    continue
            
            # 檢查地段建號
            if (re.search(r'.*?[區鄉鎮市].*?段.*?小段\s+\d{5}-\d{3}[建地]號', line) or
                re.search(r'.*?[區鄉鎮市].*?段.*?小段\s+\d{4}-\d{4}[建地]號', line)):
                if seen_location is None:
                    seen_location = line
                    filtered_lines.append(lines[i])
                    i += 1
                    continue
            
            # 🔧 關鍵修正：保護年月地價行
            # 檢查是否為年月地價行（055年06月 *******30.3元／平方公尺）
            if re.search(r'\d{3}年\d{2}月.*?元[／/]平方公尺', line):
                print(f"         ✅ 保護年月地價行：{line}")
                filtered_lines.append(lines[i])
                i += 1
                continue
            
            # 🗑️ 清理重複的跨頁資料
            
            # 檢查續次頁區塊
            if '續次頁' in line:
                print(f"         🔍 發現跨頁區塊開始：第{i+1}行 - {line}")
                
                # 尋找區塊結束
                j = i + 1
                block_lines = [line]
                found_end = False
                
                while j < len(lines) and j < i + 15:  # 增加搜尋範圍到15行
                    next_line = lines[j].strip()
                    block_lines.append(next_line)
                    
                    if ('頁次：' in next_line or re.search(r'頁次[:：]\s*\d+', next_line)):
                        found_end = True
                        break
                    j += 1
                
                if found_end:
                    # 🔧 顯示區塊內容以便調試
                    for idx, block_line in enumerate(block_lines):
                        print(f"           第{i+idx+1}行: {block_line}")
                    i = j + 1
                    continue
                else:
                    print(f"         ⚠️ 跨頁區塊沒找到結束，保留此行")
            
            # 檢查重複的地政事務所
            elif (re.search(r'.*?地政事務所\s*', line) and 
                seen_office is not None and 
                line != seen_office):
                i += 1
                continue
            
            # 檢查重複的地段建號
            elif ((re.search(r'.*?[區鄉鎮市].*?段.*?小段\s+\d{5}-\d{3}[建地]號', line) or
                re.search(r'.*?[區鄉鎮市].*?段.*?小段\s+\d{4}-\d{4}[建地]號', line)) and
                seen_location is not None):
                i += 1
                continue
            
            # 🔧 修正：更嚴格的十六進制檢查（避免誤刪年月地價）
            # 只有純十六進制（沒有其他內容）才刪除
            elif line in ['EG', 'F0', 'RE', 'CB', '76', '40', '58', '31', '4D', '64', '2B', '14', '48', '11', '7C', '1R', '2F', '3A', '4B', '5C', '6D', '87', '4T', '2E', '12', 'FF']:
                print(f"         🗑️ 清理孤立十六進制：{line}")
                i += 1
                continue
            
            # 檢查包含列印時間和頁次的行
            elif re.search(r'列印時間：.*?頁次：\d+', line):
                i += 1
                continue
            
            # 保留其他重要內容
            else:
                filtered_lines.append(lines[i])
            
            i += 1

        # 重新組合文本
        cleaned_text = '\n'.join(filtered_lines)

        # 清理多餘的空行
        cleaned_text = re.sub(r'\n\s*\n\s*\n', '\n\n', cleaned_text)

        # 🔧 最終檢查（修正版：不要刪除年月地價）
        remaining_lines = []
        hex_patterns = ['EG', 'F0', 'RE', 'CB', '76', '40', '58', '31', '4D', '64', '2B', '14', '48', '11', '7C', '1R', '2F', '3A', '4B', '5C', '6D', '87', '4T', '2E', '12', 'FF']

        for line in cleaned_text.split('\n'):
            line_stripped = line.strip()
            # 檢查跨頁標記和殘留十六進位（但不包括年月地價）
            if ('續次頁' in line_stripped or 
                (re.search(r'列印時間：.*?頁次：\d+', line_stripped)) or
                line_stripped in hex_patterns):
                remaining_lines.append(line_stripped)
        
        if remaining_lines:
            for line in remaining_lines[:2]:
                print(f"        - {line[:60]}...")
        else:
            print(f"      ✅ 跨頁清理完成，共移除 {total_removed} 項")
        
        # 🔧 清理行內的單字母十六進位殘留
        # 這些字母後面通常會接地名或其他內容
        final_lines = []
        for line in cleaned_text.split('\n'):
            # 移除明顯的十六進位單字母模式
            # G 前鎮區 → 前鎮區
            # F 本抵押權 → 本抵押權
            cleaned_line = line
            
            # 🔧 新增更多清理模式
            cleanup_patterns = [
                # 原有模式
                (r'\b([A-HG])\s+([前後左右中東西南北])', r'\2'),
                (r'\b([A-HG])\s+(本)', r'\2'),
                (r'\b([A-HG])\s+(\S*[區鄉鎮市]\S*段)', r'\2'),
                
                # 🔧 新增：清理完整的地段建號
                (r'\b[A-HG]\s+.*?[區鄉鎮市].*?段.*?\d{4}-\d{4}[地建]號', ''),
                (r'\b[A-HG]\s+.*?[區鄉鎮市].*?段(?!\d)', ''),

                # 🔧 新增：清理建號列表中的十六進制殘留
                (r'\b[0-9A-F]{1,2}[A-F]\b(?=\s+\d{5}-\d{3})', ''),  # 如 "1R 01509-000"
                (r'(\d{5}-\d{3})\s+[0-9A-F]{1,2}[A-F]\b\s+(\d{5}-\d{3})', r'\1 \2'),  # 建號之間的殘留
            ]
            
            for pattern, replacement in cleanup_patterns:
                old_line = cleaned_line
                cleaned_line = re.sub(pattern, replacement, cleaned_line)
                if cleaned_line != old_line:
                    print(f"      🧹 清理行內殘留: {old_line[:50]}... → {cleaned_line[:50]}...")
            
            final_lines.append(cleaned_line)

        cleaned_text = '\n'.join(final_lines)
        print(f"  ✅ 行內十六進位清理完成")
        
        # 🔧 恢復被保護的內容
        for placeholder, original_content in protected_content.items():
            cleaned_text = cleaned_text.replace(placeholder, original_content)
            print(f"         ✅ 恢復保護內容: {original_content[:50]}...")

        # 針對已合併的跨頁內容進行分離與清理
        if "共同擔保地號" in cleaned_text:
            # 通用清理：移除共同擔保地號/建號後面的重複地段內容
            cleaned_text = re.sub(r'(共同擔保[地建]號[:：][^，。\n]*?)\s+[^，。\n]*?[區鄉鎮市][^，。\n]*?段\s*\d{4,5}-\d{3,4}[地建]號', r'\1', cleaned_text)
            print(f"         🧹 清理合併的共同擔保地號跨頁雜訊")

        return cleaned_text


    def clean_pagination_method1(self, text: str) -> str:
        """通用跨頁清理 - 徹底清理所有跨頁重複"""
        
        # 🔍 調試：顯示清理前的跨頁內容
        lines_with_續次頁 = [line.strip() for line in text.split('\n') if '續次頁' in line]
        for i, line in enumerate(lines_with_續次頁[:5]):  # 只顯示前5行
            print(f"  {i+1}. {line}")
        if len(lines_with_續次頁) > 5:
            print(f"  ... 還有 {len(lines_with_續次頁) - 5} 行")
        
        # 🔧 新的徹底清理模式
        patterns_to_remove = [
            # 模式1：任何包含「續次頁」的完整行
            r'^.*?[（\(]續次頁[）\)].*?',
            
            # 模式2：跨頁後的地政事務所獨立行
            r'^\s*.*?地政事務所\s*',
            
            # 模式3：十六進制代碼 + 地段建號 + 列印時間的行
            r'^\s*[0-9A-F]{1,3}\s+[0-9A-F]{1,3}\s+.*?[區鄉鎮市].*?段.*?建號.*?列印時間.*?',
            
            # 模式4：只有十六進制代碼的行
            r'^\s*[0-9A-F]{1,3}\s+[0-9A-F]{1,3}\s*',
            
            # 模式5：單獨的頁次行
            r'^\s*頁次[:：]\s*\d+\s*'
        ]
        
        cleaned_text = text
        total_removed = 0
        
        for i, pattern in enumerate(patterns_to_remove, 1):
            
            # 找匹配項
            matches = re.findall(pattern, cleaned_text, re.MULTILINE)
            
            if matches:
                print(f"     ✅ 找到 {len(matches)} 個匹配")
                # 顯示前3個匹配內容
                for j, match in enumerate(matches[:3]):
                    print(f"       - {match}")
                if len(matches) > 3:
                    print(f"       ... 還有 {len(matches) - 3} 個")
                
                # 執行清理
                cleaned_text = re.sub(pattern, "", cleaned_text, flags=re.MULTILINE)
                total_removed += len(matches)
            else:
                print(f"     ❌ 沒有匹配")
        
        # 🔍 最後檢查
        remaining_續次頁 = [line.strip() for line in cleaned_text.split('\n') if '續次頁' in line]
        
        if remaining_續次頁:
            for line in remaining_續次頁[:3]:
                print(f"  - {line}")
        else:
            print("✅ 所有跨頁內容已清理完畢")
        
        print(f"📊 總共清理項目: {total_removed}")
        
        return cleaned_text

    def extract_indicator_section(self, section_text: str, doc_type: str) -> OrderedDict:
        """標示部提取 - 修正版：避免建物謄本重複處理其他登記事項"""
        data = OrderedDict()
        search_text = section_text if section_text else ""
        
        try:
            if doc_type == "建物謄本":
                # 建物標示部的實際順序
                ordered_fields = [
                    "標示_登記日期", "標示_登記原因", "建物門牌", "建物坐落地號", 
                    "主要用途", "主要建材", "層數", "總面積"
                ]
                
                # 提取基本欄位
                for field in ordered_fields:
                    if field in self.patterns:
                        try:
                            match = re.search(self.patterns[field], search_text, re.MULTILINE | re.DOTALL)
                            if match:
                                clean_name = field.replace("標示_", "")
                                raw_value = match.group(1)
                                cleaned_value = self.clean_value(raw_value)
                                
                                # 🔧 加強檢查：確保值不為空
                                if cleaned_value and cleaned_value.strip():
                                    data[clean_name] = cleaned_value
                                    print(f"      ✅ 提取 {clean_name}: {data[clean_name]}")
                                else:
                                    print(f"      ⚠️ {clean_name} 清理後為空: 原始='{raw_value}' 清理後='{cleaned_value}'")
                            else:
                                print(f"      ❌ {field} 沒有匹配到任何內容")
                        except Exception as e:
                            print(f"      ❌ 提取 {field} 失敗: {e}")
                            import traceback
                            traceback.print_exc()
                            continue
                
                # 🔍 特別調試：主要建材的原始資料
                if "主要建材" in self.patterns:
                    print(f"🔍 調試主要建材原始提取:")
                    
                    # 顯示正則表達式
                    pattern = self.patterns["主要建材"]
                    print(f"  📋 使用的正則表達式: {pattern}")
                    
                    # 在原始文本中找主要建材
                    all_matches = re.finditer(pattern, search_text, re.MULTILINE | re.DOTALL)
                    for i, match in enumerate(all_matches):
                        print(f"  🎯 匹配{i+1}: 完整匹配='{match.group(0)}'")
                        print(f"        提取內容='{match.group(1)}'")
                        
                        # 🔧 新增：測試clean_value的結果
                        raw_content = match.group(1)
                        cleaned_content = self.clean_value(raw_content)
                        # print(f"        清理前內容: '{raw_content}'")
                        # print(f"        清理後內容: '{cleaned_content}'")
                        
                        # 顯示匹配位置前後的內容
                        start_pos = match.start()
                        end_pos = match.end()
                        before = search_text[max(0, start_pos-20):start_pos]
                        after = search_text[end_pos:end_pos+20]
                        # print(f"        前20字: '{before}'")
                        # print(f"        後20字: '{after}'")
                        
                        # 檢查是否包含RC造
                        full_context = search_text[max(0, start_pos-10):end_pos+30]
                        # print(f"        完整上下文: '{full_context}'")
                        
                        # 🔧 修正：檢查各種RC造的變體
                        rc_variants = ['RC造', 'ＲＣ造', 'rc造', 'Rc造']
                        found_rc = False
                        for variant in rc_variants:
                            if variant in full_context:
                                # print(f"        ✅ 上下文包含'{variant}'")
                                found_rc = True
                                break
                        
                        if not found_rc:
                            pass
                            # print(f"        ❌ 上下文不包含任何RC造變體")
                
                # 🔧 關鍵修正：建物層次面積的正確順序和獨立提取
                try:
                    layer_data = self._extract_building_layers_corrected_order(search_text)
                    data.update(layer_data)
                    print(f"      ✅ 層次面積提取完成，共 {len(layer_data)} 個欄位")
                except Exception as e:
                    print(f"      ❌ 層次面積提取失敗: {e}")
                
                # 🔧 關鍵修正：建築完成日期（在層次之後）
                try:
                    completion_match = re.search(self.patterns["建築完成日期"], search_text)
                    if completion_match:
                        data["建築完成日期"] = self.clean_value(completion_match.group(1))
                        print(f"      ✅ 建築完成日期: {data['建築完成日期']}")
                except Exception as e:
                    print(f"      ❌ 建築完成日期提取失敗: {e}")
                
                # 🔧 關鍵修正：附屬建物（在建築完成日期之後）
                try:
                    accessory_data = self._extract_accessory_buildings_fixed(search_text)
                    data.update(accessory_data)
                    print(f"      ✅ 附屬建物提取完成，共 {len(accessory_data)} 個欄位")
                except Exception as e:
                    print(f"      ❌ 附屬建物提取失敗: {e}")
                
                # 🔧 關鍵修正：共有部分和停車位解析（在附屬建物之後）
                try:
                    shared_data = self._extract_building_shared_parts_and_parking_enhanced(search_text)
                    data.update(shared_data)
                    print(f"      ✅ 共有部分提取完成，共 {len(shared_data)} 個欄位")
                except Exception as e:
                    print(f"      ❌ 共有部分和停車位提取失敗: {e}")
                        
            else:  # 土地謄本
                # 土地標示部的實際順序
                ordered_fields = [
                    "標示_登記日期", "標示_登記原因", "面積", "使用分區", 
                    "使用地類別", "公告土地現值", "地上建物建號"
                ]
                
                for field in ordered_fields:
                    if field in self.patterns:
                        try:
                            match = re.search(self.patterns[field], search_text, re.MULTILINE)
                            if match:
                                clean_name = field.replace("標示_", "")
                                if field == "公告土地現值" and len(match.groups()) == 2:
                                    # 特別處理公告土地現值：合併日期和金額
                                    date_part = match.group(1)
                                    amount_part = self.clean_value(match.group(2))
                                    data[clean_name] = f"{date_part} {amount_part}"
                                else:
                                    if field == "地上建物建號":
                                        
                                        # ↓↓↓ 多行抓取 ↓↓↓
                                        multi_val = self.get_multiline_indented(
                                            search_text,
                                            "地上建物建號",
                                            stop_labels=["其他登記事項"]  # 下一個可能 label
                                        )
                                        final_val = multi_val if multi_val else match.group(1)
                 
                                        data[clean_name] = self.clean_value(final_val)
                                    else:
                                        data[clean_name] = self.clean_value(match.group(1))

                                    # → 共用列印（保持在同一縮排層級）
                                    print(f"      ✅ 提取 {clean_name}: {data[clean_name]}")

                        except Exception as e:
                            print(f"      ❌ 提取 {field} 失敗: {e}")
                            continue
                
                # 🔧 土地謄本：處理其他登記事項（因為沒有共有部分函數處理）
                try:
                    other_items = self.extract_other_items_comprehensive(search_text, "標示部")
                    if other_items:
                        if isinstance(other_items, dict):
                            data.update(other_items)
                        else:
                            data["其他登記事項"] = other_items
                    else:
                        print(f"      ❌ 土地標示部沒有其他登記事項")
                except Exception as e:
                    print(f"      ❌ 土地標示部其他登記事項提取失敗: {e}")
            
            # 🔍 調試：顯示最終欄位順序
            for i, (key, value) in enumerate(data.items()):
                print(f"      {i+1:2d}. {key}: {str(value)[:50]}...")
            
        except Exception as e:
            print(f"    ❌ {doc_type}標示部整體提取失敗: {e}")
        
        return data

    def _extract_building_layers_corrected_order(self, text: str) -> OrderedDict:
        """建物層次面積的正確順序和命名"""
        layer_data = OrderedDict()
        
        # 步驟1：提取主層次（層次：一層）
        main_layer_pattern = r'層\s*次[:：]\s*([^\s]+)'
        main_layer_match = re.search(main_layer_pattern, text)
        if main_layer_match:
            layer_data["層次"] = main_layer_match.group(1)
        
        # 步驟2：提取主層次面積（層次面積：67.43平方公尺）
        main_area_pattern = r'層次面積[:：]\s*(\**[\d,]+\.?\d*\s*平方公尺)'
        main_area_match = re.search(main_area_pattern, text)
        if main_area_match:
            layer_data["層次面積"] = self.clean_value(main_area_match.group(1))
        
        # 步驟3：提取其他層次（編號格式）
        layer_line_pattern = r'((?:[一二三四五六七八九十夾]+層|騎樓|地下[一二三四五]?層|防空避難室)\s+\**[\d,]+\.?\d*\s*平方公尺)+'
        layer_line_matches = re.findall(layer_line_pattern, text)
        
        layer_index = 1
        
        for line_match in layer_line_matches:
            individual_layers = re.findall(r'([一二三四五六七八九十夾]+層|騎樓|地下[一二三四五]?層|防空避難室)\s+(\**[\d,]+\.?\d*\s*平方公尺)', line_match)
            
            for layer_name, area in individual_layers:
                clean_area = self.clean_value(area)
                
                # 如果是主層次（與layer_data["層次"]相同），跳過
                if main_layer_match and layer_name == main_layer_match.group(1):
                    continue
                
                # 按照格式A編號
                layer_data[f"層次{layer_index}"] = layer_name
                layer_data[f"層次面積{layer_index}"] = clean_area
                layer_index += 1

        # 步驟4：提取特殊面積格式（如：停車空間面積：３９．９２平方公尺）
        # 🔥 支援全形數字和冒號格式
        special_area_pattern = r'([^：\n]+面積)[:：]\s*([\d\.\,\uFF10-\uFF19\uFF0E\uFF0C]+)\s*平方公尺'
        special_area_matches = re.findall(special_area_pattern, text)

        for area_name, area_value in special_area_matches:
            area_name = area_name.strip()

            # 🔥 去掉末尾的「面積」二字（如：停車空間面積 → 停車空間）
            if area_name.endswith('面積'):
                area_name = area_name[:-2]

            # 🔥 將全形數字轉換為半形
            area_value_halfwidth = self.fullwidth_to_halfwidth(area_value)

            # 組合完整面積字串
            clean_area = f"{area_value_halfwidth}平方公尺"

            # 使用層次編號格式
            layer_data[f"層次{layer_index}"] = area_name
            layer_data[f"層次面積{layer_index}"] = clean_area
            layer_index += 1

            print(f"      ✅ 提取特殊面積: {area_name} = {clean_area} (原始: {area_value})")

        return layer_data

    def _extract_accessory_buildings_fixed(self, text: str) -> OrderedDict:
        """修正：附屬建物的獨立提取（支援跨行格式）"""
        accessory_data = OrderedDict()
        
        # 🔧 修正1：先找到附屬建物用途的起始位置
        accessory_line_pattern = r'附屬建物用途[:：]\s*([^\n]+)'
        accessory_match = re.search(accessory_line_pattern, text)
        
        if not accessory_match:
            print("    ❌ 沒有找到附屬建物用途")
            return accessory_data
        
        # 🔧 修正2：取得完整的附屬建物區塊（包括跨行內容）
        start_pos = accessory_match.start()
        lines = text.split('\n')
        start_line_index = text[:start_pos].count('\n')
        
        # 收集附屬建物相關的所有行
        all_accessory_content = []
        
        # 第一行：附屬建物用途：陽台 面積：4.37平方公尺
        first_line_content = accessory_match.group(1).strip()
        all_accessory_content.append(first_line_content)
        
        # 🔧 修正3：繼續收集後續行，直到遇到其他欄位
        i = start_line_index + 1
        while i < len(lines):
            line = lines[i].strip()
            
            if not line:
                i += 1
                continue
            
            # 停止條件：遇到其他欄位
            if (re.search(r'^共有部分[:：]', line) or 
                re.search(r'^權利範圍[:：]', line) or
                re.search(r'^其他登記事項[:：]', line) or
                re.search(r'^\*+.*?\*+', line) or
                re.search(r'^建物所有權部', line)):
                print(f"    🛑 遇到停止條件: {line}")
                break
            
            # 檢查是否為附屬建物相關內容（包含中文用途名稱 + 面積）
            if re.search(r'[\u4e00-\u9fff]+.*?[\d,]+\.?\d*\s*平方公尺', line):
                all_accessory_content.append(line)
                print(f"    ➕ 收集附屬建物行: {line}")
            else:
                print(f"    🛑 不符合附屬建物格式，停止: {line}")
                break
            
            i += 1
        
        # 🔧 修正4：將所有收集到的行合併成一個內容字串
        accessory_content = ' '.join(all_accessory_content)
        # 🔥 全形 → 半形正規化（NFKC）：把「３６．２９」轉成「36.29」、全形空白轉半形
        # 不然下面的數字 regex `[\d,]` 認不出全形數字，會漏抓
        try:
            import unicodedata as _ud
            accessory_content_normalized = _ud.normalize('NFKC', accessory_content)
            if accessory_content_normalized != accessory_content:
                print(f"    🔄 全形→半形正規化後: {accessory_content_normalized}")
            accessory_content = accessory_content_normalized
        except Exception:
            pass
        print(f"    📊 合併後的附屬建物內容: {accessory_content}")

        processed_items = set()  # 用於避免重複
        accessory_index = 0

        # 格式1：陽台 面積：4.37平方公尺
        pattern1 = r'([^\s]+)\s+面積[:：]\s*(\**[\d,]+\.?\d*\s*平方公尺)'
        matches1 = re.findall(pattern1, accessory_content)
        
        for accessory_type, area in matches1:
            clean_area = self.clean_value(area)
            item_key = f"{accessory_type}_{clean_area}"
            
            if item_key not in processed_items:
                processed_items.add(item_key)
                
                if accessory_index == 0:
                    accessory_data["附屬建物用途"] = accessory_type
                    accessory_data["附屬建物面積"] = clean_area
                else:
                    accessory_data[f"附屬建物用途{accessory_index}"] = accessory_type
                    accessory_data[f"附屬建物面積{accessory_index}"] = clean_area
                accessory_index += 1
        
        # 🔧 修正5：格式2處理（移除已匹配的部分後再匹配）
        remaining_content = accessory_content
        for accessory_type, area in matches1:
            # 更精確的替換模式
            patterns_to_remove = [
                f"{accessory_type} 面積：{area}",
                f"{accessory_type}面積：{area}",
                f"{accessory_type} 面積: {area}",
                f"{accessory_type}面積: {area}"
            ]
            for pattern_to_remove in patterns_to_remove:
                remaining_content = remaining_content.replace(pattern_to_remove, "")
        
        # 格式2：雨遮 2.27平方公尺（沒有"面積："標籤）
        pattern2 = r'([^\s]+)\s+(\**[\d,]+\.?\d*\s*平方公尺)'
        matches2 = re.findall(pattern2, remaining_content.strip())
        
        for accessory_type, area in matches2:
            clean_area = self.clean_value(area)
            item_key = f"{accessory_type}_{clean_area}"
            
            # 🔧 修正6：更嚴格的重複檢查
            already_processed = item_key in processed_items
            already_in_matches1 = accessory_type in [m[0] for m in matches1]
            
            if not already_processed and not already_in_matches1:
                processed_items.add(item_key)
                
                if accessory_index == 0:
                    accessory_data["附屬建物用途"] = accessory_type
                    accessory_data["附屬建物面積"] = clean_area
                else:
                    accessory_data[f"附屬建物用途{accessory_index}"] = accessory_type
                    accessory_data[f"附屬建物面積{accessory_index}"] = clean_area
                accessory_index += 1
            else:
                print(f"      ⏭️ 跳過重複項目: {accessory_type}")

        # 🔥 格式3 (fallback)：中文用途與數字之間沒空白，例「住宅、廚房、自用36.29平方公尺」
        # 用 [一-鿿、] 匹配中文+頓號，數字部分用 \s* 容許 0+ 空白
        if not accessory_data:
            pattern3 = r'([一-鿿、]+)\s*([\d,]+\.?\d*)\s*平方公尺'
            matches3 = re.findall(pattern3, accessory_content)
            for accessory_type, area_num in matches3:
                clean_area = self.clean_value(f"{area_num}平方公尺")
                item_key = f"{accessory_type}_{clean_area}"
                if item_key in processed_items:
                    continue
                processed_items.add(item_key)
                if accessory_index == 0:
                    accessory_data["附屬建物用途"] = accessory_type
                    accessory_data["附屬建物面積"] = clean_area
                else:
                    accessory_data[f"附屬建物用途{accessory_index}"] = accessory_type
                    accessory_data[f"附屬建物面積{accessory_index}"] = clean_area
                accessory_index += 1
                print(f"      ✅ [格式3 fallback] 用途={accessory_type}，面積={clean_area}")

        print(f"    🎯 附屬建物提取完成，共 {len(accessory_data)} 個欄位")

        return accessory_data

    # ========= 最終修正版：建物共有部分解析 =========
    def _extract_building_shared_parts_and_parking_enhanced(self, text: str) -> OrderedDict:
        """
        ◎ 功能總覽
        1. 有「共有部分：」── 完整解析共有部分、車位與組別內的其他登記事項
        2. 沒有「共有部分：」── 仍能擷取『標示部整體 其他登記事項』
        3. 「其他登記事項」多行續行：不再靠縮排深度，而用「停止條件」判斷
        
        🔧 修正重點：確保群組完整性，防止中間項目插入打亂順序
        """
        shared_data = OrderedDict()

        lines = text.split('\n')

        # ---------- 第一階段：掃描 ----------
        shared_parts_lines, other_items_lines = [], []
        for idx, raw in enumerate(lines):
            st = raw.strip()
            if '共有部分：' in st:
                shared_parts_lines.append((idx, st))
                print(f"      📦 第{idx+1}行 - 共有部分: {st}")
            if '其他登記事項：' in st:
                other_items_lines.append((idx, st))
                print(f"      📝 第{idx+1}行 - 其他登記事項: {st}")

        print(f"    📊 掃描結果: {len(shared_parts_lines)}個共有部分, "
            f"{len(other_items_lines)}個其他登記事項")

        # ---------- 無共有部分：直接抓標示部整體其他登記事項 ----------
        if not shared_parts_lines:
            print("    ⚠️ 沒有找到『共有部分：』— 僅解析標示部整體其他登記事項")

            def collect_other(start_i: int) -> str:
                """從 start_i 開始，把所有續行（直到遇停止條件）串成一行"""
                pieces = []
                m = re.search(r'其他登記事項[:：]\s*(.*)', lines[start_i])
                first = (m.group(1).strip() if m else "")
                # 「（空白）」也是有意義的字面標記，要保留（與「建築完成日期: （空白）」一致）
                if first:
                    pieces.append(first)

                j = start_i + 1
                while j < len(lines):
                    txt = lines[j].strip()
                    if not txt:
                        j += 1
                        continue
                    # 停止條件：遇章節標題／*****／括號開頭等
                    if (any(k in txt for k in [
                            '*****建物所有權部', '*****所有權部',
                            '建物所有權部', '所有權部', '共有部分：', '其他登記事項：']) or
                            txt.startswith('*****') or
                            '本謄本' in txt):
                        break
                    pieces.append(txt)
                    j += 1
                return self.clean_value(" ".join(pieces).strip())

            idx = other_items_lines[-1][0] if other_items_lines else -1
            if idx != -1:
                other_text = collect_other(idx)
                if other_text:
                    shared_data["其他登記事項"] = other_text
                    print(f"    ✅ 標示部整體其他登記事項: {other_text}")
                else:
                    print("    ⚠️ 找到標題但內容為空或（空白）")
            else:
                print("    ❌ 完全沒有『其他登記事項：』行")
            print("    ✅ 解析完成（無共有部分）")
            return shared_data

        # 以下為「有共有部分」的解析流程

        # -------------------------------------------------
        # 🔍 第二階段：確定最後一個其他登記事項（屬於標示部整體）
        last_other_idx = other_items_lines[-1][0] if other_items_lines else -1
        last_other_item_index = last_other_idx

        if last_other_idx != -1:
            print(f"      🎯 最後一個其他登記事項在第{last_other_idx+1}行: {lines[last_other_idx].strip()}")
            print("      📝 此項目將被視為【標示部整體】的其他登記事項")
        else:
            print("      ❌ 沒有找到其他登記事項")
        
        # 🔍 第三階段：兩階段處理 - 先收集所有群組基本信息，再處理細節
        print("    🔄 第三階段：兩階段群組處理開始")
        
        # 📋 階段3.1：收集群組邊界和基本信息（穩健版）
        group_boundaries = []
        for i, (idx, line) in enumerate(shared_parts_lines):
            # 🔧 穩健的群組邊界識別
            if i + 1 < len(shared_parts_lines):
                # 有下一個共有部分群組
                next_shared_idx = shared_parts_lines[i + 1][0]
                print(f"      📍 群組{i+1}邊界：下一個共有部分在第{next_shared_idx+1}行")
            else:
                # 最後一個群組，需要找到合適的結束點
                next_shared_idx = len(lines)
                
                # 尋找可能的結束標記
                for j in range(idx + 1, len(lines)):
                    line_text = lines[j].strip()
                    
                    # 明確的結束標記
                    if any(keyword in line_text for keyword in [
                        '所有權部', '*****所有權部', '建物所有權部', '*****建物所有權部'
                    ]):
                        next_shared_idx = j
                        print(f"      📍 群組{i+1}邊界：所有權部開始於第{j+1}行")
                        break
                    
                    # 如果遇到文檔結束提示
                    if '本謄本' in line_text or '詳細' in line_text:
                        next_shared_idx = j
                        print(f"      📍 群組{i+1}邊界：文檔結束於第{j+1}行")
                        break
                
                if next_shared_idx == len(lines):
                    print(f"      📍 群組{i+1}邊界：延伸到文檔末尾")
            
            group_info = {
                'group_num': i + 1,
                'start_line': idx,
                'end_line': next_shared_idx - 1,
                'building_num': '',
                'area': '',
                'rights_range': '',
                'parking_slots': [],
                'group_other_items': []
            }
            
            # 從共有部分行提取建號和面積
            shared_match = re.search(r'共有部分[:：]\s*([^*]+?)(\**[\d,]+\.?\d*\s*平方公尺)', line)
            if shared_match:
                group_info['building_num'] = shared_match.group(1).strip()
                group_info['area'] = self.clean_value(shared_match.group(2))
            
            group_boundaries.append(group_info)
            print(f"      🆕 群組{i + 1}: 第{idx + 1}行 - 第{next_shared_idx}行")
            print(f"          建號: {group_info['building_num']}")
            print(f"          面積: {group_info['area']}")
            print(f"          範圍涵蓋: {next_shared_idx - idx} 行")
            
            # 🔍 安全檢查：群組範圍是否合理
            if next_shared_idx - idx < 2:
                print(f"      ⚠️ 警告：群組{i+1}範圍過小，可能有邊界識別問題")
            elif next_shared_idx - idx > 50:
                print(f"      ⚠️ 警告：群組{i+1}範圍過大，可能有邊界識別問題")
        
        print(f"    📊 群組邊界識別完成：共識別出 {len(group_boundaries)} 個群組")
        
        # 📋 階段3.2：在各自範圍內填充詳細信息
        for group in group_boundaries:
            print(f"    🔍 處理群組{group['group_num']}詳細信息 (第{group['start_line'] + 1}-{group['end_line'] + 1}行)")
            
            # 🔍 先收集該群組範圍內的所有停車位（連續處理）
            print(f"          🔍 第一階段：收集群組{group['group_num']}的所有停車位")
            for line_idx in range(group['start_line'], min(group['end_line'] + 1, len(lines))):
                line = lines[line_idx].strip()
                
                # 停車位識別（括號內容）
                if line.startswith('（') and line.endswith('）'):
                    parking_content = line.strip('（）').strip()
                    
                    # 車位識別邏輯
                    is_parking = (
                        '停車位編號' in parking_content or
                        '車位編號' in parking_content or
                        '停車位' in parking_content or
                        '車位' in parking_content or
                        re.search(r'^\d+\s*,', parking_content) or
                        re.search(r'^\d+\s+', parking_content) or
                        re.search(r'權利範圍[:：]', parking_content)
                    )
                    
                    if is_parking:
                        # 清理車位內容
                        if parking_content.startswith('含停車位編號'):
                            parking_content = parking_content.replace('含停車位編號', '').strip()
                        elif parking_content.startswith('停車位編號'):
                            parking_content = parking_content.replace('停車位編號', '').strip()
                        
                        cleaned_parking = self.clean_value(parking_content)
                        group['parking_slots'].append(cleaned_parking)
                        print(f"          🚗 停車位: {cleaned_parking}")
            
            # 🔍 第二階段：收集其他欄位
            print(f"          🔍 第二階段：收集群組{group['group_num']}的其他欄位")
            for line_idx in range(group['start_line'], min(group['end_line'] + 1, len(lines))):
                line = lines[line_idx].strip()
                
                # 🔍 權利範圍 (排除括號開頭的)
                if '權利範圍：' in line and not line.startswith('（'):
                    rights_match = re.search(r'權利範圍[:：]\s*(\**[\d,]+分之[\d,]+)', line)
                    if rights_match:
                        group['rights_range'] = self.clean_value(rights_match.group(1))
                        print(f"          ⚖️ 權利範圍: {group['rights_range']}")
                
                # 🔍 群組內的其他登記事項 (排除最後一個全局的)
                elif '其他登記事項：' in line and line_idx != last_other_item_index:
                    match = re.search(r'其他登記事項[:：]\s*(.*)', line)
                    if match:
                        other_content = match.group(1).strip()
                        
                        # 收集多行內容
                        full_content = other_content
                        j = line_idx + 1
                        
                        # 在群組範圍內收集續行
                        while j <= group['end_line'] and j < len(lines):
                            next_line = lines[j].strip()
                            if not next_line:
                                j += 1
                                continue
                            
                            # 停止條件：遇到新的主要欄位
                            if (next_line.startswith('共有部分：') or 
                                next_line.startswith('其他登記事項：') or
                                next_line.startswith('權利範圍：') or
                                (next_line.startswith('（') and next_line.endswith('）'))):
                                break
                            
                            full_content += " " + next_line
                            j += 1
                        
                        # 清理並保存
                        full_content = re.sub(r'\*+', '', full_content).strip()
                        # 「（空白）」也是有意義的字面標記，要保留
                        if full_content:
                            group['group_other_items'].append(full_content)
                            print(f"          📝 群組其他登記事項: {full_content}")
        
        # 🔍 第四階段：提取標示部整體的其他登記事項
        print("    🔍 第四階段：提取標示部整體其他登記事項")
        indicator_other_item = ""
        
        if last_other_item_index != -1:
            print(f"      🎯 標示部整體其他登記事項起點：第{last_other_item_index + 1}行")
            
            parts = []
            m = re.search(r'其他登記事項[:：]\s*(.*)', lines[last_other_item_index])
            first = (m.group(1).strip() if m else "")
            # 「（空白）」也是有意義的字面標記，要保留（與「建築完成日期: （空白）」一致）
            if first:
                parts.append(first)

            j = last_other_item_index + 1
            while j < len(lines):
                txt = lines[j].strip()
                if not txt:
                    j += 1
                    continue
                # 停止條件
                if (any(k in txt for k in [
                        '*****建物所有權部', '*****所有權部',
                        '建物所有權部', '所有權部', '共有部分：', '其他登記事項：']) or
                        txt.startswith('*****') or
                        '本謄本' in txt):
                    break
                parts.append(txt)
                j += 1

            merged = self.clean_value(" ".join(parts).strip())
            if merged:
                indicator_other_item = merged
                print(f"      ✅ 標示部整體其他登記事項: {merged}")

        # 🔍 第五階段：按正確順序組裝輸出
        print("    🔍 第五階段：按群組順序組裝輸出")
        
        for i, group in enumerate(group_boundaries):
            group_num = i + 1
            print(f"    🔄 輸出群組 {group_num}:")
            
            # 🔧 確保每組按正確順序：建號 → 面積 → 權利範圍 → 所有停車位 → 群組其他登記事項
            
            # 1. 基本信息：建號 → 面積 → 權利範圍
            if group_num == 1:
                if group['building_num']:
                    shared_data["共有部分建號"] = group['building_num']
                    print(f"        ✅ 共有部分建號: {group['building_num']}")
                
                if group['area']:
                    shared_data["共有部分面積"] = group['area']
                    print(f"        ✅ 共有部分面積: {group['area']}")
                
                if group['rights_range']:
                    shared_data["權利範圍"] = group['rights_range']
                    print(f"        ✅ 權利範圍: {group['rights_range']}")
                
                # 2. 🔧 關鍵修正：所有停車位一次性全部輸出
                print(f"        🚗 第{group_num}組停車位一次性輸出 (共{len(group['parking_slots'])}個)")
                for j, parking in enumerate(group['parking_slots']):
                    if j == 0:
                        shared_data["含停車位編號"] = parking
                        print(f"        ✅ 含停車位編號: {parking}")
                    else:
                        shared_data[f"含停車位編號_1_{j}"] = parking
                        print(f"        ✅ 含停車位編號_1_{j}: {parking}")
                
                # 3. 🔧 停車位完成後才處理群組其他登記事項
                print(f"        📝 第{group_num}組其他登記事項 (共{len(group['group_other_items'])}個)")
                for j, other in enumerate(group['group_other_items']):
                    if j == 0:
                        shared_data["共有部分其他登記事項"] = other
                        print(f"        ✅ 共有部分其他登記事項: {other}")
                    else:
                        shared_data[f"共有部分其他登記事項_1_{j}"] = other
                        print(f"        ✅ 共有部分其他登記事項_1_{j}: {other}")
                
            else:
                # 其他組的處理邏輯相同
                if group['building_num']:
                    shared_data[f"共有部分建號{group_num}"] = group['building_num']
                    print(f"        ✅ 共有部分建號{group_num}: {group['building_num']}")
                
                if group['area']:
                    shared_data[f"共有部分面積{group_num}"] = group['area']
                    print(f"        ✅ 共有部分面積{group_num}: {group['area']}")
                
                if group['rights_range']:
                    shared_data[f"權利範圍{group_num}"] = group['rights_range']
                    print(f"        ✅ 權利範圍{group_num}: {group['rights_range']}")
                
                # 2. 🔧 關鍵修正：所有停車位一次性全部輸出
                print(f"        🚗 第{group_num}組停車位一次性輸出 (共{len(group['parking_slots'])}個)")
                for j, parking in enumerate(group['parking_slots']):
                    if j == 0:
                        shared_data[f"含停車位編號{group_num}"] = parking
                        print(f"        ✅ 含停車位編號{group_num}: {parking}")
                    else:
                        shared_data[f"含停車位編號{group_num}_{j}"] = parking
                        print(f"        ✅ 含停車位編號{group_num}_{j}: {parking}")
                
                # 3. 🔧 停車位完成後才處理群組其他登記事項
                print(f"        📝 第{group_num}組其他登記事項 (共{len(group['group_other_items'])}個)")
                for j, other in enumerate(group['group_other_items']):
                    if j == 0:
                        shared_data[f"共有部分其他登記事項{group_num}"] = other
                        print(f"        ✅ 共有部分其他登記事項{group_num}: {other}")
                    else:
                        shared_data[f"共有部分其他登記事項{group_num}_{j}"] = other
                        print(f"        ✅ 共有部分其他登記事項{group_num}_{j}: {other}")

        # 🔧 修正：輸出標示部整體的其他登記事項
        if indicator_other_item:
            shared_data["其他登記事項"] = indicator_other_item
            print(f"    ✅ 標示部整體其他登記事項輸出: {indicator_other_item}")

        # 🔧 修正：確保有共有部分時也要返回結果
        return shared_data

    # ========== 第1部分：修正OCR住址入列函數（無數量限制） ==========

    # ========== 第2部分：完整的 extract_ownership_section 函數（替換原本的） ==========
    def extract_ownership_section(self, section_text: str, ocr_data: Dict[str, str], doc_type: str) -> List[OrderedDict]:
        """修復版謄本所有權部提取 - 正確使用所有OCR住址"""
        global ocr_address_queue, used_address_count
        ownership_list = []
        
        print(f"🔍 開始解析謄本所有權部...")

        # 🔧 新增：清理殘留的跨頁內容
        # 移除單獨的字母行（如 G, EG, F0 等）
        lines = section_text.split('\n')
        cleaned_lines = []
        for line in lines:
            line_stripped = line.strip()
            # 跳過單獨的1-2個字母/數字
            if re.match(r'^[A-Z0-9]{1,2}$', line_stripped):
                print(f"  🗑️ 清理殘留字符: {line_stripped}")
                continue
            # 跳過孤立的地號行（沒有其他內容）
            if re.match(r'^.*?[區鄉鎮市].*?段\s+\d{4}-\d{4}地號$', line_stripped):
                print(f"  🗑️ 清理孤立地號: {line_stripped}")
                continue
            cleaned_lines.append(line)
        
        section_text = '\n'.join(cleaned_lines)
        
        # 檢查是否有地價資料
        price_pattern = r'\d+\.\d+元[／/]平方公尺'
        price_matches = re.findall(price_pattern, section_text)
        print(f"找到 {len(price_matches)} 個地價資料：")
        for price in price_matches[:10]:  # 顯示前10個
            print(f"  - {price}")
        
        print(f"  📊 OCR住址佇列長度: {len(ocr_address_queue)}")
        print(f"  📊 已使用住址數: {used_address_count}")
        
        # 顯示前10個可用住址
        if ocr_address_queue:
            print(f"  📋 可用OCR住址預覽:")
            start_idx = used_address_count
            for i in range(min(10, len(ocr_address_queue) - start_idx)):
                idx = start_idx + i
                if idx < len(ocr_address_queue):
                    print(f"    {idx+1}. {ocr_address_queue[idx][:50]}...")
        
        # 處理OCR資料
        if isinstance(ocr_data, dict):
            final_ocr_data = [ocr_data] if any(ocr_data.get(k, "").strip() for k in ["所有權人", "住址", "統一編號", "出生日期"]) else []
        elif isinstance(ocr_data, list):
            final_ocr_data = [d for d in ocr_data if isinstance(d, dict) and any(d.get(k, "").strip() for k in ["所有權人", "住址", "統一編號", "出生日期"])]
        else:
            final_ocr_data = []
        
        print(f"  📊 可用的OCR資料組數: {len(final_ocr_data)}")

        # 分割記錄
        records = re.findall(r'（(\d+)）登記次序[:：]\s*(\d+)(.*?)(?=（\d+）登記次序|本謄本|〈.*?〉|\*{10,}.*?他項權利部|$)', section_text, re.DOTALL)
        
        if not records:
            records = [('0001', '0001', section_text)]
        
        print(f"  📋 找到 {len(records)} 個所有權記錄")
        
        # 🔧 關鍵修正：確保正確使用全域索引
        start_index = used_address_count
        end_index = start_index + len(records)
        
        print(f"  📍 此謄本將使用住址索引 {start_index} 到 {end_index-1}")
        
        # 檢查住址是否足夠
        if end_index > len(ocr_address_queue):
            print(f"  ⚠️ 警告：OCR住址不足！需要{end_index}個，但只有{len(ocr_address_queue)}個")
            print(f"     將重複使用最後的住址或留空")
        
        # 建立住址分配計畫
        # 🔧 修正：不預先分配，改為按需使用
        print(f"  📍 此謄本從住址索引 {start_index} 開始按需使用")
        current_ocr_index = start_index
        # 🔧 修正：更新實際使用的住址數
 
        print(f"  📊 更新已使用住址數: {used_address_count}")
        
        # 處理每個記錄
        for i, (num, seq, content) in enumerate(records):
            print(f"\n  🔄 處理第 {i+1} 個所有權記錄（編號: {num}, 次序: {seq}）...")
            
            info = OrderedDict()
            info["記錄編號"] = num
            info["登記次序"] = seq
            
            # 登記相關資訊
            early_fields = [
                ("所有權_登記日期", "登記日期"),
                ("所有權_登記原因", "登記原因"), 
                ("所有權_原因發生日期", "原因發生日期")
            ]
            
            for field_pattern, field_name in early_fields:
                if field_pattern in self.patterns:
                    match = re.search(self.patterns[field_pattern], content, re.MULTILINE)
                    if match:
                        info[field_name] = self.clean_value(match.group(1))
                    else:
                        info[field_name] = ""
            
            # 所有權人資訊解析
            owner_fields = [("所有權人", "所有權人"), ("統一編號", "統一編號"), ("出生日期", "出生日期"), ("住址", "住址")]

            for field_pattern, field_name in owner_fields:
                if field_pattern in self.patterns:
                    # 🔥 所有權人欄位需要 DOTALL 旗標支援跨行匹配
                    flags = re.MULTILINE | re.DOTALL if field_pattern == "所有權人" else re.MULTILINE
                    match = re.search(self.patterns[field_pattern], content, flags)
                    if match:
                        if field_name == "統一編號":
                            raw_value = match.group(1).strip()
                            info[field_name] = raw_value  # 保留*號
                        else:
                            parsed_value = self.clean_value(match.group(1))
                            # 🔥 清理多餘的空白和換行
                            parsed_value = re.sub(r'\s+', '', parsed_value).strip()
                            info[field_name] = parsed_value
                        print(f"    ✅ {field_name}(文字解析): {info[field_name]}")
                    else:
                        info[field_name] = ""
                        print(f"    ❌ {field_name}(文字解析): 未找到")
            
            # 住址分配邏輯
            current_address = info.get("住址", "").strip()
            owner_name = info.get("所有權人", "").strip()
            unified_number = info.get("統一編號", "").strip()

            # 🔥 所有權人姓氏補字（第二類謄本中姓氏可能是圖片）
            _glyphs = globals().get('inline_glyphs', [])
            if owner_name and _glyphs:
                # 檢測是否只有遮罩符號（如「＊＊」或「**」）
                if re.fullmatch(r'[＊\*]+', owner_name):
                    glyph = _glyphs.pop(0)
                    owner_name = glyph + owner_name
                    info["所有權人"] = owner_name
                    print(f"    🔥 所有權人姓氏補字: → {owner_name}")

            # 🔍 調試：顯示當前記錄的文字資料
            print(f"    🔍 當前記錄分析:")
            print(f"      文字所有權人: '{owner_name}'")
            print(f"      文字統一編號: '{unified_number}'")
            print(f"      文字住址: '{current_address[:30]}...'")
            
            print(f"    🔍 住址分配分析:")
            print(f"       文字解析住址: '{current_address}'")
            print(f"       可用OCR住址: '{ocr_address_queue[current_ocr_index] if current_ocr_index < len(ocr_address_queue) else '無'}'")
            
            # 判斷是否需要使用OCR住址
            # 🔧 修正：根據謄本類型判斷需要OCR的欄位
            need_ocr_data = False
            missing_fields = []

            # 🔧 修正：檢查所有欄位是否都空白，如果都空白就需要完整OCR
            all_text_empty = (
                not owner_name and 
                not unified_number and 
                not current_address
            )

            if all_text_empty:
                # 所有文字都空白，需要完整OCR資料
                if not owner_name or len(owner_name) < 2:
                    missing_fields.append("所有權人")
                if not unified_number:
                    missing_fields.append("統一編號") 
                if not current_address or len(current_address) < 8:
                    missing_fields.append("住址")
                # 檢查出生日期
                # 🔧 修正：從OCR資料判斷是否為個人謄本
                if current_ocr_index < len(final_ocr_data):
                    ocr_record = final_ocr_data[current_ocr_index]
                    ocr_unified = ocr_record.get('統一編號', '')
                    # 判斷是否為個人（統一編號以英文字母開頭）
                    is_individual = ocr_unified and re.match(r'^[A-Z]', ocr_unified)
                    if is_individual:
                        birth_date = info.get("出生日期", "").strip()
                        if not birth_date:
                            missing_fields.append("出生日期")
                            print(f"       🔍 個人謄本需要出生日期")
                print(f"       ✅ 文字資料全空，需要完整OCR")
            else:
                # 🔧 修正：部分文字有資料，但要檢查統一編號
                # 即使所有權人和住址都有值，如果統一編號為空，也要從OCR補充
                print(f"       [DEBUG 1022-01] 檢查統一編號: '{unified_number}'")
                if not unified_number:
                    missing_fields.append("統一編號")
                    print(f"       ✅ [修正版] 統一編號為空，需要OCR補充")

                if not current_address or len(current_address) < 8 or "*" in current_address:
                    missing_fields.append("住址")
                    print(f"       ✅ 部分缺失，需要OCR住址")
            # else:
            #     # 二類謄本：主要檢查住址
            #     if not current_address or len(current_address) < 8 or "*" in current_address:
            #         missing_fields.append("住址")
            #         need_ocr_data = True
            #         print(f"       ✅ 二類謄本需要OCR住址")

            if missing_fields:
                need_ocr_data = True
                print(f"       ✅ 需要OCR資料: 缺失欄位 {missing_fields}")
            else:
                need_ocr_data = False
                print(f"       ❌ 文字資料足夠，跳過OCR")
            
            # 應用住址分配
            # 🔧 修正：統一處理所有OCR資料
            if need_ocr_data:
                # 🔧 關鍵修正：即使住址佇列為空，仍然可以使用OCR記錄中的其他欄位（如統一編號）
                ocr_record_index = current_ocr_index

                # 先處理住址（如果有的話）
                if current_ocr_index < len(ocr_address_queue) and "住址" in missing_fields:
                    info["住址"] = ocr_address_queue[current_ocr_index]
                    print(f"       ✅ 使用OCR住址[{current_ocr_index}]: {ocr_address_queue[current_ocr_index]}")

                # 🔧 關鍵：無論住址佇列如何，都要檢查OCR記錄中的其他欄位
                # 🔧 修正1022-04：如果索引超出範圍，重複使用最後一筆OCR資料
                ocr_record = None
                if ocr_record_index < len(final_ocr_data):
                    ocr_record = final_ocr_data[ocr_record_index]
                    print(f"       🔍 [修正版1022-03] 使用OCR記錄[{ocr_record_index}]補充其他欄位")
                elif len(final_ocr_data) > 0:
                    # 索引超出但有OCR資料，重複使用最後一筆
                    ocr_record = final_ocr_data[-1]
                    print(f"       🔍 [修正版1022-04] OCR索引超出，重複使用最後一筆OCR記錄")

                if ocr_record:
                    print(f"       📋 missing_fields: {missing_fields}")
                    print(f"       📋 ocr_record內容: {ocr_record}")

                    # 補充缺失的欄位
                    for field_name in missing_fields:
                        if field_name == "住址" and current_ocr_index >= len(ocr_address_queue):
                            # 住址佇列不足時跳過
                            print(f"       ⏭️ 跳過住址（佇列不足）")
                            continue
                        elif field_name == "住址":
                            # 住址已在上面處理
                            continue

                        ocr_value = ocr_record.get(field_name, "").strip()
                        original_value = info.get(field_name, "").strip()
                        print(f"       🔍 欄位'{field_name}': OCR值='{ocr_value}', 原始值='{original_value}'")
                        if ocr_value:
                            info[field_name] = ocr_value
                            print(f"       ✅ [修正版1022-04] 使用OCR{field_name}: {ocr_value}")
                        else:
                            # 🔥 修正：OCR 為空時保留原始值，不要覆蓋
                            print(f"       ⏭️ OCR{field_name}為空，保留原始值: '{original_value}'")

                    current_ocr_index += 1
                    print(f"       📊 OCR索引更新為: {current_ocr_index}")
                else:
                    print(f"       ❌ 完全沒有OCR記錄資料可用")
            else:
                print(f"       ⏭️ 文字資料完整，跳過OCR（索引保持: {current_ocr_index}）")
            
            
            # 權利範圍和權狀字號
            common_fields = [
                ("權利範圍", "權利範圍"), 
                ("權狀字號", "權狀字號"),
                # ("相關他項權利登記次序", "相關他項權利登記次序")  # 🔧 新增：建物也可能有此欄位
            ]
            for field_pattern, field_name in common_fields:
                if field_pattern in self.patterns:
                    match = re.search(self.patterns[field_pattern], content, re.MULTILINE)
                    if match:
                        info[field_name] = self.clean_value(match.group(1))
                    # else:
                    #     info[field_name] = ""
            
            # 土地特有欄位
            if doc_type == "土地謄本":
                print(f"    📊 開始解析土地特有欄位...")
                
                # 當期申報地價
                if "當期申報地價" in self.patterns:
                    match = re.search(self.patterns["當期申報地價"], content, re.MULTILINE)
                    if match:
                        info["當期申報地價"] = self.clean_value(match.group(1))
                        print(f"    ✅ 當期申報地價: {info['當期申報地價']}")
                    else:
                        info["當期申報地價"] = ""
                
                # 前次移轉現值和歷次取得權利範圍解析
                transfer_entries = []  # 儲存 (年月, 地價) 元組
                historical_values = []
                
                print(f"    🔍 開始解析前次移轉現值和歷次取得權利範圍...")
                
                # 逐行解析
                lines = content.split('\n')
                i = 0
                while i < len(lines):
                    line = lines[i].strip()
                    line_cleaned = re.sub(r'\*+', '', line)  # 🔧 修正：定義 line_cleaned
                    
                    # A. 檢查是否為前次移轉現值標題行
                    if '前次移轉現值或原規定地價' in line:
                        print(f"      找到前次移轉現值標題 (行{i+1})")
                        print(f"      標題行內容: '{line}'")
                        i += 1
                        
                        # 繼續讀取後續的年月和地價行
                        while i < len(lines):
                            next_line = lines[i].strip()
                            next_line_cleaned = re.sub(r'\*+', '', next_line)  # 移除星號
                            print(f"      檢查第{i+1}行: '{next_line}'")
                            print(f"      移除星號後: '{next_line_cleaned}'")
                            
                            # 🔧 修正：更嚴格的停止條件
                            # 只有遇到這些明確的欄位標題才停止
                            stop_keywords = [
                                '相關他項權利登記次序：',
                                '其他登記事項：',
                                '權狀字號：',
                                '當期申報地價：',
                                '（0002）登記次序',  # 下一筆記錄開始
                                '（0003）登記次序',  # 下一筆記錄開始
                                '〈 本謄本列印完畢 〉'
                                ]
                            
                            should_stop = False
                            for keyword in stop_keywords:
                                if keyword in next_line:
                                    print(f"      遇到停止關鍵字 '{keyword}'，停止")
                                    should_stop = True
                                    break
                            
                            if should_stop:
                                break

                            # 🔧 修正：跳過地號行但繼續處理
                            if re.search(r'[區鄉鎮市].*?段.*?\d{4}-\d{4}地號', next_line_cleaned):
                                print(f"      跳過地號行，繼續處理")
                                i += 1
                                continue
                            
                            # 檢查是否為歷次取得權利範圍
                            if '歷次取得權利範圍' in next_line_cleaned:
                                # 提取權利範圍值
                                if '：' in next_line_cleaned:
                                    rights_value = next_line_cleaned.split('：', 1)[1].strip()
                                elif ':' in next_line_cleaned:
                                    rights_value = next_line_cleaned.split(':', 1)[1].strip()
                                else:
                                    rights_value = ""
                                
                                if rights_value and rights_value not in ['（空白）', '(空白)', '無']:
                                    # 提取分數格式
                                    rights_match = re.search(r'(\d+分之\d+)', rights_value)
                                    if rights_match:
                                        historical_values.append(rights_match.group(1))
                                        print(f"      歷次取得權利範圍: {rights_match.group(1)}")
                                i += 1
                                continue
                            
                            # 檢查是否為年月和地價行
                            # 先嘗試匹配完整的年月和地價
                            year_month_match = re.search(r'(\d{3}年\d{2}月)', next_line_cleaned)
                            price_match = re.search(r'([\d,]+\.?\d*)\s*元[／/]平方公尺', next_line_cleaned)
                            
                            if year_month_match and price_match:
                                year_month = year_month_match.group(1)
                                price = f"{price_match.group(1)}元/平方公尺"
                                transfer_entries.append((year_month, price))
                                print(f"      ✅ 找到前次移轉: 年月={year_month}, 地價={price}")
                                i += 1
                                continue
                            elif year_month_match:
                                print(f"      ⚠️ 只找到年月: {year_month_match.group(1)}")
                                # 檢查下一行是否有地價
                                if i + 1 < len(lines):
                                    next_next_line = lines[i + 1].strip()
                                    next_next_cleaned = re.sub(r'\*+', '', next_next_line)
                                    price_match_2 = re.search(r'([\d,]+\.?\d*)\s*元[／/]平方公尺', next_next_cleaned)
                                    if price_match_2:
                                        year_month = year_month_match.group(1)
                                        price = f"{price_match_2.group(1)}元/平方公尺"
                                        transfer_entries.append((year_month, price))
                                        print(f"      ✅ 找到前次移轉(跨行): 年月={year_month}, 地價={price}")
                                        i += 2
                                        continue
                                i += 1
                            elif price_match:
                                print(f"      ⚠️ 只找到地價: {price_match.group(1)}元/平方公尺")
                                i += 1
                            elif not next_line.strip():
                                print(f"      空行，繼續")
                                i += 1
                            else:
                                print(f"      無法匹配任何模式，檢查內容...")
                                # 顯示更多調試信息
                                if len(next_line_cleaned) > 0:
                                    print(f"        原始長度: {len(next_line)}, 清理後長度: {len(next_line_cleaned)}")
                                    print(f"        前20字符: '{next_line_cleaned[:20] if len(next_line_cleaned) > 20 else next_line_cleaned}'")
                                break
                        continue
                    
                    # B. 單獨檢查歷次取得權利範圍（可能不在前次移轉現值區段內）
                    if '歷次取得權利範圍' in line_cleaned and '前次移轉現值' not in line_cleaned:
                        if '：' in line_cleaned:
                            rights_value = line_cleaned.split('：', 1)[1].strip()
                        elif ':' in line_cleaned:
                            rights_value = line_cleaned.split(':', 1)[1].strip()
                        else:
                            rights_value = ""
                        
                        if rights_value and rights_value not in ['（空白）', '(空白)', '無']:
                            rights_match = re.search(r'(\d+分之\d+)', rights_value)
                            if rights_match:
                                historical_values.append(rights_match.group(1))
                                print(f"      歷次取得權利範圍(單獨): {rights_match.group(1)}")
                    
                    i += 1
                
                # 插入資料到 info
                if transfer_entries or historical_values:
                    # 計算總筆數
                    total_count = max(len(transfer_entries), len(historical_values))
                    info["前次移轉現值或原規定"] = str(total_count)
                    print(f"    📊 前次移轉總筆數: {total_count}")
                    
                    # 插入年月和地價
                    for idx, (year_month, price) in enumerate(transfer_entries):
                        seq_num = f"{idx+1:04d}"
                        info[f"序號{idx+1}"] = seq_num
                        info[f"年月{idx+1}"] = year_month
                        info[f"地價{idx+1}"] = price
                        print(f"      ✅ 第{idx+1}筆: 年月={year_month}, 地價={price}")
                    
                    # 補充空欄位
                    for idx in range(len(transfer_entries), total_count):
                        seq_num = f"{idx+1:04d}"
                        info[f"序號{idx+1}"] = seq_num
                        info[f"年月{idx+1}"] = ""
                        info[f"地價{idx+1}"] = ""
                    
                    # 歷次取得權利範圍
                    for idx, value in enumerate(historical_values):
                        info[f"歷次取得權利範圍{idx+1}"] = value
                        print(f"      ✅ 歷次取得權利範圍{idx+1}: {value}")
                else:
                    info["前次移轉現值或原規定"] = "0"
                    print(f"    ❌ 沒有找到前次移轉現值記錄")

            # 相關他項權利登記次序（在其他登記事項之前）
            if "相關他項權利登記次序" in self.patterns:
                match = re.search(self.patterns["相關他項權利登記次序"], content, re.MULTILINE)
                if match:
                    extracted_value = self.clean_value(match.group(1))
                    if extracted_value.strip() and extracted_value.strip() not in ['', '（空白）', '(空白)', '無']:
                        info["相關他項權利登記次序"] = extracted_value
                        print(f"    ✅ 相關他項權利登記次序: {info['相關他項權利登記次序']}")

            # 其他登記事項
            other_items = self.extract_other_items_comprehensive(content, "所有權部")
            if other_items:
                if isinstance(other_items, dict):
                    info.update(other_items)
                else:
                    info["其他登記事項"] = other_items
            
            ownership_list.append(info)
            print(f"    📋 記錄{i+1}完成，共{len(info)}個欄位")
        
        # 🔧 修正：更新全域使用計數
        used_address_count = current_ocr_index
        # 最終統計
        print(f"\n  📊 所有權部處理完成: {len(ownership_list)} 筆記錄")
        print(f"  🔍 此謄本使用了OCR住址索引 {start_index} 到 {current_ocr_index - 1}")
        print(f"  📊 總共已使用: {used_address_count} 個住址")
        print(f"  📊 剩餘未使用: {len(ocr_address_queue) - used_address_count} 個住址")
        
        return ownership_list

# ==================== 圖片提取和文字提取函數 ====================
def extract_height60_images_only(pdf_path: str) -> List[Tuple[str, np.ndarray]]:
    """增強調試版圖片提取 - 加上頁面位置信息"""
    print(f"🔍 [DEBUG 1022-02修正版] 開始從PDF提取圖片: {os.path.basename(pdf_path)}")
    print(f"   目標高度TARGET_HEIGHT = {TARGET_HEIGHT}")
    
    if not os.path.exists(pdf_path):
        print(f"❌ PDF檔案不存在: {pdf_path}")
        return []

    try:
        doc = fitz.open(pdf_path)
        print(f"✅ PDF開啟成功，總頁數: {len(doc)}")
        
        images = []
        total_blocks = 0
        total_images = 0
        
        for page_num, page in enumerate(doc, 1):
            print(f"\n📄 處理第 {page_num} 頁...")
            
            # 方法1：從文字區塊中提取圖片
            page_dict = page.get_text("dict")
            blocks = page_dict.get("blocks", [])
            page_blocks = 0
            page_images = 0
            
            for block_idx, block in enumerate(blocks):
                total_blocks += 1
                page_blocks += 1
                
                if block.get("type") == 1:  # 圖片區塊
                    xref = block.get("xref")
                    
                    # 🔧 新增：獲取圖片的位置資訊
                    bbox = block.get("bbox", [0, 0, 0, 0])
                    x, y, width, height = bbox[0], bbox[1], bbox[2] - bbox[0], bbox[3] - bbox[1]
                    
                    print(f"  🖼️ 找到圖片區塊 {block_idx}, xref: {xref}")
                    print(f"    📍 圖片位置: 第{page_num}頁, x={x:.1f}, y={y:.1f}, 寬={width:.1f}, 高={height:.1f}")
                    
                    if xref:
                        try:
                            pix = fitz.Pixmap(doc, xref)
                            if pix.alpha:
                                pix = fitz.Pixmap(fitz.csRGB, pix)
                            img_arr = np.frombuffer(pix.tobytes(), dtype=np.uint8)
                            img_data = cv2.imdecode(img_arr, cv2.IMREAD_COLOR)
                            pix = None
                            
                            if img_data is not None:
                                img_height, img_width = img_data.shape[:2]
                                print(f"    📐 圖片尺寸: {img_width}x{img_height}")
                                total_images += 1
                                page_images += 1
                                
                                if img_height == TARGET_HEIGHT:
                                    # 🔧 新增：更詳細的標籤，包含位置信息
                                    detailed_label = f"頁{page_num}_Block{block_idx}_位置({x:.0f},{y:.0f})"
                                    images.append((detailed_label, img_data))
                                    print(f"    ✅ 符合目標高度 {TARGET_HEIGHT}，已收集")
                                    print(f"    📍 完整位置信息: {detailed_label}")
                                else:
                                    print(f"    ❌ 高度 {img_height} 不符合目標 {TARGET_HEIGHT}")
                            else:
                                print(f"    ❌ 圖片解碼失敗")
                        except Exception as e:
                            print(f"    ❌ 圖片處理錯誤: {e}")
                            continue

            # 方法2：從頁面圖片列表中提取
            try:
                page_image_list = page.get_images(full=True)
                print(f"  📋 頁面圖片列表包含 {len(page_image_list)} 個圖片")
                
                processed_xrefs = set()
                # 記錄已處理的xref，避免重複
                for block in blocks:
                    if block.get("type") == 1:
                        processed_xrefs.add(block.get("xref"))
                
                for img_idx, img_info in enumerate(page_image_list):
                    xref = img_info[0]
                    
                    # 跳過已經在方法1中處理過的圖片
                    if xref in processed_xrefs:
                        continue
                    
                    # 🔧 新增：嘗試獲取圖片在頁面上的位置（透過transform matrix）
                    try:
                        # 獲取圖片的轉換矩陣（如果可用）
                        transform = img_info[1] if len(img_info) > 1 else None
                        bbox_info = f"transform={transform}" if transform else "位置未知"
                    except:
                        bbox_info = "位置未知"
                        
                    print(f"  🖼️ 處理圖片列表項目 {img_idx}, xref: {xref}")
                    print(f"    📍 圖片位置: 第{page_num}頁, {bbox_info}")
                    
                    try:
                        pix = fitz.Pixmap(doc, xref)
                        if pix.alpha:
                            pix = fitz.Pixmap(fitz.csRGB, pix)
                        img_arr = np.frombuffer(pix.tobytes(), dtype=np.uint8)
                        img_data = cv2.imdecode(img_arr, cv2.IMREAD_COLOR)
                        pix = None
                        
                        if img_data is not None:
                            img_height, img_width = img_data.shape[:2]
                            print(f"    📐 圖片尺寸: {img_width}x{img_height}")
                            total_images += 1
                            page_images += 1
                            
                            if img_height == TARGET_HEIGHT:
                                # 🔧 新增：更詳細的標籤
                                detailed_label = f"頁{page_num}_列表{img_idx}_{bbox_info}"
                                images.append((detailed_label, img_data))
                                print(f"    ✅ 符合目標高度 {TARGET_HEIGHT}，已收集")
                                print(f"    📍 完整位置信息: {detailed_label}")
                            else:
                                print(f"    ❌ 高度 {img_height} 不符合目標 {TARGET_HEIGHT}")
                        else:
                            print(f"    ❌ 圖片解碼失敗")
                    except Exception as e:
                        print(f"    ❌ 圖片處理錯誤: {e}")
                        continue
                        
            except Exception as e:
                print(f"  ❌ 獲取頁面圖片列表失敗: {e}")
            
            print(f"  📊 第{page_num}頁統計: {page_blocks}個區塊, {page_images}個圖片")

        doc.close()
        
        print(f"\n📊 總體統計:")
        print(f"  總區塊數: {total_blocks}")
        print(f"  總圖片數: {total_images}")
        print(f"  目標高度: {TARGET_HEIGHT}")
        print(f"  符合條件的圖片: {len(images)}")
        
        if images:
            print(f"✅ 成功收集到 {len(images)} 張目標圖片:")
            for i, (label, img) in enumerate(images):
                height, width = img.shape[:2]
                print(f"  圖片{i+1}: {label}, 尺寸: {width}x{height}")
        else:
            print(f"❌ 沒有找到高度為 {TARGET_HEIGHT} 的圖片")
            print(f"💡 建議檢查:")
            print(f"   1. TARGET_HEIGHT 設定是否正確 (當前: {TARGET_HEIGHT})")
            print(f"   2. PDF是否包含圖片內容")
            print(f"   3. 圖片是否為其他高度")
        
        return images
        
    except Exception as e:
        print(f"❌ 圖片提取失敗: {e}")
        import traceback
        print(f"📋 詳細錯誤: {traceback.format_exc()}")
        return []
        
def split_documents_by_title(text: str) -> List[str]:
    """分割多份謄本"""
    patterns = [
        r"土地登記第[一二]類謄本.*?(?=土地登記第[一二]類謄本|建物登記第[一二]類謄本|$)",
        r"建物登記第[一二]類謄本.*?(?=土地登記第[一二]類謄本|建物登記第[一二]類謄本|$)"
    ]
    
    documents = []
    for pattern in patterns:
        matches = re.findall(pattern, text, re.DOTALL)
        documents.extend(matches)
    
    return [doc.strip() for doc in documents if doc.strip()]

# 修正的OCR住址佇列處理函數
def __push_addr_from_ocr(addr: str):
    """
    OCR住址入列函數 - 正式版
    
    設計說明：
    1. ✅ 允許重複住址（正常情況：夫妻、家人共同持有房產，住址相同）
    2. ✅ 保留100個上限（保護機制：避免異常情況導致記憶體問題）
    3. ✅ 按照OCR圖片出現順序依序加入佇列
    """
    global ocr_address_queue
    print(f"🔥🔥🔥 住址入列函數被呼叫: {addr}")
    
    try:
        if isinstance(addr, str) and addr.strip():
            clean_addr = addr.strip()
            
            # ✅ 允許重複住址，直接加入佇列（重複住址是正常的）
            # 🛡️ 保護機制：限制佇列最大100個，避免異常情況
            if len(ocr_address_queue) < 100:
                ocr_address_queue.append(clean_addr)
                print(f"    🧺 OCR住址入列: {clean_addr} (佇列位置: {len(ocr_address_queue)})")
            else:
                print(f"    🚫 佇列已滿(100個)，跳過: {clean_addr}")
    except Exception as e:
        print(f"    ❌ OCR住址入列失敗: {e}")


# 完整的 extract_comprehensive_ocr_data 函數（保留所有原始功能）
def extract_comprehensive_ocr_data(images: List[Tuple[str, np.ndarray]], ocr_mgr: TranscriptOCRManager) -> Tuple[List[Dict[str, str]], List[str]]:
    """統一處理一類和二類謄本的OCR資料組合邏輯 - 完整修正版"""
    global ocr_address_queue
    #     with open("ocr_debug.txt", "a", encoding="utf-8") as f:
    #         f.write(f"🔥 OCR函數開始，佇列狀態: {len(ocr_address_queue)}\n")
    #         f.write(f"🔥 OCR函數開始，佇列內容: {ocr_address_queue}\n")
    
    print(f"🔥 OCR函數開始，佇列狀態: {len(ocr_address_queue)}")
    print(f"🔥 OCR函數開始，佇列內容: {ocr_address_queue}")
    
    all_ocr_results = []
    NEED_KEYS = ("所有權人", "統一編號", "住址", "出生日期")
    
    print(f"🔍 開始處理 {len(images)} 張OCR圖片（統一邏輯：支援一類+二類）...")
    
    # 通用欄位模式（兼容一類和二類）
    field_patterns = {
        "所有權人": r"所有權人[:：]\s*(.+?)(?=\s*統一編號)",  # 🔥 支援跨行匹配
        "統一編號": r"統一編號[:：]\s*([^\s\n]+)",
        "出生日期": r"出生日[期日][:：]\s*([^\s\n]+)",
        "住址": r"住\s*址[:：]\s*([^\n]+)",
        "地址": r"地\s*址[:：]\s*([^\n]+)",
        "公司地址": r"公司地址[:：]\s*([^\n]+)",
        "住址_缺字": r"(?:^|\s)址[:：]\s*([^\n]+)",
    }
    
    print(f"📋 統一OCR邏輯特性:")
    print(f"   - 支援一類謄本（完整資料顯示）")
    print(f"   - 支援二類謄本（隱私保護）")
    print(f"   - 自動識別謄本類型並採用對應策略")
    
    # 第一階段：收集所有OCR識別結果
    raw_ocr_data = []
    current_owner_data = {}
    inline_glyphs = []  # 收集『單字小圖片』字元（例：館）
    
    for i, (img_label, img_data) in enumerate(images):
        # 🔍 記錄圖片處理資訊
        image_process_log = {
            'image_index': i,
            'image_label': img_label,
            'processed_time': time.strftime("%H:%M:%S")
        }
        result = ocr_mgr.ocr_image(img_data)
        if result:
            print(f"\n  📝 圖片{i+1} OCR結果: {result}")
            # 🔍 記錄OCR結果和圖片資訊
            image_process_log['ocr_result'] = result
            image_process_log['has_fields'] = False  # 先設為False，後面會更新
            raw_ocr_data.append(image_process_log.copy())
            print(f"🔥 圖片{i+1}處理前佇列長度: {len(ocr_address_queue)}")
            
            # 保險：先建立 text_compact
            text_compact = re.sub(r"\s+", "", (result or ""))
            
            # 提取所有可能的欄位
            found_fields = {}
            for field, pattern in field_patterns.items():
                # 🔥 所有權人欄位需要 DOTALL 旗標支援跨行匹配
                flags = re.DOTALL if field == "所有權人" else 0
                match = re.search(pattern, result, flags)
                if match:
                    value = match.group(1).strip()
                    # 🔥 清理多餘的空白和換行
                    value = re.sub(r'\s+', ' ', value).strip()
                    found_fields[field] = value
                    print(f"    ✅ 提取 {field}: {value}")
            
            # 🔧 關鍵修正：處理住址 - 只選擇一個最優的住址
            final_address = None
            address_source = None
            
            # 優先順序：住址 > 地址 > 公司地址 > 住址_缺字
            if "住址" in found_fields:
                final_address = found_fields["住址"]
                address_source = "住址"
            elif "地址" in found_fields:
                final_address = found_fields["地址"]
                address_source = "地址"
                found_fields["住址"] = final_address  # 統一存為"住址"
            elif "公司地址" in found_fields:
                final_address = found_fields["公司地址"]
                address_source = "公司地址"
                found_fields["住址"] = final_address  # 統一存為"住址"
            elif "住址_缺字" in found_fields:
                final_address = found_fields["住址_缺字"]
                address_source = "住址_缺字"
                found_fields["住址"] = final_address  # 統一存為"住址"
            
            # 清理多餘的住址欄位，只保留統一的"住址"
            for field_to_remove in ["地址", "公司地址", "住址_缺字"]:
                if field_to_remove in found_fields:
                    del found_fields[field_to_remove]
            
            # 🔧 只將最終選定的住址入列一次
            if final_address:
                __push_addr_from_ocr(final_address)
                current_queue_length = len(ocr_address_queue)
                print(f"    🏠 最終住址入佇列(來源:{address_source}, 第{current_queue_length}位): {final_address}")
            
            if found_fields:
                # 🔍 標記這張圖片有找到欄位
                image_process_log['has_fields'] = True
                image_process_log['found_fields'] = list(found_fields.keys())
                # 如果遇到新的所有權人，保存之前的資料並開始新的一組
                if "所有權人" in found_fields:
                    # 只要四大欄位有其一就保留
                    if current_owner_data and any(current_owner_data.get(k) for k in NEED_KEYS):
                        all_ocr_results.append(current_owner_data.copy())
                        owner_show = current_owner_data.get('所有權人') or current_owner_data.get('住址') or current_owner_data.get('統一編號') or '（無）'
                        print(f"    💾 完成第 {len(all_ocr_results)} 組OCR資料: {owner_show}")
                    
                    # 開始新的一組
                    current_owner_data = found_fields.copy()
                    first_show = (found_fields.get('所有權人') or found_fields.get('住址') or found_fields.get('統一編號') or '（無）')
                    print(f"    🆕 開始新的 OCR 資料組: {first_show}")
                else:
                    # 補充到當前組
                    current_owner_data.update(found_fields)
                    print(f"    📋 補充到當前組資料: {list(found_fields.keys())}")
                
                print(f"🔥 圖片{i+1}處理後佇列長度: {len(ocr_address_queue)}")
            else:
                # 🔧 修正：只有沒找到欄位時才處理內嵌字元
                text_compact = re.sub(r"\s+", "", result)
                
                # 檢查是否為內嵌字元（移到else內，避免重複處理）
                if is_inline_glyph(text_compact):
                    print(f"    🔍 檢測到內嵌字元原始: '{text_compact}'")
                    # 🔧 修正：將多字內嵌字元拆分為個別字元
                    for char in text_compact:
                        if re.match(r"[\u4e00-\u9fff]", char):  # 確保是中文字
                            inline_glyphs.append(char)
                            print(f"    🧩 捕捉到內嵌字元(拆分): {char}")
                    print(f"    📌 拆分完成，內嵌字元列表: {inline_glyphs}")
    # 🔍 OCR處理分析報告
    print(f"\n📊 OCR圖片處理分析:")
    print(f"   總圖片數: {len(raw_ocr_data)}")

    # 統計有欄位 vs 無欄位的圖片
    with_fields = [img for img in raw_ocr_data if img.get('has_fields', False)]
    without_fields = [img for img in raw_ocr_data if not img.get('has_fields', False)]

    print(f"   有欄位圖片: {len(with_fields)} 張")
    print(f"   無欄位圖片: {len(without_fields)} 張")

    # 分析重複的OCR結果
    ocr_content_count = {}
    for img_data in raw_ocr_data:
        content = img_data.get('ocr_result', '')
        if content:
            ocr_content_count[content] = ocr_content_count.get(content, 0) + 1

    print(f"\n🔍 重複OCR內容分析:")
    for content, count in ocr_content_count.items():
        if count > 1:
            print(f"   '{content}' 出現 {count} 次")

    
    # 保存最後一組（只要四大欄位有其一就保留）
    if current_owner_data and any(current_owner_data.get(k) for k in NEED_KEYS):
        all_ocr_results.append(current_owner_data.copy())
        owner_show = (current_owner_data.get('所有權人') or current_owner_data.get('住址') or current_owner_data.get('統一編號') or '（無）')
        print(f"    💾 完成第 {len(all_ocr_results)} 組OCR資料: {owner_show}")
    
    print(f"\n🔄 第二階段：標準化和清理OCR資料...")
    
    # 第二階段：標準化處理
    cleaned_results = []
    
    for i, raw_data in enumerate(all_ocr_results):
        print(f"\n  🔄 處理第 {i+1} 組OCR資料:")
        
        # 先在原始欄位上做通用補字（涵蓋住址/地段/地號/姓名…）
        if 'inline_glyphs' in locals() and inline_glyphs:
            patch_mapping_all_strings(raw_data, inline_glyphs)
            print(f"    🔧 已先對原始 OCR 欄位補字，剩餘內嵌字: {inline_glyphs}")
        
        # 建立標準化記錄
        cleaned_record = {}
        
        # 1. 所有權人
        owner = raw_data.get("所有權人", "").strip()
        cleaned_record["所有權人"] = owner
        print(f"    📋 所有權人: {owner}")
        
        # 2. 統一編號
        id_number = raw_data.get("統一編號", "").strip()
        cleaned_record["統一編號"] = id_number
        print(f"    📋 統一編號: {id_number}")
        
        # 3. 出生日期
        birth_date = raw_data.get("出生日期", "").strip()
        cleaned_record["出生日期"] = birth_date
        print(f"    📋 出生日期: {birth_date}")
        
        # 4. 住址（已經在第一階段統一處理）
        address = raw_data.get("住址", "").strip()
        
        # 住址清理
        if address:
            address = address.replace("毫南市", "臺南市")
            address = address.replace("镇", "鎮")
            address = address.replace("县", "縣")
            address = address.replace("鳳山品", "鳳山區")
            address = address.replace("燕巢品", "燕巢區")
            address = address.replace("左警", "左營")
            address = address.replace(" 後段", "廍後段")
            address = address.replace("廊", "廍")
            address = address.replace("苓稚", "苓雅")
            # 🔧 新增：補充其他簡體字
            address = address.replace("岗", "崗")
            address = address.replace("东", "東")
            address = address.replace("区", "區")
            address = address.replace("术", "術")
        
        cleaned_record["住址"] = address
        
        # 住址補字（OCR小圖 → 章節即時套用）
        _self = locals().get('self')
        _glyphs = ((_self and getattr(_self, 'inline_glyphs', None)) or locals().get('inline_glyphs') or globals().get('inline_glyphs'))
        if _glyphs and 'address' in locals() and address:
            try:
                address = patch_inline_glyphs_in_text(address, _glyphs)
                cleaned_record["住址"] = address
            except Exception as _e:
                print(f"    ⚠️ 住址補字失敗: {_e}")
        
        print(f"    📋 住址: {address}")
        
        # 5. 自動識別謄本類型
        if len(id_number) == 8 and id_number.isdigit():
            transcript_type = "法人"
            data_complete = bool(owner and id_number)
        elif "*" in owner or "*" in id_number:
            transcript_type = "二類個人"
            data_complete = bool(owner and address)  # 二類個人重點是住址
        elif owner and id_number and birth_date and address:
            transcript_type = "一類個人"
            data_complete = True
        else:
            transcript_type = "待補充"
            data_complete = False
        
        cleaned_record["謄本類型"] = transcript_type
        cleaned_record["資料完整"] = data_complete
        
        print(f"    🏷️ 識別類型: {transcript_type}")
        print(f"    📊 資料完整: {'✅' if data_complete else '❌'}")
        
        # 6. 資料驗證和修正
        if transcript_type == "一類個人":
            # 一類個人：確保所有欄位都有值
            if not owner or not id_number or not birth_date or not address:
                print(f"    ⚠️ 一類個人資料不完整，標記為待補充")
                cleaned_record["資料完整"] = False
        elif transcript_type == "二類個人":
            # 二類個人：重點是住址
            if not address:
                print(f"    ⚠️ 二類個人缺少住址，標記為待補充")
                cleaned_record["資料完整"] = False
        elif transcript_type == "法人":
            # 法人：所有權人和統一編號
            if not owner or not id_number:
                print(f"    ⚠️ 法人基本資料不完整，標記為待補充")
                cleaned_record["資料完整"] = False
        
        # 在寫入前再補一次（處理清理/合成後仍殘留的空洞）
        if 'inline_glyphs' in locals() and inline_glyphs:
            patch_mapping_all_strings(cleaned_record, inline_glyphs)
            print(f"    🔧 已對標準化記錄補字，剩餘內嵌字: {inline_glyphs}")
        
        cleaned_results.append(cleaned_record)
    
    # 統計結果
    print(f"\n📊 統一OCR最終統計:")
    print(f"   總記錄數: {len(cleaned_results)}")
    print(f"   OCR住址佇列總數: {len(ocr_address_queue)}")
    
    # 顯示所有OCR住址
    print(f"\n📋 OCR住址佇列內容:")
    for i, addr in enumerate(ocr_address_queue[:20]):  # 顯示前20個
        print(f"   {i+1}. {addr}")
    if len(ocr_address_queue) > 20:
        print(f"   ... 還有 {len(ocr_address_queue) - 20} 個住址")
    
    # 檢查重複情況
    from collections import Counter
    addr_counter = Counter(ocr_address_queue)
    duplicates = {addr: count for addr, count in addr_counter.items() if count > 1}
    
    if duplicates:
        print(f"\n⚠️ 發現重複住址:")
        for addr, count in duplicates.items():
            print(f"   {addr[:40]}... (重複{count}次)")
        print(f"   💡 建議檢查OCR是否重複掃描了相同的圖片")
    else:
        print(f"\n✅ 沒有重複住址，每個住址都是唯一的")
    
    # 統計各類型
    type_stats = {}
    for record in cleaned_results:
        t_type = record.get("謄本類型", "未知")
        type_stats[t_type] = type_stats.get(t_type, 0) + 1
    
    for t_type, count in type_stats.items():
        complete_count = sum(1 for r in cleaned_results if r.get("謄本類型") == t_type and r.get("資料完整"))
        print(f"   {t_type}: {count} 個 (完整: {complete_count})")
    
    # 返回標準格式（移除內部欄位）
    final_results = []
    for record in cleaned_results:
        final_record = {
            "所有權人": record["所有權人"],
            "統一編號": record["統一編號"],
            "出生日期": record["出生日期"],
            "住址": record["住址"]
        }
        final_results.append(final_record)
    
    print(f"\n🎯 統一OCR處理完成: 共 {len(final_results)} 套標準化資料")
    
    # 🔧 新增：住址與統一編號的配對分析
    print(f"\n🔍 OCR配對分析:")
    for i, record in enumerate(final_results):
        owner = record["所有權人"]
        id_num = record["統一編號"]
        birth = record["出生日期"]
        addr = record["住址"]
        
        # 分析統一編號模式
        id_pattern = "unknown"
        if id_num:
            if len(id_num) == 8 and id_num.isdigit():
                id_pattern = "法人"
            elif "*" in id_num:
                id_pattern = "二類個人"
            else:
                id_pattern = "一類個人"
        
        print(f"  資料{i+1}: {owner} | {id_num} ({id_pattern}) | {addr[:30]}...")
        
        # 檢查住址合理性
        if addr:
            district_count = len(re.findall(r'[區鄉鎮市]', addr))
            if district_count > 2:
                print(f"    ⚠️ 住址可能有問題（包含{district_count}個行政區）")
    
    # 🔧 新增：為智慧配對準備額外資訊
    print(f"\n📋 為智慧配對準備統一編號對應表:")
    id_address_map = {}
    for i, record in enumerate(final_results):
        id_num = record["統一編號"]
        addr = record["住址"]
        if id_num and addr:
            id_address_map[id_num] = addr
            print(f"  {id_num} → {addr[:40]}...")
    
    # 儲存全域變數
    globals()['inline_glyphs'] = inline_glyphs
    globals()['ocr_id_address_map'] = id_address_map  # 儲存配對資訊供後續使用
    
    return final_results, inline_glyphs


class UnifiedTranscriptProcessor:
    def __init__(self):
        """初始化處理器"""
        print("🔧 初始化謄本處理器...")
        self.ocr_mgr = TranscriptOCRManager()
        self.parser = UnifiedTranscriptParser()
        self.inline_glyphs = []
        print("✅ 謄本處理器初始化完成")
    
    # 修正 process_multiple_pdfs 中的佇列清空邏輯
    def process_multiple_pdfs(self, pdf_paths: List[str]) -> List[Dict[str, Any]]:
        """處理多個謄本PDF - 修正住址佇列管理"""
        global ocr_address_queue
        all_results = []
        
        for i, pdf_path in enumerate(pdf_paths, 1):
            print(f"\n{'='*60}")
            print(f"📄 處理第 {i} 個PDF: {os.path.basename(pdf_path)}")
            print(f"{'='*60}")
            
            # 🔧 修正：清空住址佇列，為新PDF準備
            ocr_address_queue = []  # 使用賦值而不是 clear()
            
            # 重置全域計數器
            global used_address_count
            used_address_count = 0
            print(f"🗑️ 已清空OCR住址佇列（處理新PDF）")
            
            if not os.path.exists(pdf_path):
                print(f"❌ 檔案不存在: {pdf_path}")
                continue

            # OCR 處理
            print("🎯 提取高度60的圖片進行 OCR...")
            images = extract_height60_images_only(pdf_path)

            if images:
                print(f"🔄 開始OCR識別 ({len(images)}張圖片)...")
                all_ocr_results, self.inline_glyphs = extract_comprehensive_ocr_data(images, self.ocr_mgr)
                
                # 🔧 修正：顯示實際的住址數量
                total_addresses = len(ocr_address_queue)
                print(f"📊 OCR處理完成:")
                print(f"   圖片數量: {len(images)}")
                print(f"   OCR住址總數: {total_addresses}")
                print(f"   OCR記錄組數: {len(all_ocr_results)}")
                
                # 顯示住址統計
                if ocr_address_queue:
                    # 統計重複的住址
                    from collections import Counter
                    addr_counter = Counter(ocr_address_queue)
                    unique_count = len(addr_counter)
                    
                    print(f"\n📋 住址統計:")
                    print(f"   總住址數: {total_addresses}")
                    print(f"   唯一住址數: {unique_count}")
                    
                    # 顯示最常見的住址
                    most_common = addr_counter.most_common(5)
                    print(f"   最常見的住址:")
                    for addr, count in most_common:
                        print(f"     {addr[:40]}... (出現{count}次)")
                    
            else:
                all_ocr_results = []
                print("⚠️ 沒有找到可 OCR 的圖片")
            
            # 提取文字內容
            print("📄 提取 PDF 文字內容...")
            full_text = extract_pdf_text_only(pdf_path)
            if not full_text:
                print("❌ 無法提取 PDF 文字內容")
                continue
            
            # 分割謄本
            print("✂️ 分割謄本...")
            raw_documents = split_documents_by_title(full_text)
            print(f"📋 找到 {len(raw_documents)} 份謄本")
            
            # 🔧 新增：檢查住址數量是否足夠
            total_ownership_records = 0
            for raw_doc in raw_documents:
                # 簡單計算所有權記錄數（通過登記次序）
                ownership_count = len(re.findall(r'（\d+）登記次序[:：]', raw_doc))
                if ownership_count == 0:
                    ownership_count = 1  # 至少有一筆
                total_ownership_records += ownership_count
            
            print(f"\n📊 住址分配預測:")
            print(f"   預計所有權記錄總數: {total_ownership_records}")
            print(f"   可用OCR住址數: {len(ocr_address_queue)}")
            
            if len(ocr_address_queue) < total_ownership_records:
                print(f"   ⚠️ 警告：OCR住址可能不足 (差{total_ownership_records - len(ocr_address_queue)}個)")
            else:
                print(f"   ✅ OCR住址充足")
            
            # 處理每份謄本
            for doc_i, raw_doc in enumerate(raw_documents):
                print(f"\n📄 處理第 {doc_i+1} 份謄本...")
                
                # 🔧 關鍵修正：在清理跨頁前先提取列印時間
                print_time = None
                print_time_pattern = r"列印時間[:：]\s*([^\n]+?)(?:\s+頁次[:：].*?)?(?:\n|$)"
                print_time_match = re.search(print_time_pattern, raw_doc)
                if print_time_match:
                    print_time = print_time_match.group(1).strip()
                    print(f"    ✅ 提取列印時間（清理前）: {print_time}")
                
                # 🔧 先修復跨行問題，遇到有造字圖片被折行時可在這裡加入
                raw_doc = raw_doc.replace('共同擔保地號\n:', '共同擔保地號:')
                raw_doc = raw_doc.replace('共同擔保建號\n:', '共同擔保建號:')
                raw_doc = raw_doc.replace('鳥松區大\n腿段', '鳥松區大脚腿段')
                raw_doc = raw_doc.replace('鳥松區大\n\n腿段', '鳥松區大脚腿段')
                raw_doc = raw_doc.replace('左營區\n後段', '左營區廍後段')
                raw_doc = raw_doc.replace('左營區\n\n後段', '左營區廍後段')
                raw_doc = raw_doc.replace('磚子\n段', '磚子磘段')
                raw_doc = raw_doc.replace('磚子\n\n段', '磚子磘段')


                
                # 立即清理該份謄本的跨頁資料
                cleaned_doc = self.parser.clean_single_transcript_cross_page(raw_doc)
                
                try:
                    # 傳遞OCR資料
                    if all_ocr_results:
                        ocr_data = all_ocr_results
                        print(f"    📋 傳遞OCR資料: {len(all_ocr_results)} 套")
                    else:
                        ocr_data = {}
                        print(f"    📋 沒有OCR資料")
                    
                    # 解析謄本（使用已清理的文本）
                    structured_data = self.parser.extract_structured_data(cleaned_doc, ocr_data)
                    
                    # 🔧 關鍵修正：將預先提取的列印時間加入基本資訊
                    if print_time and "基本資訊" in structured_data:
                        structured_data["基本資訊"]["列印時間"] = print_time
                        print(f"    ✅ 已將列印時間加入結構化資料")

                    # 添加源文件資訊
                    structured_data["源檔案"] = os.path.basename(pdf_path)
                    structured_data["檔案序號"] = i
                    structured_data["謄本序號"] = doc_i + 1
                    
                    all_results.append(structured_data)
                    print(f"  ✅ 謄本 {doc_i+1} 解析完成")
                    
                except Exception as e:
                    print(f"  ❌ 謄本 {doc_i+1} 解析失敗: {e}")
                    import traceback
                    traceback.print_exc()
            
            # 🔧 顯示最終使用情況
            print(f"\n📊 住址使用統計:")
            print(f"   總共使用: {used_address_count} 個住址")
            print(f"   剩餘未用: {len(ocr_address_queue) - used_address_count} 個住址")
        
        return all_results

# ==================== 輸出器（修正Excel欄位順序） ====================
class UnifiedTranscriptExporter:
    """謄本專用輸出器 - 修正Excel欄位順序"""
    
    def __init__(self, output_dir: str = "output", gui_instance=None):
        # 🔥 確保 output_dir 是絕對路徑，避免在 _internal 目錄下創建
        if not os.path.isabs(output_dir):
            # 如果是相對路徑，轉換為絕對路徑
            if hasattr(sys, '_MEIPASS'):
                # PyInstaller 打包環境：基於執行檔所在目錄
                exe_dir = os.path.dirname(os.path.abspath(sys.executable))
                if '_internal' in exe_dir or exe_dir.endswith('_internal'):
                    exe_dir = os.path.dirname(exe_dir)
                output_dir = os.path.join(exe_dir, output_dir)
            else:
                # 開發環境：基於腳本所在目錄
                output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), output_dir)

        self.output_dir = output_dir
        self.gui_instance = gui_instance  # 保留 GUI 實例引用（如果有）
        os.makedirs(output_dir, exist_ok=True)

        # 🔧 記錄輸出目錄，讓 atexit 知道要保存日誌到哪裡
        _set_output_dir(output_dir)
        
        auto_install("pandas")
        auto_install("openpyxl")

    def _create_flexible_filename(self, data: List[Dict], format_suffix: str) -> str:
        """靈活檔案命名系統 - 支援一類+二類，友善的中文檔名"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")

        # 分析謄本類型
        first_class_count = sum(1 for t in data if "第一類" in str(t.get("謄本類型", "")))
        second_class_count = sum(1 for t in data if "第二類" in str(t.get("謄本類型", "")))
        building_count = sum(1 for t in data if "建物謄本" in str(t.get("謄本類型", "")))
        land_count = sum(1 for t in data if "土地謄本" in str(t.get("謄本類型", "")))

        # 🔧 新增：提取地段地號資訊
        location_part = ""
        if data and len(data) > 0:
            # 從第一份謄本的基本資訊中提取地號建號
            first_transcript = data[0]
            basic_info = first_transcript.get("基本資訊", {})
            lot_info = basic_info.get("地號建號", "")

            if lot_info:
                # 解析地號建號，例如："大寮區仁德段 0193-0000地號"
                import re
                # 提取區段和地號/建號
                match = re.match(r'([\u4e00-\u9fff]+區)?([\u4e00-\u9fff]+段)\s*(\d+[\-\d]*)(地號|建號)', lot_info)
                if match:
                    area = match.group(1) or ""  # 可能沒有區
                    section = match.group(2)
                    number = match.group(3).replace('-0000', '').replace('-000', '').lstrip('0') or '0'
                    lot_type = match.group(4)

                    # 移除多餘的0，例如 0193 -> 193
                    if '-' in number:
                        parts = number.split('-')
                        number = parts[0].lstrip('0') or '0'
                        if parts[1] != '0000' and parts[1] != '000':
                            number += '-' + parts[1].lstrip('0')

                    location_part = f"{area}{section}{number}"

        # 智慧命名邏輯（中文化）
        if first_class_count > 0 and second_class_count > 0:
            class_part = "混合類"
        elif second_class_count > 0:
            class_part = "第二類"
        else:
            class_part = "第一類"

        if building_count > 0 and land_count > 0:
            type_part = "土地建物"
        elif building_count > 0:
            type_part = "建物"
        else:
            type_part = "土地"

        # 組合檔名（加上 trans_ 前綴表示結構化謄本）
        if location_part:
            # 有地段地號資訊
            if format_suffix == "txt":
                return f"trans_{location_part}_{type_part}謄本_{class_part}_{timestamp}.{format_suffix}"
            else:
                return f"trans_{location_part}_{type_part}謄本_{class_part}_{timestamp}.{format_suffix}"
        else:
            # 沒有地段地號資訊，使用舊格式
            if format_suffix == "txt":
                return f"trans_{type_part}謄本_{class_part}_報告_{timestamp}.{format_suffix}"
            else:
                return f"trans_{type_part}謄本_{class_part}_{timestamp}.{format_suffix}"

    def _get_land_ownership_field_order(self) -> List[str]:
        """土地所有權部欄位順序 - 預先定義足夠多的序號"""
        return [
            # 基本資訊
            "編號", "源檔案", "地號建號", "記錄編號", "登記次序", 
            "登記日期", "登記原因", "原因發生日期", 
            "所有權人", "統一編號", "出生日期", "住址", 
            "權利範圍", "權狀字號", "當期申報地價", 
            "前次移轉現值或原規定",
            
            # 🔧 預先定義大量序號（1-50）
            "序號1", "年月1", "地價1", "歷次取得權利範圍1",
            "序號2", "年月2", "地價2", "歷次取得權利範圍2",
            "序號3", "年月3", "地價3", "歷次取得權利範圍3",
            "序號4", "年月4", "地價4", "歷次取得權利範圍4",
            "序號5", "年月5", "地價5", "歷次取得權利範圍5",
            "序號6", "年月6", "地價6", "歷次取得權利範圍6",
            "序號7", "年月7", "地價7", "歷次取得權利範圍7",
            "序號8", "年月8", "地價8", "歷次取得權利範圍8",
            "序號9", "年月9", "地價9", "歷次取得權利範圍9",
            "序號10", "年月10", "地價10", "歷次取得權利範圍10",
            "序號11", "年月11", "地價11", "歷次取得權利範圍11",
            "序號12", "年月12", "地價12", "歷次取得權利範圍12",
            "序號13", "年月13", "地價13", "歷次取得權利範圍13",
            "序號14", "年月14", "地價14", "歷次取得權利範圍14",
            "序號15", "年月15", "地價15", "歷次取得權利範圍15",
            "序號16", "年月16", "地價16", "歷次取得權利範圍16",
            "序號17", "年月17", "地價17", "歷次取得權利範圍17",
            "序號18", "年月18", "地價18", "歷次取得權利範圍18",
            "序號19", "年月19", "地價19", "歷次取得權利範圍19",
            "序號20", "年月20", "地價20", "歷次取得權利範圍20",
            # ... 繼續到序號50
            
            # 其他登記事項放在最後
            "相關他項權利登記次序",
            "其他登記事項"
        ]
    
    def _get_building_indicator_field_order(self) -> List[str]:
        """最終修正版：建物標示部Excel欄位順序 - 支援更多停車位，按組別完整排列"""
        return [
            # 基本資訊
            "編號", "源檔案", "謄本類型", "地號建號", "列印時間", "謄本類別",
            
            # 標示部基本欄位（保持舊程式處理方式）
            "登記日期", "登記原因", "建物門牌", "建物坐落地號", 
            "主要用途", "主要建材", "層數", "總面積",
            
            # 主層次
            "層次", "層次面積",
            
            # 其他層次（動態生成，最多支援10層）
            "層次1", "層次面積1", "層次2", "層次面積2", "層次3", "層次面積3",
            "層次4", "層次面積4", "層次5", "層次面積5", "層次6", "層次面積6",
            "層次7", "層次面積7", "層次8", "層次面積8", "層次9", "層次面積9",
            "層次10", "層次面積10",
            
            # 建築完成日期在層次之後
            "建築完成日期",
            
            # 附屬建物（在建築完成日期之後）
            "附屬建物用途", "附屬建物面積",
            "附屬建物用途1", "附屬建物面積1", "附屬建物用途2", "附屬建物面積2",
            "附屬建物用途3", "附屬建物面積3", "附屬建物用途4", "附屬建物面積4",
            "附屬建物用途5", "附屬建物面積5",
            
            # 🔧 修正：共有部分按組別完整排列（支援更多停車位）
            # 第1組完整：建號 → 面積 → 權利範圍 → 所有停車位 → 其他登記事項
            "共有部分建號", "共有部分面積", "權利範圍", 
            
            # 🔧 關鍵修正：支援更多停車位（最多20個）
            "含停車位編號", 
            "含停車位編號_1_1", "含停車位編號_1_2", "含停車位編號_1_3", "含停車位編號_1_4",
            "含停車位編號_1_5", "含停車位編號_1_6", "含停車位編號_1_7", "含停車位編號_1_8",
            "含停車位編號_1_9", "含停車位編號_1_10", "含停車位編號_1_11", "含停車位編號_1_12",
            "含停車位編號_1_13", "含停車位編號_1_14", "含停車位編號_1_15", "含停車位編號_1_16",
            "含停車位編號_1_17", "含停車位編號_1_18", "含停車位編號_1_19", "含停車位編號_1_20",
            
            # 第1組其他登記事項
            "共有部分其他登記事項", "共有部分其他登記事項_1_1", "共有部分其他登記事項_1_2",
            
            # 第2組完整：建號 → 面積 → 權利範圍 → 所有停車位 → 其他登記事項
            "共有部分建號2", "共有部分面積2", "權利範圍2",
            "含停車位編號2", 
            "含停車位編號2_1", "含停車位編號2_2", "含停車位編號2_3", "含停車位編號2_4",
            "含停車位編號2_5", "含停車位編號2_6", "含停車位編號2_7", "含停車位編號2_8",
            "含停車位編號2_9", "含停車位編號2_10", "含停車位編號2_11", "含停車位編號2_12",
            "含停車位編號2_13", "含停車位編號2_14", "含停車位編號2_15", "含停車位編號2_16",
            "含停車位編號2_17", "含停車位編號2_18", "含停車位編號2_19", "含停車位編號2_20",
            "共有部分其他登記事項2", "共有部分其他登記事項2_1", "共有部分其他登記事項2_2",
            
            # 第3組完整
            "共有部分建號3", "共有部分面積3", "權利範圍3",
            "含停車位編號3", 
            "含停車位編號3_1", "含停車位編號3_2", "含停車位編號3_3", "含停車位編號3_4",
            "含停車位編號3_5", "含停車位編號3_6", "含停車位編號3_7", "含停車位編號3_8",
            "含停車位編號3_9", "含停車位編號3_10", "含停車位編號3_11", "含停車位編號3_12",
            "含停車位編號3_13", "含停車位編號3_14", "含停車位編號3_15", "含停車位編號3_16",
            "含停車位編號3_17", "含停車位編號3_18", "含停車位編號3_19", "含停車位編號3_20",
            "共有部分其他登記事項3", "共有部分其他登記事項3_1", "共有部分其他登記事項3_2",
            
            # 第4組完整
            "共有部分建號4", "共有部分面積4", "權利範圍4",
            "含停車位編號4", 
            "含停車位編號4_1", "含停車位編號4_2", "含停車位編號4_3", "含停車位編號4_4",
            "含停車位編號4_5", "含停車位編號4_6", "含停車位編號4_7", "含停車位編號4_8",
            "含停車位編號4_9", "含停車位編號4_10", "含停車位編號4_11", "含停車位編號4_12",
            "含停車位編號4_13", "含停車位編號4_14", "含停車位編號4_15", "含停車位編號4_16",
            "含停車位編號4_17", "含停車位編號4_18", "含停車位編號4_19", "含停車位編號4_20",
            "共有部分其他登記事項4", "共有部分其他登記事項4_1", "共有部分其他登記事項4_2",
            
            # 第5組完整
            "共有部分建號5", "共有部分面積5", "權利範圍5",
            "含停車位編號5", 
            "含停車位編號5_1", "含停車位編號5_2", "含停車位編號5_3", "含停車位編號5_4",
            "含停車位編號5_5", "含停車位編號5_6", "含停車位編號5_7", "含停車位編號5_8",
            "含停車位編號5_9", "含停車位編號5_10", "含停車位編號5_11", "含停車位編號5_12",
            "含停車位編號5_13", "含停車位編號5_14", "含停車位編號5_15", "含停車位編號5_16",
            "含停車位編號5_17", "含停車位編號5_18", "含停車位編號5_19", "含停車位編號5_20",
            "共有部分其他登記事項5", "共有部分其他登記事項5_1", "共有部分其他登記事項5_2",
            
            # 標示部整體的其他登記事項（最後）
            "其他登記事項", "其他登記事項1", "其他登記事項2", "其他登記事項3",
            "其他登記事項4", "其他登記事項5"
        ]

    def _get_rights_field_order(self) -> List[str]:
        """定義他項權利部的正確欄位順序"""
        return [
            # 基本資訊
            "編號", "源檔案", "地號建號",
            
            # 記錄資訊
            "記錄編號", "登記次序",
            
            # 權利基本資訊
            "權利種類", "收件年期", "字號", "登記日期", "登記原因",
            
            # 🔧 關鍵修正：權利人資訊的正確順序
            "權利人", "統一編號", "住址",  # 統一編號在權利人和住址之間
            
            # 債權資訊
            "債權額比例", "擔保債權總金額", "存續期間",
            
            # 其他欄位...
            "擔保債權種類及範圍", "擔保債權確定期日", "清償日期",
            "利息(率)", "遲延利息(率)", "違約金", "其他擔保範圍約定",
            "債務人及債務額比例", "權利標的", "標的登記次序", "設定權利範圍",
            "證明書字號", "設定義務人",
            "共同擔保地號", "共同擔保建號",
            "其他登記事項", "其他登記事項1", "其他登記事項2", "其他登記事項3",
            "其他登記事項4", "其他登記事項5"
        ]


    def _reorder_dataframe_columns(self, df: pd.DataFrame, field_order: List[str]) -> pd.DataFrame:
        """根據指定順序重新排列DataFrame的欄位"""
        
        # 獲取實際存在的欄位
        existing_columns = list(df.columns)
        
        # 按照指定順序排列存在的欄位
        ordered_columns = []
        for field in field_order:
            if field in existing_columns:
                ordered_columns.append(field)
        
        # 添加不在順序列表中但存在於DataFrame的欄位
        for col in existing_columns:
            if col not in ordered_columns:
                ordered_columns.append(col)
        
        return df[ordered_columns]

    def export_to_excel(self, data: List[Dict], filename: str = None):
        """輸出到Excel - 謄本專用格式 - 修正建物標示部欄位順序"""
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = self._create_flexible_filename(data, "xlsx")

        filepath = os.path.join(self.output_dir, filename)

        # 🔧 定義自動調整欄寬的輔助函數
        def auto_adjust_column_width(worksheet):
            """自動調整工作表的所有欄位寬度"""
            for column_cells in worksheet.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter

                for cell in column_cells:
                    try:
                        if cell.value:
                            cell_value = str(cell.value)
                            # 中文字元計為2個單位，英文計為1個單位
                            cell_length = sum(2 if ord(c) > 127 else 1 for c in cell_value)
                            max_length = max(max_length, cell_length)
                    except:
                        pass

                # 設定欄寬，增加邊距
                adjusted_width = max_length + 2
                adjusted_width = max(10, min(adjusted_width, 100))
                worksheet.column_dimensions[column_letter].width = adjusted_width

        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            sheets_created = 0

            print("📊 開始輸出Excel（修正欄位順序）...")
            
            # 1. 建物標示部（修正欄位順序）
            building_indicator_data = []
            for i, transcript in enumerate(data):
                if transcript is None or "建物謄本" not in str(transcript.get("謄本類型", "")):
                    continue
                    
                basic_info = {
                    "編號": i+1, 
                    "源檔案": transcript.get("源檔案", ""),
                    "謄本類型": transcript.get("謄本類型", "")
                }
                
                # 添加基本資訊
                basic_from_transcript = transcript.get("基本資訊", {})
                if basic_from_transcript:
                    excluded_fields = ["謄本類型", "申請人", "謄本種類碼", "機關名稱", "謄本類別"]
                    for key, value in basic_from_transcript.items():
                        if key not in excluded_fields:
                            basic_info[key] = value
                
                # 添加標示部資訊
                indicator = transcript.get("標示部", {})
                if indicator:
                    for key, value in indicator.items():
                        basic_info[key] = value
                
                building_indicator_data.append(basic_info)
            
            if building_indicator_data:
                df_building_indicator = pd.DataFrame(building_indicator_data)
                
                # 🔧 關鍵修正：重新排列欄位順序
                field_order = self._get_building_indicator_field_order()
                df_building_indicator = self._reorder_dataframe_columns(df_building_indicator, field_order)
                
                df_building_indicator.to_excel(writer, sheet_name='建物標示部', index=False)
                auto_adjust_column_width(writer.sheets['建物標示部'])  # 🔧 自動調整欄寬
                sheets_created += 1
                print(f"✅ 建物標示部工作表: 已創建({len(building_indicator_data)} 筆記錄)")
                
                # 調試：顯示欄位順序
                # print(f"   📋 欄位順序: {list(df_building_indicator.columns)}")
            
            # 2. 建物所有權部
            building_ownership_data = []
            for i, transcript in enumerate(data):
                if transcript is None or "建物謄本" not in str(transcript.get("謄本類型", "")):
                    continue
                    
                ownership_list = transcript.get("所有權部", [])
                if ownership_list is None:
                    ownership_list = []
                    
                for j, ownership in enumerate(ownership_list):
                    if ownership is None:
                        continue
                        
                    ownership_record = {
                        "編號": f"{i+1}-{j+1}",
                        "源檔案": transcript.get("源檔案", ""),
                        "地號建號": transcript.get("基本資訊", {}).get("地號建號", "")
                    }
                    
                    for key, value in ownership.items():
                        ownership_record[key] = value
                    
                    building_ownership_data.append(ownership_record)
            
            if building_ownership_data:
                df_building_ownership = pd.DataFrame(building_ownership_data)
                df_building_ownership.to_excel(writer, sheet_name='建物所有權部', index=False)
                auto_adjust_column_width(writer.sheets['建物所有權部'])  # 🔧 自動調整欄寬
                sheets_created += 1
                print(f"✅ 建物所有權部工作表: 已創建 ({len(building_ownership_data)} 筆記錄)")
            
            # 3. 建物他項權利部
            building_rights_data = []
            for i, transcript in enumerate(data):
                if transcript is None or "建物謄本" not in str(transcript.get("謄本類型", "")):
                    continue
                    
                rights_list = transcript.get("他項權利部", [])
                if rights_list is None:
                    rights_list = []
                    
                for j, rights in enumerate(rights_list):
                    if rights is None:
                        continue
                        
                    rights_record = {
                        "編號": f"{i+1}-{j+1}",
                        "源檔案": transcript.get("源檔案", ""),
                        "地號建號": transcript.get("基本資訊", {}).get("地號建號", "")
                    }
                    
                    for key, value in rights.items():
                        rights_record[key] = value
                    
                    building_rights_data.append(rights_record)
            
            if building_rights_data:
                df_building_rights = pd.DataFrame(building_rights_data)
                
                # 🔧 關鍵修正：重新排列他項權利部欄位順序
                rights_field_order = self._get_rights_field_order()
                df_building_rights = self._reorder_dataframe_columns(df_building_rights, rights_field_order)
                
                df_building_rights.to_excel(writer, sheet_name='建物他項權利部', index=False)
                auto_adjust_column_width(writer.sheets['建物他項權利部'])  # 🔧 自動調整欄寬
                sheets_created += 1
                print(f"✅ 建物他項權利部工作表: 已創建 ({len(building_rights_data)} 筆記錄)")

            # 4. 土地標示部
            land_indicator_data = []
            for i, transcript in enumerate(data):
                if transcript is None or "土地謄本" not in str(transcript.get("謄本類型", "")):
                    continue
                    
                basic_info = {
                    "編號": i+1, 
                    "源檔案": transcript.get("源檔案", ""),
                    "謄本類型": transcript.get("謄本類型", "")
                }
                
                basic_from_transcript = transcript.get("基本資訊", {})
                if basic_from_transcript:
                    excluded_fields = ["謄本類型", "申請人", "謄本種類碼", "機關名稱", "謄本類別"]
                    for key, value in basic_from_transcript.items():
                        if key not in excluded_fields:
                            basic_info[key] = value
                
                indicator = transcript.get("標示部", {})
                if indicator:
                    for key, value in indicator.items():
                        basic_info[key] = value
                
                land_indicator_data.append(basic_info)
            
            if land_indicator_data:
                df_land_indicator = pd.DataFrame(land_indicator_data)
                df_land_indicator.to_excel(writer, sheet_name='土地標示部', index=False)
                auto_adjust_column_width(writer.sheets['土地標示部'])  # 🔧 自動調整欄寬
                sheets_created += 1
                print(f"✅ 土地標示部工作表已創建 ({len(land_indicator_data)} 筆記錄)")
            
            # 5. 土地所有權部
            land_ownership_data = []
            for i, transcript in enumerate(data):
                if transcript is None or "土地謄本" not in str(transcript.get("謄本類型", "")):
                    continue
                    
                ownership_list = transcript.get("所有權部", [])
                if ownership_list is None:
                    ownership_list = []
                    
                for j, ownership in enumerate(ownership_list):
                    if ownership is None:
                        continue
                        
                    ownership_record = {
                        "編號": f"{i+1}-{j+1}",
                        "源檔案": transcript.get("源檔案", ""),
                        "地號建號": transcript.get("基本資訊", {}).get("地號建號", "")
                    }
                    
                    for key, value in ownership.items():
                        ownership_record[key] = value
                    
                    land_ownership_data.append(ownership_record)
            
            # 5. 土地所有權部
            if land_ownership_data:
                df_land_ownership = pd.DataFrame(land_ownership_data)
                
                # 🔧 關鍵修正：重新排列土地所有權部欄位順序
                land_ownership_field_order = self._get_land_ownership_field_order()
                df_land_ownership = self._reorder_dataframe_columns(df_land_ownership, land_ownership_field_order)
                
                df_land_ownership.to_excel(writer, sheet_name='土地所有權部', index=False)
                auto_adjust_column_width(writer.sheets['土地所有權部'])  # 🔧 自動調整欄寬
                sheets_created += 1
                print(f"✅ 土地所有權部工作表已創建 ({len(land_ownership_data)} 筆記錄)")

            # 6. 土地他項權利部
            land_rights_data = []
            for i, transcript in enumerate(data):
                if transcript is None or "土地謄本" not in str(transcript.get("謄本類型", "")):
                    continue
                    
                rights_list = transcript.get("他項權利部", [])
                if rights_list is None:
                    rights_list = []
                    
                for j, rights in enumerate(rights_list):
                    if rights is None:
                        continue
                        
                    rights_record = {
                        "編號": f"{i+1}-{j+1}",
                        "源檔案": transcript.get("源檔案", ""),
                        "地號建號": transcript.get("基本資訊", {}).get("地號建號", "")
                    }
                    
                    for key, value in rights.items():
                        rights_record[key] = value
                    
                    land_rights_data.append(rights_record)
            
            if land_rights_data:
                df_land_rights = pd.DataFrame(land_rights_data)
                
                # 🔧 關鍵修正：重新排列他項權利部欄位順序
                rights_field_order = self._get_rights_field_order()
                df_land_rights = self._reorder_dataframe_columns(df_land_rights, rights_field_order)
                
                df_land_rights.to_excel(writer, sheet_name='土地他項權利部', index=False)
                auto_adjust_column_width(writer.sheets['土地他項權利部'])  # 🔧 自動調整欄寬
                sheets_created += 1
                print(f"✅ 土地他項權利部工作表已創建 ({len(land_rights_data)} 筆記錄)")
            
            # 如果沒有創建任何工作表，創建摘要
            if sheets_created == 0:
                summary_data = [{
                    "檔案名稱": data[0].get("源檔案", "未知") if data else "無資料",
                    "謄本數量": len(data),
                    "狀態": "謄本解析完成，但無有效數據" if data else "無資料",
                    "說明": "謄本可能沒有他項權利部，或數據解析遇到問題"
                }]
                
                df_summary = pd.DataFrame(summary_data)
                df_summary.to_excel(writer, sheet_name='處理摘要', index=False)
                auto_adjust_column_width(writer.sheets['處理摘要'])  # 🔧 自動調整欄寬
                sheets_created += 1
        
        # print(f"📈 謄本Excel 檔案已輸出（修正欄位順序）: {filepath}")
        return filepath

    def export_to_json(self, data: List[Dict], filename: str = None):
        """輸出到JSON"""
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = self._create_flexible_filename(data, "json")
        
        filepath = os.path.join(self.output_dir, filename)
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        # print(f"📄 謄本JSON 檔案已輸出: {filepath}")
        return filepath

    def export_to_txt_report(self, data: List[Dict], filename: str = None):
        """生成謄本專用的詳細txt報告"""
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = self._create_flexible_filename(data, "txt")
        
        filepath = os.path.join(self.output_dir, filename)
        
        with open(filepath, 'w', encoding='utf-8') as f:
            # 直接在寫入前進行內容清理
            print("🔧 執行TXT最終清理...")

            f.write("="*80 + "\n")
            f.write("📋 登記謄本詳細分析報告\n")
            f.write("="*80 + "\n")
            f.write(f"生成時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"總謄本數: {len(data)}\n")
            
            # 統計各類型謄本數量
            second_class_count = sum(1 for t in data if "第二類" in str(t.get("謄本類型", "")))
            first_class_count = sum(1 for t in data if "第一類" in str(t.get("謄本類型", "")))
            building_count = sum(1 for t in data if "建物謄本" in str(t.get("謄本類型", "")))
            land_count = sum(1 for t in data if "土地謄本" in str(t.get("謄本類型", "")))
            
            f.write(f"第二類謄本: {second_class_count} 份, 第一類謄本: {first_class_count} 份\n")
            f.write(f"建物謄本: {building_count} 份, 土地謄本: {land_count} 份\n\n")
            
            # 謄本特色說明
            f.write("🛠️ 謄本解析特色（修正版）:\n")
            f.write("  ✅ 支援多筆他項權利記錄的完整解析\n")
            f.write("  ✅ 正確處理縮排的擔保債權種類及範圍內容\n")
            f.write("  ✅ 精確分割各欄位邊界，避免內容串接\n")
            f.write("  ✅ 完整提取共同擔保地號和建號列表\n")
            f.write("  ✅ 強化跨頁續次頁內容處理\n")
            f.write("  🔧 修正建物標示部Excel輸出欄位順序\n")
            f.write("     ↳ 建築完成日期現在正確顯示在層次面積之後\n\n")
            
            # 按源檔案分組
            file_groups = {}
            for transcript in data:
                file_name = transcript.get("源檔案", "未知檔案")
                if file_name not in file_groups:
                    file_groups[file_name] = []
                file_groups[file_name].append(transcript)
            
            for file_name, transcripts in file_groups.items():
                f.write(f"{'='*60}\n")
                f.write(f"📁 檔案: {file_name} ({len(transcripts)} 份謄本)\n")
                f.write(f"{'='*60}\n")
                
                for i, transcript in enumerate(transcripts, 1):
                    f.write(f"\n📄 第 {i} 份謄本 - {transcript.get('謄本類型', '未知類型')}\n")
                    f.write(f"{'-'*40}\n")
                    
                    # 基本資訊
                    basic = transcript.get("基本資訊", {})
                    f.write("📌 基本資訊:\n")
                    for key, value in basic.items():
                        f.write(f"  {key}: {value}\n")
                    
                    # 標示部
                    indicator = transcript.get("標示部", {})
                    if indicator:
                        f.write("\n📍 標示部:\n")
                        
                        # 🔧 重要：按正確順序顯示建物標示部資訊
                        if "建物謄本" in str(transcript.get("謄本類型", "")):
                            field_order = self._get_building_indicator_field_order()
                            # 只顯示實際存在的欄位，按順序
                            for field in field_order:
                                if field in indicator:
                                    f.write(f"  {field}: {indicator[field]}\n")
                            
                            # 顯示不在順序列表中的其他欄位
                            for key, value in indicator.items():
                                if key not in field_order:
                                    f.write(f"  {key}: {value}\n")
                        else:
                            # 土地謄本正常顯示
                            for key, value in indicator.items():
                                f.write(f"  {key}: {value}\n")
                    
                    # 所有權部
                    ownership_list = transcript.get("所有權部", [])
                    if ownership_list:
                        f.write(f"\n👤 所有權部 ({len(ownership_list)} 筆記錄):\n")
                        for j, ownership in enumerate(ownership_list, 1):
                            f.write(f"  記錄 {j}:\n")
                            for key, value in ownership.items():
                                f.write(f"    {key}: {value}\n")
                    
                    # 他項權利部
                    rights_list = transcript.get("他項權利部", [])
                    if rights_list:
                        f.write(f"\n⚖️ 他項權利部 ({len(rights_list)} 筆記錄):\n")
                        for j, rights in enumerate(rights_list, 1):
                            f.write(f"  記錄 {j} (登記次序: {rights.get('登記次序', '未知')}):\n")
                            
                            # 🔧 修正：按照資料原本順序輸出，不分重點和其他
                            for key, value in rights.items():
                                if len(str(value)) > 80:
                                    f.write(f"    {key}:\n")
                                    # 對長內容進行適當的縮排顯示
                                    lines = str(value).split('\n')
                                    for line in lines:
                                        f.write(f"      {line}\n")
                                else:
                                    f.write(f"    {key}: {value}\n")
                    
                    f.write("\n")
                
                f.write("\n")
        
        # print(f"📑 謄本完整報告已輸出（修正版）: {filepath}")
        return filepath

    def finalize(self):
        """
        完成所有匯出後的最終處理
        - 保存完整的除錯日誌
        """
        try:
            log_file = os.path.join(self.output_dir, "transcript_debug.txt")
            debug_logger.save_logs(log_file)
            print(f"📄 完整除錯日誌已保存到: {log_file}")
        except Exception as e:
            print(f"⚠️ 日誌保存失敗: {e}")

# ==================== 主程式 ====================
def main():
    #     with open("ocr_debug.txt", "w", encoding="utf-8") as f:
    #         f.write("=== OCR調試開始 ===\n")
    
    # 🔧 日誌記錄已在模組匯入時啟動，不需要重複呼叫
    # debug_logger.start_logging()  # ← 已移到模組層級
    print("🔧 不動產登記謄本通用解析器（修正Excel欄位順序版）")
    print("="*50)
    print("專門處理土地登記謄本和建物登記謄本（支援第一類及第二類）")
    print("特色功能:")
    print("  ✅ 支援第一類謄本的完整資料顯示")
    print("  ✅ 支援第二類謄本的隱私保護格式")
    print("  ✅ 自動識別謄本類型並採用對應解析策略")
    print("  ✅ 完整解析他項權利記錄（支援多筆記錄）")
    print("  ✅ 處理共同擔保地號和建號資訊")
    print("  ✅ 增強跨頁續次頁處理")
    print("  ✅ 優化長列表資料解析")
    print("  ✅ 統一OCR資料補充邏輯（兼容一類+二類）")
    print("  🔧 修正建物標示部Excel欄位順序")
    print("     ↳ 建築完成日期現在正確顯示在層次面積之後")
    print("="*50)
    # 指定要處理的PDF檔案
    specified_pdfs = [
        # "建A+B+C.pdf",
        # "地A+B+C.pdf",
        # "地-585-4地號A+B+C.pdf",
        # "路地A+B+C.pdf",
        # "土地登記第二類謄本-大寮區翁公園段二小段 5077-0000地號.pdf",
        # "土地一類-小港區青島段二小段 0190-0002地號.pdf",
        # "土地一類-林園區王公廟段苦苓腳小段 1183-0000等11筆地號.pdf",---
        # "土地一類-高樹鄉源興段 0291-0000地號.pdf",
        # "鳳山區頂新段605建號-一類.pdf",
        "建物一類-前鎮區興邦段 00150-000建號.pdf",
        # "土地一類+建物-左營區新民段181地號+3232建號.pdf"
        # "土地一類-多人.pdf",
        # "土地一類-多次多人.pdf",
        # "113-土地+建物-二類.pdf",
        # "土地登記第二類謄本-安平區漁光段 0891-0002地號+安平區漁光段 01632-000、01661-000建號.pdf",
        # "土地一類-林園區王公廟段苦苓腳小段 1182-0000地號.pdf",
        # "土地+建物-橋頭區經武路50號22樓-1.pdf",
        # ".\\1140327許智豪謄本\\1140327鳳山區文山段847地號.pdf",
        # ".\\1140327許智豪謄本\\1140327鳳山區鳳青段169.170.171.172地號.pdf",
        # "土地二類-鳥松區大脚腿段 1780-0016地號.pdf",
        # "土地登記第二類謄本-左營區廍後段 0019-0003地號+左營區廍後段 01447-000建號.pdf",
        # "土地登記第二類謄本-大寮區磚子磘段 3800-0000地號.pdf",
    ]
    
    # 檢查指定檔案是否存在
    existing_pdfs = []
    for pdf in specified_pdfs:
        if os.path.exists(pdf):
            existing_pdfs.append(pdf)
            print(f"✅ 找到檔案: {pdf}")
        else:
            print(f"❌ 檔案不存在: {pdf}")
    
    if not existing_pdfs:
        print("\n❌ 沒有找到任何指定的PDF檔案!")
        return
    
    print(f"\n📋 將處理 {len(existing_pdfs)} 個PDF檔案:")
    for i, pdf in enumerate(existing_pdfs, 1):
        print(f"  {i}. {pdf}")
    
    confirm = input("\n繼續處理? (Y/n): ").lower().strip()
    if confirm == 'n' or confirm == 'no':
        print("❌ 處理已取消")
        return

    # 初始化處理器
    processor = UnifiedTranscriptProcessor()
    exporter = UnifiedTranscriptExporter()
    
    print("\n🚀 開始謄本解析（通用版）...")
    
    # 處理所有PDF檔案
    all_results = processor.process_multiple_pdfs(existing_pdfs)
    
    # 統計結果
    print(f"\n📊 處理完成統計:")
    print(f"  總謄本數: {len(all_results)}")
    
    # 統計類型
    first_class_count = sum(1 for r in all_results if "第一類" in str(r.get("謄本類型", "")))
    second_class_count = sum(1 for r in all_results if "第二類" in str(r.get("謄本類型", "")))

    print(f"  第一類謄本: {first_class_count} 份")
    print(f"  第二類謄本: {second_class_count} 份")

    
    # 輸出結果
    print(f"\n📊 開始輸出謄本結構化資料（修正版）...")
    

    # 🔍 檢查補字前的狀態
    print(f"\n🔍 補字階段檢查:")
    _glyphs_check = globals().get('inline_glyphs', [])
    print(f"   全域內嵌字元: {_glyphs_check}")

    # ⚙️ 最後補字：在輸出前對所有結果做遞迴補字（涵蓋住址/地段/地號等）
    _self = locals().get('self')
    _locals_glyphs = locals().get('inline_glyphs')
    _globals_glyphs = globals().get('inline_glyphs')

    print(f"   🔍 _self: {_self}")
    print(f"   🔍 locals中的glyphs: {_locals_glyphs}")
    print(f"   🔍 globals中的glyphs: {_globals_glyphs}")

    _glyphs = ((_self and getattr(_self, 'inline_glyphs', None)) or _locals_glyphs or _globals_glyphs)
    print(f"   🔍 最終選用的_glyphs: {_glyphs}")

    if _glyphs:
        print(f"   🧩 開始執行deep_patch_glyphs，字元: {_glyphs}")
        try:
            all_results = deep_patch_glyphs(all_results, _glyphs)
            print(f"   ✅ deep_patch_glyphs執行完成")
            # 加入全域地上建物建號清理
            def global_cleanup_building_numbers(data):
                """全域清理重複的地上建物建號 + 修正缺失的廍字（支援多種格式）"""
                import re
                
                # 🔧 先檢查共同擔保地號是否存在
                print("檢查共同擔保地號:")
                data_str = str(data)
                if "共同擔保" in data_str:
                    print("  發現共同擔保相關內容")
                else:
                    print("  未發現共同擔保相關內容")
                
                if "磚子" in data_str:
                    print("  發現磚子相關內容")
                else:
                    print("  未發現磚子相關內容")
                
                def clean_recursive(obj):
                    if isinstance(obj, dict):
                        cleaned = {}
                        for key, value in obj.items():
                            if isinstance(value, str):
                                original = value
                                cleaned_value = value
                                
                                # 🔧 保護共同擔保內容，避免被清理
                                if "共同擔保" in str(value):
                                    print(f"    🛡️ 跳過共同擔保內容清理: {key}")
                                    cleaned[key] = cleaned_value
                                    continue
                                
                                # 原有的重複清理（支援地號和建號）
                                if key == "地上建物建號":
                                    cleaned_value = re.sub(r'後段\s+(\d{4,5}-\d{3,4})\s*廍?後段\s+\1', r'廍後段 \1', cleaned_value)
                                
                                # 檢測是否包含地號或建號相關內容
                                has_land_building_ref = any([
                                    re.search(r'\d{4,5}-\d{3,4}', cleaned_value),  # 標準格式：0019-0003, 01311-000
                                    re.search(r'１９之３地號', cleaned_value),        # 特殊格式：１９之３地號  
                                    re.search(r'\d+之\d+地號', cleaned_value),       # 通用格式：數字之數字地號
                                    re.search(r'地號|建號', cleaned_value)            # 包含地號或建號關鍵字
                                ])
                                
                                if has_land_building_ref:
                                    # 修正各種「後段」格式為「廍後段」
                                    patterns = [
                                        # 格式1: 後段 + 標準地號建號 (0019-0003, 01311-000)
                                        (r'(?<!廍)後段\s+(\d{4,5}-\d{3,4})', r'廍後段 \1'),
                                        
                                        # 格式2: 後段 + 中文數字地號 (１９之３地號)
                                        (r'(?<!廍)後段\s*([１２３４５６７８９０\d]+之[１２３４５６７８９０\d]+地號)', r'廍後段\1'),
                                        
                                        # 格式3: 後段 + 一般數字地號 (19之3地號)
                                        (r'(?<!廍)後段\s*(\d+之\d+地號)', r'廍後段\1'),
                                        
                                        # 格式4: 單純的後段（如果上下文包含地號建號）
                                        (r'(?<!廍)後段(?=\s|$|，|。)', r'廍後段'),
                                        
                                        # 格式5: 磘字修正 - 磚子 段 → 磚子磘段
                                        (r'磚子\s+段', r'磚子磘段')
                                    ]
                                    
                                    for pattern, replacement in patterns:
                                        old_value = cleaned_value
                                        cleaned_value = re.sub(pattern, replacement, cleaned_value)
                                        if cleaned_value != old_value:
                                            break  # 只執行第一個匹配的規則，避免過度替換
                                
                                if cleaned_value != original:
                                    print(f"    🔧 修正: {key} - {original[:50]}... → {cleaned_value[:50]}...")
                                
                                cleaned[key] = cleaned_value
                            else:
                                cleaned[key] = clean_recursive(value)
                        return cleaned
                    elif isinstance(obj, list):
                        return [clean_recursive(item) for item in obj]
                    return obj
                
                return clean_recursive(data)

            print(f"\n🔧 執行全域地上建物建號清理...")
            all_results = global_cleanup_building_numbers(all_results)
            print(f"✅ 全域建號清理完成")
        except Exception as _e:
            print(f"   ⚠️ 最後補字失敗: {_e}")
            import traceback
            traceback.print_exc()
    else:
        print(f"   ❌ _glyphs為空，跳過補字")
    # 🔧 新增：最終全域清理十六進制地段建號殘留（完全移除）
    def final_cleanup_hex_geo(obj):
        """通用清理所有十六進制字母 + 地段建號的殘留"""
        if isinstance(obj, str):
            original = obj
            result = obj
            
            # 🔧 擴展保護共同擔保相關內容
            protection_patterns = [
                "共同擔保地號",
                "共同擔保建號", 
            ]

            for pattern in protection_patterns:
                if pattern in result:
                    print(f"    🛡️ 跳過清理受保護內容: {pattern} - {original[:50]}...")
                    return result

            # 🔧 新增：保護包含多個地號的磚子磘段內容
            if re.search(r'磚子磘段.*?\d{4}-\d{4}.*?\d{4}-\d{4}', result):
                print(f"    🛡️ 跳過清理共同擔保地號內容: {original[:50]}...")
                return result
            
            # 通用清理模式（適用於所有謄本）- 修正版
            cleanup_patterns = [
                # 模式1: 清理 "字母 + 地區地段 + 地號/建號" (支援各種格式)
                r'\b[A-HG]\s+[^，。\n]*?[區鄉鎮市][^，。\n]*?段(?:[^，。\n]*?小段)?[^，。\n]*?\d{4,5}-\d{3,4}[地建]號',
                
                # 模式2: 清理 "字母 + 地區地段" (沒有地號，支援小段)
                r'\b[A-HG]\s+[^，。\n]*?[區鄉鎮市][^，。\n]*?段(?:[^，。\n]*?小段)?(?!\d)',
                
                # 模式3: 清理孤立的 "地段 + 地號/建號" (支援各種格式)
                r'(?<=\s)[^，。\n]*?[區鄉鎮市][^，。\n]*?段(?:[^，。\n]*?小段)?\s+\d{4,5}-\d{3,4}[地建]號(?=\s|$|，|。)',
                
                # 模式4: 清理換行後的地段建號 (支援小段)
                r'\n[A-HG]\s+[^，。\n]*?[區鄉鎮市][^，。\n]*?段(?:[^，。\n]*?小段)?[^，。\n]*?',
                
                # 模式5: 更寬鬆的地號建號格式 (1-5位數字)
                r'\b[A-HG]\s+[^，。\n]*?[區鄉鎮市][^，。\n]*?段(?:[^，。\n]*?小段)?[^，。\n]*?\d{1,5}-\d{1,5}[地建]號',
            ]
            
            cleaned_count = 0
            for i, pattern in enumerate(cleanup_patterns, 1):
                old_result = result
                result = re.sub(pattern, '', result, flags=re.MULTILINE)
                if result != old_result:
                    cleaned_count += 1
                    print(f"    🗑️ 模式{i}清理成功")
            
            if cleaned_count > 0:
                # 清理多餘的空格和標點
                result = re.sub(r'\s+', ' ', result)  # 多個空格變一個
                result = re.sub(r'，\s*，', '，', result)  # 重複逗號
                result = re.sub(r'。\s*。', '。', result)  # 重複句號
                result = re.sub(r'：\s+', '：', result)  # 冒號後多餘空格
                result = result.strip()
                
                print(f"    ✅ 清理前: {original[:80]}...")
                print(f"    ✅ 清理後: {result[:80]}...")
            
            return result
        elif isinstance(obj, dict):
            return {k: final_cleanup_hex_geo(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [final_cleanup_hex_geo(item) for item in obj]
        return obj

    print(f"\n🗑️ 執行通用十六進制地段建號清理...")
    cleaned_count = 0
    for i, result in enumerate(all_results):
        old_result_str = str(result)
        all_results[i] = final_cleanup_hex_geo(result)
        new_result_str = str(all_results[i])
        if old_result_str != new_result_str:
            cleaned_count += 1

    print(f"\n🗑️ 執行最終十六進制地段建號完全清除...")
    all_results = final_cleanup_hex_geo(all_results)
    print(f"✅ 十六進制地段建號完全清除完成")

    
    # JSON 輸出
    json_file = exporter.export_to_json(all_results)
    
    # Excel 輸出（修正欄位順序）
    excel_file = exporter.export_to_excel(all_results)
    
    # TXT 報告輸出
    txt_file = exporter.export_to_txt_report(all_results)
    
    # 🔧 可選：住址校正檢查
    print(f"\n🔍 住址配對檢查:")
    address_issues = []
    
    for transcript in all_results:
        ownership_list = transcript.get("所有權部", [])
        for i, owner in enumerate(ownership_list):
            addr = owner.get("住址", "")
            owner_id = owner.get("統一編號", "")
            
            # 🔧 修正：正確的住址檢查邏輯
            if addr:
                # 檢查是否有真正的重複或錯誤模式
                problem_patterns = [
                    # 模式1：同一類型重複（如：台北市台北市）
                    r'(.+市)\1',
                    r'(.+縣)\1',
                    
                    # 模式2：不合理的組合（如：台北市高雄市）
                    r'台北市.*?高雄市|高雄市.*?台北市',
                    r'台中市.*?台南市|台南市.*?台中市',
                    
                    # 模式3：超過2個不同的縣市名稱
                    # 但排除正常的縣轄市格式
                ]
                
                has_problem = False
                for pattern in problem_patterns:
                    if re.search(pattern, addr):
                        has_problem = True
                        break
                
                # 🔧 特別檢查：排除正常的縣轄市格式
                if re.search(r'高雄縣鳳山市|台北縣.*?市|桃園縣.*?市', addr):
                    has_problem = False  # 這些是正常格式
                
                if has_problem:
                    issue = {
                        "檔案": transcript.get("源檔案", ""),
                        "所有權人": owner.get("所有權人", ""),
                        "統一編號": owner_id,
                        "住址": addr,
                        "問題": "住址格式異常"
                    }
                    address_issues.append(issue)
    
    if address_issues:
        print(f"⚠️ 發現 {len(address_issues)} 個可疑的住址配對:")
        for issue in address_issues:
            print(f"  📄 {issue['檔案']}")
            print(f"     所有權人: {issue['所有權人']} ({issue['統一編號']})")
            print(f"     住址: {issue['住址']}")
            print(f"     問題: {issue['問題']}")
            print()
        
        print(f"💡 建議:")
        print(f"  1. 檢查OCR圖片順序是否與PDF中的所有權人順序一致")
        print(f"  2. 考慮手動校正上述住址")
        print(f"  3. 如果經常出現此問題，可能需要調整OCR配對邏輯")
    else:
        print(f"✅ 所有住址配對看起來都正常")

    print(f"📁 輸出目錄: {exporter.output_dir}")
    print(f"📄 檔案清單:")
    print(f"   - JSON:  {os.path.basename(json_file)}")
    print(f"   - Excel: {os.path.basename(excel_file)}")
    print(f"   - TXT:   {os.path.basename(txt_file)}")

    # 🔧 完成處理，保存除錯日誌
    exporter.finalize()

if __name__ == "__main__":
    main()