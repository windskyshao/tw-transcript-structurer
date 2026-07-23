# -*- mode: python ; coding: utf-8 -*-
# 完整版 - 包含所有相依模組和套件
# 包括 cv2 (OpenCV) 和 numpy

import os
import rapidocr_onnxruntime
_rapidocr_path = os.path.dirname(rapidocr_onnxruntime.__file__)

block_cipher = None

a = Analysis(
    ['GUI_transcript_pdf 1141021-01.py'],
    pathex=[],
    binaries=[],
    datas=[
        # 本地相依模組
        ('base_dir_helper.py', '.'),
        ('pdf_parser.py', '.'),

        # RapidOCR 模型檔案和資料（動態取得套件路徑，不寫死）
        (os.path.join(_rapidocr_path, 'models'), 'rapidocr_onnxruntime/models'),
        (os.path.join(_rapidocr_path, 'config.yaml'), 'rapidocr_onnxruntime'),

        # 🔥 如果您還有其他功能模組,請在這裡加入:
        # ('export_module.py', '.'),
        # ('其他模組.py', '.'),
    ],
    hiddenimports=[
        # 🔥 僅保留必要的核心模組
        'rapidocr_onnxruntime',
        'onnxruntime',
        'pyclipper',

        # PDF 和圖像處理 (程式實際需要)
        'fitz',
        'PIL',

        # 本地模組
        'base_dir_helper',
        'pdf_parser',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 排除不需要的大型套件
        'torch',
        'torch.utils',
        'torchvision',
        'tensorflow',
        'tensorflow-cpu',
        'scipy',
        'matplotlib',
        'sklearn',
        'scikit-learn',
        'scikit-image',
        # 🔥 排除不必要的資料庫和語法高亮套件（優化啟動速度）
        'sqlalchemy',           # 資料庫 ORM (不需要)
        'pygments',             # 語法高亮 (不需要)
        'lxml',                 # XML 處理 (pandas 可用其他引擎)
        'cryptography',         # 加密庫 (大部分功能不需要)
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    # 🔥 noarchive=False 可減小檔案大小（壓縮 Python 模組）
    # 啟動速度會稍慢（需要解壓縮）,但檔案會小很多
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='電子謄本結構化v1.7',
    debug=False,
    bootloader_ignore_signals=False,
    # 🔥 strip=True 可移除除錯符號,減小檔案大小並稍微加快啟動
    strip=True,
    # 🔥 upx=True 壓縮可執行檔,但可能略微增加啟動時間
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 改為 True 可以看到除錯訊息
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
