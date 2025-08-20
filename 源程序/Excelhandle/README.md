# Excel Data Analyst

Excel数据分析师是一款基于Python和PyQt5开发的桌面应用程序，用于处理和分析Excel文件。

## 功能特性

- 导入Excel文件
- 提取指定列数据
- 定向筛选数据
- 多表合并
- 多表统计排行
- 生成图表

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法
```bash
python dataEXCEL.py
```

## 打包可执行文件（方式一 直接打包）
> 可以直接使用PyInstaller命令打包，但需要包含必要的资源文件.

1. 安装PyInstaller：
```bash
pip install pyinstaller
```
2. Windows系统：
```bash
pyinstaller --onefile --windowed --add-data "image/*.png;image" dataEXCEL.py
```
3. Macos系统：
```bash
pyinstaller --onefile --windowed --add-data "image/*.png:image" dataEXCEL.py
```

## 打包可执行文件（方式二 使用.spec文件）
1. 安装PyInstaller：
```bash
pip install pyinstaller
```
2. 创建.spec文件
> 首先运行以下命令生成基本的.spec文件
```bash
pyinstaller --name=ExcelDataAnalyst --windowed ./dataEXCEL.py
```
3. 修改.spec文件
> 编辑生成的ExcelDataAnalyst.spec文件，添加资源文件路径
```python
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['dataEXCEL.py'],
    pathex=[],
    binaries=[],
    datas=[('image/*.png', 'image')],  # 添加图片资源
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
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
    name='ExcelDataAnalyst',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Windows下设为False以避免显示命令行窗口
    disable_windowed_traceback=False,
    argv_emulation=False,  # Mac下可设为True以处理文件关联
    target_arch=None,  # Mac下可指定'x86_64'或'arm64'
    codesign_identity=None,
    entitlements_file=None,
)
```
4. 编译为Windows可执行文件，在Windows系统上运行以下命令：
```bash
pyinstaller ExcelDataAnalyst.spec
# 或者
pyinstaller --onefile --windowed ExcelDataAnalyst.spec
```
5. 编译为Mac可执行文件，在Mac系统上运行以下命令：
```bash
pyinstaller ExcelDataAnalyst.spec
# 如果需要指定架构，可以使用：
pyinstaller --target-architecture x86_64 ExcelDataAnalyst.spec  # Intel Mac
pyinstaller --target-architecture arm64 ExcelDataAnalyst.spec   # Apple Silicon Mac
```
6. 创建.app bundle (Mac专用)
> 对于Mac系统，还可以创建标准的应用程序包：
```bash
pyinstaller --windowed --name=ExcelDataAnalyst --add-data "image/*.png:image" --osx-bundle-identifier=com.yourcompany.ExcelDataAnalyst dataEXCEL.py
```
7. 优化建议
为了减小可执行文件的大小和避免潜在问题，建议添加以下选项：
```bash
pyinstaller --onefile --windowed --name=ExcelDataAnalyst --add-data "image/*.png:image" --exclude-module tkinter --exclude-module matplotlib.backends._backend_tk dataEXCEL.py
```