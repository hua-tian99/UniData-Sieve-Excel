# 📊 UniData-Sieve-Excel

UniData-Sieve 是一个**从海量、格式杂乱的 Excel / ZIP 中精确筛选并汇总数据**的小工具。  
典型用途包括：按学号/姓名/班级，从一堆老师上交的 Excel 和压缩包里，批量捞出某一批学生的数据，并可选地给每个人拆一份独立表格。

本仓库同时面向两类用户：

- 🧑‍💻 普通用户：直接使用打包好的 **Windows 免安装版 exe**，不用管 Python 和代码。
- 👨‍🔧 技术人员：基于源码进行本地运行、调试和二次开发。

---

## 一、普通用户：免安装版 exe 使用说明

### 1. 获取程序

1. 打开仓库主页：  
   `https://github.com/hua-tian99/UniData-Sieve-Excel`
2. 点击右侧或上方的 **Releases**（版本发布）入口。
3. 在最新的 Release 中下载 `UniData-Sieve.exe`（文件名以实际为准），**以及它所在的整个文件夹**。

> 如果暂时没有 Release，请联系项目维护者获取打包好的完整文件夹。

---

### 2. 绝对禁忌：不要单独拷 exe

你拿到的通常是一个文件夹，里面至少包含：

- `UniData-Sieve.exe`
- `_internal`（或类似名字的内部组件目录）
- 以及其他若干运行所需文件

⚠️ 请务必遵守：

- **错误做法**：  
  把 `UniData-Sieve.exe` 单独剪切/复制到桌面运行（会因缺少内部组件而直接闪退）。
- **正确做法**：  
  保持整个大文件夹结构不变，可以把**整个文件夹**移动到你喜欢的位置。  
  如果想在桌面快捷启动：  
  右键 `UniData-Sieve.exe` → `发送到` → `桌面快捷方式`。

---

### 3. 启动步骤

1. 在该完整文件夹内，双击运行 `UniData-Sieve.exe`。
2. 会弹出一个**黑色命令行窗口**（请不要关掉它）。
3. 等待约 3–5 秒，程序会自动打开你的默认浏览器（推荐 Edge / Chrome），访问：  
   `http://127.0.0.1:8501`
4. 浏览器中会出现 UniData-Sieve 的控制面板。

> 只要你还在使用网页，这个黑色窗口就必须保持打开。  
> 用完后直接关闭黑色窗口即可退出所有服务，网页也会随之失效。

---

### 4. 基本操作流程

网页面板大致分为：

- 左侧：**控制区**（配置匹配规则、列筛选等）
- 右侧：**工作区**（上传文件、查看进度和结果）

#### 4.1 设置过滤规则（左侧）

1. **提取关键词**  
   - 输入你要查找的关键字，支持多个，例如：  
     `24级软件1班` 或一串学号：  
     `24405101001 24405101002 24405101003`  
   - 词与词之间用**空格**分隔。

2. **匹配模式**（在界面中选择）  
   - `AND`：一行中必须同时包含你输入的**所有**词才算命中。  
   - `OR`：一行只要包含其中**任意一个**词就会被抓取（常用于成批学号）。  
   - `REGEX`：高级正则模式，仅推荐熟悉正则表达式的用户使用。

3. **列截取/列模式**  
   - 如果原始表格列非常多，你只关心“学号 / 姓名 / 班级”等核心字段，可以选择较精简的列模式；
   - 程序会自动裁掉明显无关的空列或重复列，保持结果表简洁。

#### 4.2 上传数据源（右侧）

- 支持的输入格式：
  - 单个或多个 `.xlsx` / `.xls` 文件；
  - 含有这些表格的 `.zip` 压缩包（支持嵌套压缩包，最多约 5 层）。
- 你可以直接把**一大堆表格和压缩包一起拖拽**到上传区域；
- 程序会自动：
  - 递归解压压缩包；
  - 解决大多数中文文件名乱码问题；
  - 找出所有可以处理的 Excel 表格。

#### 4.3 一键处理

点击按钮（如“开始处理”或界面上的启动按钮），程序会开始：

1. 递归解压 ZIP；
2. 逐行扫描 Excel；
3. 用你设定的关键词和模式进行匹配；
4. 聚合所有命中的结果行。

数据量越大，耗时越久。期间你可以看到右侧的日志/进度反馈。

---

### 5. 查看和导出结果

#### 5.1 汇总大表

- 处理完成后，页面中间会展示**前若干行结果预览**（如 100 行）。
- 确认结果无误后，点击“下载结果汇总 Excel”等按钮，保存完整的结果表，一般类似：  
  `result.xlsx`

#### 5.2 智能拆分打包（按人生成多份文件）

如果你想把汇总结果按人拆分，每个人单独一份 Excel（适合发给学生确认）：

1. 在页面底部找到 **“智能拆分/打包”** 模块。
2. 准备一份“花名册” Excel：只需 **一列** 如“学号”或“姓名”。（注意：表头需要在第一列）
3. 上传这份花名册。
4. 程序会：
   - 在总结果中按花名册逐人查找；
   - 为名单中的每一个人生成一份独立的 `.xlsx`；
   - 最后打包成一个 `.zip`，供你一键下载。

---

### 6. 常见问题（Q&A）

**Q1：双击 exe 黑窗一闪而过就没了？**  
A：十有八九是把 `UniData-Sieve.exe` 从原有文件夹里单独拽出来运行了，导致找不到 `_internal` 等内部依赖。请回到完整目录内运行，或重新解压整套程序。

**Q2：黑窗口在，浏览器开了，但页面显示“无法访问此页面 / 拒绝连接”？**  
A：可能是安全软件/公司防火墙拦截了本地服务，或端口被占用。可以先关闭黑窗，等 1 分钟后重试。必要时尝试临时关闭安全软件或换一台机器测试。

**Q3：运行很久没反应，是死机了吗？**  
A：大概率不是。你可以切回那个黑色窗口观察日志输出——如果在“刷字”，说明正在努力解析（尤其是超大的、嵌套复杂的 ZIP）。请耐心等待。

**Q4：结果表中“姓名 / 学号 / 班级”对不齐或错列？**  
A：原始 Excel 的表头可能非常不规范（比如`班级`,`班 级`,`班级 `等各种变体）。程序内部会做一定的清洗（去空格、去换行等），但如果差异太大，建议在源表中先稍微统一下字段名再用本工具处理。

---

## 二、技术人员：源码运行与开发指南

如果你需要查看源码、参与维护或进行二次开发，请参考本节。

### 1. 仓库结构概览

```text
UniData-Sieve-Excel/
├── app.py              # Streamlit Web 界面入口
├── cli.py              # 命令行入口
├── launcher.py         # 打包后 exe 的统一启动封装（可选）
├── build.bat           # Windows 下打包脚本（基于 PyInstaller 等）
├── run_app.bat         # 本地快速启动 GUI 的脚本
├── requirements.txt    # Python 依赖
├── pyproject.toml      # 项目/打包配置（可选）
├── config/
│   ├── schema.py       # AppConfig、匹配规则等配置模型
│   └── default_config.json
├── engine/
│   ├── __init__.py     # run_pipeline：串联整个数据处理流程
│   ├── processor.py    # 解压 ZIP、递归遍历文件、编码回退处理
│   ├── excel_handler.py# 逐行读取 Excel、行文本归一化与清洗
│   ├── matcher.py      # MatchMode & MatchRule，AND/OR/REGEX 匹配
│   ├── aggregator.py   # 聚合与流式写入（超阈值切换防 OOM）
│   └── exporter.py     # 按列模式导出结果到 xlsx
└── tests/              # pytest 单元测试
```

整体数据流由 `engine.run_pipeline` 串联：

1. `processor.extract_all_excels(source, temp_dir)`  
   - 支持目录或 zip 输入；
   - 递归解压（限制 max_depth）；
   - 用 `_decode_zip_filename` 按 UTF-8 → GBK → CP437 回退解析文件名。

2. `excel_handler.scan_excel(...)`  
   - 针对 `.xlsx` 使用 `openpyxl` 的 read_only 模式逐行迭代；
   - 将每行转为统一字符串，调用 `matcher.match_row` 判断命中。

3. `aggregator.aggregate(...)`  
   - 对命中记录进行缓冲；
   - 当行数超过 `STREAM_THRESHOLD`（默认 50k）时切换到 openpyxl 流式写入。

4. `exporter.to_xlsx(...)`  
   - 根据 column_mode（如 FULL / COMMON）和优先关键字控制列的保留与排序；
   - 生成最终结果 Excel。

---

### 2. 环境准备与源码运行

#### 2.1 克隆仓库

```bash
git clone https://github.com/hua-tian99/UniData-Sieve-Excel.git
cd UniData-Sieve-Excel
```

#### 2.2 创建虚拟环境并安装依赖（推荐）

要求：**Python 3.8+**

```bash
# 创建虚拟环境
python -m venv .venv

# Windows 激活
.\.venv\Scripts\activate

# Linux / macOS 激活
# source .venv/bin/activate

# 安装依赖
pip install -r requirements.txt
```

#### 2.3 启动 Streamlit Web 界面

```bash
# 在已激活虚拟环境下执行
streamlit run app.py
```

浏览器会自动打开本地地址（通常为 `http://localhost:8501`），界面和 exe 启动时看到的是同一套 UI。

也可以使用仓库自带的批处理脚本：

```bash
run_app.bat
```

你可以在 `run_app.bat` 内加入虚拟环境激活等逻辑，方便一键启动。

---

### 3. 命令行（CLI）模式

仓库包含 `cli.py`，用于在没有 GUI 的环境（服务器/定时任务）下直接调用 pipeline。  
使用方式视你在 `cli.py` 中实现的参数解析而定，一般类似：

```bash
python cli.py --help

# 伪例：按学号列表筛选
python cli.py ^
  --source "D:\data\原始数据.zip" ^
  --mode or ^
  --keywords 24405101001 24405101002 24405101003 ^
  --out result.xlsx
```

请根据 `cli.py` 实际实现补充/调整参数说明。

---

### 4. 打包为 exe（维护者参考）

项目中已包含 `build.bat`，通常里面会调用 PyInstaller 等工具。例如：

```bash
pip install pyinstaller

pyinstaller ^
  --noconfirm ^
  --onefile ^
  --name UniData-Sieve ^
  launcher.py
```

打包完成后：

- exe 一般出现在 `dist/` 目录；
- 将 exe 与其依赖（如 `_internal` 等）打包成一个发布文件夹；
- 上传到 GitHub 的 **Releases**，供普通用户下载。

---



## 许可证
本项目基于 MIT License 开源，详情见仓库中的 `LICENSE` 文件。
