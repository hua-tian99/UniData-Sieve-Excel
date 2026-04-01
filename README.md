# UniData-Sieve-Excel

UniData-Sieve-Excel 是一个基于 Python 构建的本地化数据流转与清洗工具，提供基于 Streamlit 的 Web 交互界面。本项目主要用于解决多来源、非标准化 Excel 表格的批量读取、条件过滤及自动化聚合问题。

## 🎯 设计目标

在处理如“高校综合素质测评”、“批量表单收集”等场景时，常面临数据源高度分散（层层嵌套的 ZIP）、格式字段不统一、以及无效脏数据极多的情况。本项目通过抽象“解压 - 扫描 - 匹配 - 聚合 - 导出”的处理管线（Pipeline），实现对此类脏数据的自动化提纯与梳理。

## ⚙️ 核心特性 (Core Features)

- **嵌套压缩包解析防御机制**
  - 支持最高 5 层深度的 `.zip` 递归解压。
  - 内置 `UTF-8 -> GBK -> CP437` 编码回退探测（Fallback），有效解决非标压缩软件导致的中文文件名乱码问题。
- **低内存占用的数据 IO (OOM 防御)**
  - 读取端：针对 `.xlsx` 文件强制开启 `openpyxl(read_only=True)` 的游标迭代读取，不将全量数据加载入内存，并自动清洗表头首尾的不可见字符。
  - 写入端：内置动态流式降级机制。当命中数据超过阈值（默认 50,000 行）时，自动放弃 Pandas DataFrame，无缝切换为流式追加（Append）模式直接落盘。
- **灵活的路由匹配引擎**
  - 支持行级别的关键词过滤，提供 `AND`（全包含）、`OR`（任一包含）以及 `REGEX`（正则表达式）三种逻辑路由。
- **动态列头对齐**
  - `Full` 模式（Outer Join）：保留所有出现的字段，缺失值置空。
  - `Common` 模式（Inner Join）：交集提取，仅保留所有来源表共有的基础字段。
  - 支持基于关键字配置字段排序优先级。

## 🛠️ 技术栈

- **核心语言**：Python 3.8+
- **前端 / 交互**：Streamlit
- **数据处理**：Pandas, Openpyxl, Xlrd (提供对遗留 `.xls` 格式的老旧兼容)
- **日志与监控**：Loguru

## 📂 仓库结构

```text
├── app.py                  # 基于 Streamlit 的 GUI 主入口与 Session 管理
├── cli.py                  # CLI 命令行入口 (适用于自动化批处理脚本)
├── launcher.py / run_*.bat # 快捷启动与环境检测脚本
├── config/
│   ├── schema.py           # 运行时配置的数据模型化校验
│   └── default_config.json # 缺省业务参数
└── engine/                 # 核心数据管线模块
    ├── processor.py        # 文件 IO 层：解压提取、编码回退与临时沙盒管理
    ├── excel_handler.py    # 表格 IO 层：读游标管理与单元格级脏字符正则清洗
    ├── matcher.py          # 路由层：字符串归一化与正则引擎匹配
    ├── aggregator.py       # 缓冲池：容量监控与内存/流式落盘切换
    └── exporter.py         # 渲染层：动态表头排序与序列化导出 (xlsx)
