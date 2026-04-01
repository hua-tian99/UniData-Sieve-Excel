import streamlit as st
import tempfile
import zipfile
import io
import os
from pathlib import Path
import pandas as pd
from loguru import logger

from engine.processor import extract_all_excels
from engine.excel_handler import scan_excel
from engine.aggregator import aggregate
from engine.exporter import to_xlsx
from config.schema import AppConfig, ColumnMode
from engine.matcher import MatchRule as EngineMatchRule, MatchMode

st.set_page_config(page_title="UniData-Sieve", page_icon="📦", layout="wide")

# Task 7a: Session State Management for Temp Directory
def get_temp_dir() -> Path:
    if "temp_dir_obj" not in st.session_state:
        # Ignore cleanup errors for Windows file handles
        st.session_state.temp_dir_obj = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
    return Path(st.session_state.temp_dir_obj.name)

# Helper function to run pipeline with progress
def _run_with_progress(source_paths: list[Path], config: AppConfig):
    temp_dir = get_temp_dir()
    output_path = temp_dir / config.output_filename

    # Construct rule
    mode_map = {"and": MatchMode.AND, "or": MatchMode.OR, "regex": MatchMode.REGEX}
    match_mode = mode_map.get(config.match_rule.mode.lower(), MatchMode.AND)
    rule = EngineMatchRule(
        keywords=config.match_rule.keywords,
        mode=match_mode,
        pattern=config.match_rule.pattern,
    )

    excel_paths = []
    for sp in source_paths:
        if sp.suffix.lower() == ".zip":
            st.info(f"开始解压解析: {sp.name}")
            excel_paths.extend(extract_all_excels(sp, temp_dir, max_depth=config.max_depth))
        else:
            excel_paths.append(sp)
        
    total_files = len(excel_paths)
    
    if total_files == 0:
        st.warning("未找到任何 Excel 文件 (可能加密、损坏或为空)。")
        return None, 0, output_path

    progress_bar = st.progress(0)
    status_text = st.empty()

    all_records = []
    for i, excel_path in enumerate(excel_paths):
        status_text.text(f"正在扫描 ({i+1}/{total_files}): {excel_path.name}")
        all_records.extend(scan_excel(excel_path, rule))
        progress_bar.progress((i + 1) / total_files)

    status_text.text("扫描完成，正在聚合数据...")
    df, total_rows = aggregate(iter(all_records), output_path, stream_threshold=config.stream_threshold)

    status_text.text("正在导出最终文件...")
    if df is not None:
        to_xlsx(df, output_path, column_mode=config.column_mode, priority_keywords=config.match_rule.keywords)
    
    status_text.text("处理完成！")
    return df, total_rows, output_path

def main():
    st.title("📦 UniData-Sieve 数据提纯引擎")
    
    # Task 7b: Parameters in Sidebar
    st.sidebar.header("⚙️ 参数配置")
    
    match_mode_str = st.sidebar.radio("匹配模式", options=["AND", "OR", "REGEX"], index=0).lower()
    
    if match_mode_str == "regex":
        pattern = st.sidebar.text_input("正则表达式", value="")
        keywords_str = ""
    else:
        pattern = None
        keywords_str = st.sidebar.text_input("匹配关键词 (空格分隔)", value="")
        
    column_mode_str = st.sidebar.radio("列精简模式", options=["FULL", "COMMON"], index=0).lower()
    max_depth = st.sidebar.slider("最大嵌套解压深度", 1, 10, 5)
    
    # Task 7b: File Uploader
    st.header("1. 上传数据源压缩包或表格")
    uploaded_files = st.file_uploader(
        "支持批量混传：任意数量的 .zip / .xlsx / .xls 文件组合", 
        type=["zip", "xlsx", "xls"], 
        accept_multiple_files=True
    )
    
    if "process_done" not in st.session_state:
        st.session_state.process_done = False
        st.session_state.result_df = None
        st.session_state.result_path = None
        st.session_state.total_rows = 0

    if uploaded_files and not st.session_state.process_done:
        if st.button("🚀 开始处理"):
            keywords = [k.strip() for k in keywords_str.split(" ") if k.strip()] if keywords_str else []
            
            # Use AppConfig model directly
            config = AppConfig(
                match_rule=dict(keywords=keywords, mode=match_mode_str, pattern=pattern),
                column_mode=ColumnMode(column_mode_str),
                max_depth=max_depth
            )
            
            temp_dir = get_temp_dir()
            source_paths = []
            
            # Save all uploaded files to temp directory
            for uf in uploaded_files:
                out_file_path = temp_dir / uf.name
                with open(out_file_path, "wb") as f:
                    f.write(uf.getvalue())
                source_paths.append(out_file_path)
                
            with st.spinner("系统处理中，请稍候..."):
                df, total_rows, out_path = _run_with_progress(source_paths, config)
                st.session_state.process_done = True
                st.session_state.result_df = df
                st.session_state.total_rows = total_rows
                st.session_state.result_path = out_path
                st.rerun()

    if st.session_state.process_done:
        if st.button("🔄 重新开始新任务"):
            st.session_state.process_done = False
            st.rerun()
            
        df = st.session_state.result_df
        total_rows = st.session_state.total_rows
        out_path = st.session_state.result_path
        
        st.success(f"处理成功！共计匹配 {total_rows} 条记录。")
        
        # Task 7c: Result Preview
        st.header("2. 结果预览与下载")
        if df is not None:
            if total_rows > 0:
                # 截取前100条用于预览
                preview_df = df.head(100).copy()
                
                # 修复终端告警：将复杂的混合类型（如同时包含数字和字符串的“学号”列）强制转为字符串，防止 PyArrow 崩溃
                for col in preview_df.columns:
                    if preview_df[col].dtype == 'object':
                        preview_df[col] = preview_df[col].astype(str)
                
                st.dataframe(preview_df, use_container_width=True)
                if total_rows > 100:
                    st.info("⬆️ 以上为前 100 行数据预览。")
            else:
                st.warning("⚠️ 未匹配到任何数据，结果集为空。")
        else:
            st.warning("⚠️ 匹配数据量超过 50,000 条，已触发底层引擎的流式写入模式保护内存，暂不支持在线预览。请直接下载文件查阅。")
            
        if out_path is not None and out_path.exists():
            with open(out_path, "rb") as f:
                st.download_button(
                    label="📥 下载提纯结果 (result.xlsx)",
                    data=f,
                    file_name="result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            
        # Task 8: Smart Batch Export (Split by Roster)
        st.divider()
        st.header("3. 智能批量导出 (按名单拆分)")
        st.write("上传一个名单 Excel，系统将根据您选择的列对上述提纯结果进行拆分，为名单中的每个人/项独立生成一个 Excel 表格，并一键打包为 ZIP 供您下载。")
        
        roster_file = st.file_uploader("📋 步骤1：上传您的比对名单 (.xlsx/.xls)", type=["xlsx", "xls"], key="roster")
        
        if roster_file is not None and total_rows > 0:
            roster_df = pd.read_excel(roster_file)
            
            col1, col2 = st.columns(2)
            with col1:
                # Target column in result
                result_cols = df.columns.tolist() if df is not None else pd.read_excel(out_path, nrows=0).columns.tolist()
                result_target_col = st.selectbox("🎯 步骤2：选择【刚生成的提纯结果】中用于匹配的列", options=result_cols)
            with col2:
                # Target column in roster
                roster_cols = roster_df.columns.tolist()
                roster_target_col = st.selectbox("🎯 步骤3：选择【您刚才上传的名单】中等效匹配的列", options=roster_cols)
                
            if st.button("✂️ 开始拆分并打包 ZIP"):
                with st.spinner("📦 正在按名单筛选拆分，请稍后..."):
                    # Reload full DF if we are in stream mode, otherwise use memory DF
                    full_df = df if df is not None else pd.read_excel(out_path)
                    
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        # Iterate through roster unique values
                        unique_values = roster_df[roster_target_col].dropna().unique()
                        
                        file_count = 0
                        for val in unique_values:
                            # Filter
                            subset = full_df[full_df[result_target_col] == val]
                            if not subset.empty:
                                xlsx_buffer = io.BytesIO()
                                subset.to_excel(xlsx_buffer, index=False, engine="openpyxl")
                                # Save to zip with value as filename
                                safe_val = str(val).replace("/", "_").replace("\\", "_").replace(":", "_")
                                zip_file.writestr(f"{safe_val}.xlsx", xlsx_buffer.getvalue())
                                file_count += 1
                                
                    st.success(f"✅ 批量拆分打包完成！共生成 {file_count} 个独立表格。")
                    st.download_button(
                        label="📦 立即下载拆分数据压缩包 (split_results.zip)",
                        data=zip_buffer.getvalue(),
                        file_name="split_results.zip",
                        mime="application/zip",
                        type="primary"
                    )

if __name__ == "__main__":
    main()
