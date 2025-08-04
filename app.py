import streamlit as st
import pandas as pd
from staining_logic import (
    load_excel_staining_plan,
    compute_staining,
    adjust_fmo_generic,
    export_to_single_sheet
)

# 页面配置
st.set_page_config(page_title="抗体配方计算器", layout="centered")
st.title("🧬 流式抗体配方计算器")

# 初始化 session_state
if "manual_df" not in st.session_state:
    st.session_state["manual_df"] = pd.DataFrame({
        "marker": ["" for _ in range(5)],
        "荧光染料": ["" for _ in range(5)],
        "稀释比例": ["1:100" for _ in range(5)],
        "是否作为FMO": ["" for _ in range(5)],
        "一抗/二抗/胞内抗体": ["一抗" for _ in range(5)]
    })

# 输入模式选择
use_excel = st.checkbox("📁 使用 Excel 文件上传代替网页填写(需满足格式要求)", value=False)

# 样本数输入
sample_n = st.number_input("🔢 样本数量", min_value=1, value=50, step=1)

# 表格填写说明
with st.expander("📘 表格填写说明（点击展开）", expanded=False):
    st.markdown("""
    | 字段 | 示例 | 必填 | 说明 |
    |------|------|------|------|
    | marker | CD3 | ✅ | 抗体名称 |
    | 荧光染料 | FITC | ✅ | 荧光标签名 |
    | 稀释比例 | 1:100 | ✅ | 格式为 `1:100`，不能写成 `%` |
    | 是否作为FMO | 是 / 留空 | ❓ | 写“是”表示参与 FMO |
    | 一抗/二抗/胞内抗体 | 一抗 | ✅ | 可以填写一抗/二抗/胞内抗体/自发荧光 |
    """)
    st.info("⚠️ 请确保字段名称不变，填写内容规范，请勿将live/dead和FC block计算进来，否则计算结果可能有误。")

# 数据输入
df = None
if use_excel:
    uploaded_file = st.file_uploader("上传 staining_plan.xlsx 文件", type=["xlsx"])
    if uploaded_file:
        df = load_excel_staining_plan(uploaded_file)
else:
    st.markdown("📋 请在下方表格中填写配方信息，可复制粘贴 Excel 表格区域")
    edited_df = st.data_editor(
        st.session_state["manual_df"],
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="manual_editor"
    )
    st.session_state["manual_df"] = edited_df

    if st.button("✅ 使用上方内容开始计算"):
        df = edited_df.copy()
        st.session_state["manual_df"] = df  # 再次保存，防止刷新丢失

        # 字段校验
        required_columns = ["marker", "荧光染料", "稀释比例", "是否作为FMO", "一抗/二抗/胞内抗体"]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            st.error(f"❌ 缺少必要字段：{', '.join(missing)}")
            st.stop()

        # 稀释比例格式校验
        import re
        invalid_dilutions = df[~df["稀释比例"].astype(str).str.match(r"^1:\d+$")]
        if not invalid_dilutions.empty:
            st.warning("⚠️ 以下行稀释比例格式不正确，请使用 `1:100` 格式：")
            st.dataframe(invalid_dilutions)
            st.stop()

        # 数据标准化
        df["是否作为FMO"] = df["是否作为FMO"].fillna("").apply(lambda x: str(x).strip() == "是")
        df["抗体类型"] = df["一抗/二抗/胞内抗体"].fillna("一抗").apply(str.strip)

        # 计算逻辑
        results = {}
        for ab_type in ["一抗", "二抗", "胞内抗体"]:
            df_sub = df[df["抗体类型"] == ab_type].copy()
            if not df_sub.empty:
                adjusted = adjust_fmo_generic(df, df_sub)
                results[ab_type] = compute_staining(adjusted, df_sub, sample_n)

        # 结果导出
        output_path = "staining_result.xlsx"
        export_to_single_sheet(results, output_path)
        with open(output_path, "rb") as f:
            st.download_button("📥 下载计算结果", f, file_name="staining_result.xlsx")

        # 结果预览
        for ab_type, result_df in results.items():
            st.subheader(f"📊 {ab_type} 计算结果预览")
            st.dataframe(result_df.head())
