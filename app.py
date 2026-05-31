from typing import TYPE_CHECKING

if TYPE_CHECKING:
    # Help static analyzers / linters resolve the import during type checking
    import streamlit as st  # pragma: no cover

try:
    import streamlit as st  # type: ignore
except Exception:  # pragma: no cover - friendly error when Streamlit is not installed in linting env
    # Provide a clear runtime error for missing Streamlit while keeping static analyzers happy.
    raise ImportError(
        "Streamlit is required to run this app. Install with: pip install streamlit"
    )
import pandas as pd
from staining_logic import (
    load_excel_staining_plan,
    compute_staining,
    adjust_fmo_generic,
    export_to_single_sheet,
    validate_and_prepare_df,
    build_printable_html_report,
    html_to_pdf_bytes,
)


# 页面设置
st.set_page_config(page_title="抗体配方计算器", layout="wide")
st.title("🧬 流式抗体配方计算器")

# 切换输入模式
use_excel = st.checkbox("📁 使用 Excel 文件上传代替网页填写(需满足格式要求)", value=False)

# 初始化空白表格（只在首次加载时执行）
default_df = pd.DataFrame(
    {
        "marker": ["" for _ in range(5)],
        "荧光染料": ["" for _ in range(5)],
        "稀释比例": ["1:100" for _ in range(5)],
        "是否作为FMO": ["" for _ in range(5)],
        "一抗/二抗/胞内抗体": ["一抗" for _ in range(5)],
    }
)
if "manual_df" not in st.session_state:
    st.session_state["manual_df"] = default_df.copy()

# 样本数输入
sample_n = st.number_input("🔢 样本数量", min_value=1, value=50, step=1)

with st.expander("🧾 报告信息（用于打印页眉）", expanded=True):
    m1, m2, m3 = st.columns(3)
    experiment_date = m1.text_input("实验日期", value="")
    operator = m2.text_input("操作者", value="")
    batch_id = m3.text_input("实验批次", value="")


# 表格填写说明
with st.expander("📘 表格填写说明（点击展开）", expanded=False):
    st.markdown(
        """
| 字段 | 示例 | 必填 | 说明 |
|------|------|------|------|
| marker | CD3 | ✅ | 抗体名称 |
| 荧光染料 | FITC | ✅ | 荧光标签名 |
| 稀释比例 | 1:100 | 条件必填 | 一抗/二抗/胞内抗体需要；自发荧光可留空 |
| 是否作为FMO | 是 / 留空 | 条件必填 | 一抗/胞内抗体可填写；二抗会自动按“否”处理 |
| 一抗/二抗/胞内抗体 | 一抗 | ✅ | 可以填写一抗/二抗/胞内抗体/自发荧光 |
"""
    )
    st.info("⚠️ 请确保字段名称不变，填写内容规范，请勿将live/dead和FC block计算进来，否则计算结果可能有误。")

# 数据输入
df = None
if use_excel:
    uploaded_file = st.file_uploader("上传 staining_plan.xlsx 文件", type=["xlsx"])
    if uploaded_file:
        try:
            df = load_excel_staining_plan(uploaded_file)
        except Exception as e:
            st.error(f"❌ 文件读取失败: {e}")
            st.stop()
else:
    st.markdown("📋 请在下方表格中填写配方信息，可复制粘贴 Excel 表格区域")

    edited_df = st.data_editor(
        st.session_state["manual_df"],
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="editor",
    )

    if st.button("✅ 使用上方内容开始计算"):
        st.session_state["manual_df"] = edited_df.copy()
        df = edited_df.copy()


def render_report_section(label, result_tuple):
    step1_df, step2_fsb, step2_dyes, step3_df, step4_df = result_tuple

    with st.container(border=True):
        st.subheader(f"🧪 {label} 配制方案")

        c1, c2, c3 = st.columns(3)
        c1.metric("Step1 条目数", len(step1_df))
        c2.metric("Step3 FMO配置数", len(step3_df))
        c3.metric("Step2 FSB (μL)", step2_fsb)

        st.markdown("**Step 1：染料稀释**")
        st.dataframe(step1_df, use_container_width=True, hide_index=True)

        st.markdown("**Step 2：主 Mix 配置**")
        step2_show = pd.DataFrame(
            {
                "项目": ["FSB体积 (μL)", "加入染料"],
                "值": [step2_fsb, "、".join(step2_dyes) if step2_dyes else "无"],
            }
        )
        st.table(step2_show)

        st.markdown("**Step 3：FMO 配置**")
        st.dataframe(step3_df, use_container_width=True, hide_index=True)

        st.markdown("**Step 4：剩余验证**")
        st.dataframe(step4_df, use_container_width=True, hide_index=True)


# 计算逻辑
if df is not None and not df.empty:
    prepared_df, errors = validate_and_prepare_df(df)

    if errors:
        for err in errors:
            st.error(f"❌ {err}")
    else:
        results = {}
        with st.spinner("🧪 正在计算，请稍候..."):
            for ab_type in ["一抗", "二抗", "胞内抗体"]:
                df_sub = prepared_df[prepared_df["抗体类型"] == ab_type].copy()
                if not df_sub.empty:
                    adjusted = adjust_fmo_generic(prepared_df, df_sub)
                    results[ab_type] = compute_staining(adjusted, df_sub, sample_n)

            output_path = "staining_result.xlsx"
            export_to_single_sheet(results, output_path)
            html_report = build_printable_html_report(
                results,
                sample_n,
                report_meta={
                    "实验日期": experiment_date,
                    "操作者": operator,
                    "实验批次": batch_id,
                },
            )
            pdf_report = html_to_pdf_bytes(html_report)

        st.success("✅ 计算完成！下方可直接查看报告，或下载 Excel / 打印版 HTML。")

        st.markdown("## 📑 网页报告")
        for label in ["一抗", "二抗", "胞内抗体"]:
            if label in results:
                render_report_section(label, results[label])

        d1, d2, d3 = st.columns(3)
        with d1:
            with open(output_path, "rb") as f:
                st.download_button("📥 下载 Excel 结果", f, file_name="staining_result.xlsx")
        with d2:
            st.download_button(
                "🖨️ 下载打印版 HTML 报告",
                data=html_report,
                file_name="staining_report.html",
                mime="text/html",
            )
        with d3:
            if pdf_report is not None:
                st.download_button(
                    "📄 下载 PDF 报告",
                    data=pdf_report,
                    file_name="staining_report.pdf",
                    mime="application/pdf",
                )
            else:
                st.caption("未检测到 PDF 引擎（可选安装：weasyprint）")




