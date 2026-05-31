import re
from datetime import datetime
from html import escape
from typing import Dict, List, Optional, Tuple



import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

REQUIRED_COLUMNS = ["marker", "荧光染料", "稀释比例", "是否作为FMO", "一抗/二抗/胞内抗体"]
VALID_ANTIBODY_TYPES = {"一抗", "二抗", "胞内抗体", "自发荧光"}
ANTIBODY_TYPE_ALIASES = {
    "胞内": "胞内抗体",
    "intracellular": "胞内抗体",
    "secondary": "二抗",
    "primary": "一抗",
}


def parse_dilution_ratio(value: str) -> int:
    text = str(value).strip()
    matched = re.match(r"^1:(\d+)$", text)
    if not matched:
        raise ValueError(f"无效稀释比例: {value}（应为 1:100 这类格式）")
    denominator = int(matched.group(1))
    if denominator <= 0:
        raise ValueError(f"无效稀释比例: {value}（分母必须大于 0）")
    return denominator


def normalize_antibody_type(value: str) -> str:
    raw = str(value).strip() if pd.notna(value) else ""
    if not raw:
        return "一抗"
    if raw in VALID_ANTIBODY_TYPES:
        return raw
    lowered = raw.lower()
    return ANTIBODY_TYPE_ALIASES.get(lowered, ANTIBODY_TYPE_ALIASES.get(raw, raw))


def validate_and_prepare_df(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    errors: List[str] = []

    missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_cols:
        errors.append(f"缺少必要字段: {', '.join(missing_cols)}")
        return df, errors

    prepared = df.copy()
    prepared["marker"] = prepared["marker"].fillna("").astype(str).str.strip()
    prepared["荧光染料"] = prepared["荧光染料"].fillna("").astype(str).str.strip()
    prepared["是否作为FMO"] = prepared["是否作为FMO"].fillna("").apply(lambda x: str(x).strip() == "是")
    prepared["抗体类型"] = prepared["一抗/二抗/胞内抗体"].apply(normalize_antibody_type)

    # 业务规则：
    # 1) 二抗不参与FMO，因此无论是否填写“是”，统一按 False 处理
    # 2) 自发荧光不需要稀释比例
    prepared.loc[prepared["抗体类型"] == "二抗", "是否作为FMO"] = False

    # 删除完全空行
    prepared = prepared[
        ~(
            (prepared["marker"] == "")
            & (prepared["荧光染料"] == "")
            & (prepared["一抗/二抗/胞内抗体"].fillna("").astype(str).str.strip() == "")
        )
    ].copy()

    invalid_dilution_rows: List[int] = []
    for idx, row in prepared.iterrows():
        # 自发荧光无需稀释比例
        if row["抗体类型"] == "自发荧光":
            continue

        try:
            parse_dilution_ratio(row["稀释比例"])
        except ValueError:
            invalid_dilution_rows.append(idx + 1)

    if invalid_dilution_rows:
        errors.append(f"稀释比例格式错误（行号从1开始）: {invalid_dilution_rows}")

    invalid_type_rows = prepared[~prepared["抗体类型"].isin(VALID_ANTIBODY_TYPES)]
    if not invalid_type_rows.empty:
        rows = (invalid_type_rows.index + 1).tolist()
        errors.append(f"抗体类型不支持（行号从1开始）: {rows}，允许值: {', '.join(sorted(VALID_ANTIBODY_TYPES))}")

    return prepared, errors



def load_excel_staining_plan(file_path):
    return pd.read_excel(file_path, engine="openpyxl")


def adjust_fmo_generic(df_all, df_current):
    fmo_marker_total = df_all[df_all["是否作为FMO"]]["marker"].tolist()
    df_extra = pd.DataFrame({"marker": list(set(fmo_marker_total) - set(df_current["marker"]))})
    df_extra["荧光染料"] = ""
    df_extra["稀释比例"] = ""
    df_extra["是否作为FMO"] = True
    df_extra["抗体类型"] = "非当前抗体类型FMO"
    return pd.concat([df_current, df_extra], ignore_index=True)


def compute_staining(df_fmo_ref, df_dye_only, sample_n, volume_per_well=50):
    fmo_markers = df_fmo_ref[df_fmo_ref["是否作为FMO"]]
    total_wells = sample_n + len(fmo_markers)
    total_volume = total_wells * volume_per_well

    step1_results = []
    dye_mix_info: Dict[str, Dict] = {}
    total_fmo_final_vol = 0
    step3_results = []

    for _, row in df_dye_only.iterrows():
        if row["抗体类型"] == "自发荧光":
            continue

        marker = row["marker"]
        dilution = parse_dilution_ratio(row["稀释比例"])
        is_fmo = row["是否作为FMO"]
        fmo_total = len(fmo_markers)

        if is_fmo:
            final_volume = sample_n + fmo_total - 1
            dye_vol = round(final_volume * volume_per_well / dilution, 2)
            fsb_vol = round(final_volume - dye_vol, 2)
            total_fmo_final_vol += final_volume
        else:
            final_volume = round(total_volume / dilution, 2)
            dye_vol = final_volume
            fsb_vol = 0

        step1_results.append(
            {
                "marker": marker,
                "荧光染料": row["荧光染料"],
                "稀释比例": row["稀释比例"],
                "是否为FMO": "是" if is_fmo else "否",
                "所需抗体体积 (μL)": dye_vol,
                "加入FSB体积 (μL)": fsb_vol,
                "稀释后总体积 (μL)": final_volume,
            }
        )

        dye_mix_info[marker] = {
            "dilution": dilution,
            "final_vol": final_volume,
            "used_in_mix": 0,
            "is_fmo": is_fmo,
        }

    step2_fsb = round(
        total_volume - total_fmo_final_vol - sum(info["final_vol"] for info in dye_mix_info.values() if not info["is_fmo"]),
        2,
    )
    step2_dyes = [m for m, info in dye_mix_info.items() if not info["is_fmo"]]

    for marker in fmo_markers["marker"]:
        ab_type_series = df_fmo_ref[df_fmo_ref["marker"] == marker]["抗体类型"]
        ab_type = ab_type_series.values[0] if not ab_type_series.empty else ""

        if ab_type == "自发荧光":
            all_dyes = [m for m in dye_mix_info if not dye_mix_info[m]["is_fmo"]]
            fsb_vol = round(volume_per_well - len(all_dyes), 2)
            fmo_entry = {"FMO通道": marker, "加入主mix体积 (μL)": fsb_vol}
            for m in all_dyes:
                fmo_entry[f"加 {m} 染料 (μL)"] = 1
                dye_mix_info[m]["used_in_mix"] += 1
            step3_results.append(fmo_entry)
            continue

        if marker not in dye_mix_info:
            continue

        other_fmo = [m for m in fmo_markers["marker"] if m != marker and m in dye_mix_info]
        fsb_vol = round(volume_per_well - len(other_fmo), 2)
        fmo_entry = {"FMO通道": marker, "加入主mix体积 (μL)": fsb_vol}
        for m in other_fmo:
            fmo_entry[f"加 {m} 染料 (μL)"] = 1
            dye_mix_info[m]["used_in_mix"] += 1
        step3_results.append(fmo_entry)

    master_mix_used = sum(row["加入主mix体积 (μL)"] for row in step3_results)
    master_mix_remaining = round(total_volume - total_fmo_final_vol - master_mix_used, 2)

    dye_remaining = []
    for m, info in dye_mix_info.items():
        if info["is_fmo"]:
            remain = round(info["final_vol"] - info["used_in_mix"], 2)
            if remain > 0:
                dye_remaining.append({"marker": m, "加入体积 (μL)": remain})
    dye_remaining.append({"marker": "剩余主mix", "加入体积 (μL)": master_mix_remaining})

    return (
        pd.DataFrame(step1_results),
        step2_fsb,
        step2_dyes,
        pd.DataFrame(step3_results),
        pd.DataFrame(dye_remaining),
    )


def export_to_single_sheet(results_dict, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "FMO 配制方案"

    for label, (step1_df, step2_fsb, step2_dyes, step3_df, step4_df) in results_dict.items():
        ws.append([f"【{label}】Step 1: 染料稀释"])
        for row in dataframe_to_rows(step1_df, index=False, header=True):
            ws.append(row)
        ws.append([])

        ws.append(["Step 2: 主mix配置"])
        ws.append(["FSB体积 (μL)", step2_fsb])
        ws.append(["加入染料", ", ".join(step2_dyes)])
        ws.append([])

        ws.append(["Step 3: FMO配置"])
        for row in dataframe_to_rows(step3_df, index=False, header=True):
            ws.append(row)
        ws.append([])

        ws.append(["Step 4: 剩余验证"])
        for row in dataframe_to_rows(step4_df, index=False, header=True):
            ws.append(row)
        ws.append([])

    wb.save(output_path)


def _df_to_html_table(df: pd.DataFrame) -> str:
    if df is None or df.empty:
        return '<p class="empty">无数据</p>'
    return df.to_html(index=False, classes="report-table", border=0)


def build_printable_html_report(results_dict, sample_n: int, report_meta: Optional[Dict[str, str]] = None) -> str:
    report_meta = report_meta or {}
    experiment_date = escape(str(report_meta.get("实验日期", "")))
    operator = escape(str(report_meta.get("操作者", "")))
    batch_id = escape(str(report_meta.get("实验批次", "")))
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    sections = []

    for label, (step1_df, step2_fsb, step2_dyes, step3_df, step4_df) in results_dict.items():
        step2_dyes_text = "、".join(step2_dyes) if step2_dyes else "无"

        section_html = f"""
        <section class="card">
          <h2>{escape(label)} 配制方案</h2>
          <div class="meta">样本数量：<strong>{sample_n}</strong></div>

          <h3>Step 1：染料稀释</h3>
          {_df_to_html_table(step1_df)}

          <h3>Step 2：主 Mix 配置</h3>
          <table class="report-table compact">
            <tr><th>FSB体积 (μL)</th><td>{step2_fsb}</td></tr>
            <tr><th>加入染料</th><td>{escape(step2_dyes_text)}</td></tr>
          </table>

          <h3>Step 3：FMO 配置</h3>
          {_df_to_html_table(step3_df)}

          <h3>Step 4：剩余验证</h3>
          {_df_to_html_table(step4_df)}
        </section>
        """
        sections.append(section_html)

    html = f"""
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>流式抗体配方报告</title>
  <style>
    :root {{
      --ink: #1f2937;
      --sub: #4b5563;
      --line: #d1d5db;
      --soft: #f8fafc;
      --brand: #0f766e;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      color: var(--ink);
      font-family: "Segoe UI", "PingFang SC", "Microsoft YaHei", Arial, sans-serif;
      background: white;
      line-height: 1.45;
    }}
    .wrap {{ max-width: 980px; margin: 24px auto; padding: 0 16px; }}
    h1 {{ margin: 0 0 8px 0; color: var(--brand); font-size: 28px; }}
    .subtitle {{ color: var(--sub); margin-bottom: 18px; }}
    .meta-card {{
      display: grid;
      grid-template-columns: repeat(2, minmax(220px, 1fr));
      gap: 8px 16px;
      border: 1px solid var(--line);
      border-radius: 10px;
      background: white;
      padding: 12px 14px;
      margin-bottom: 14px;
      font-size: 14px;
    }}
    .meta-card span {{ color: var(--sub); }}

    .card {{
      border: 1px solid var(--line);
      border-radius: 10px;
      padding: 16px;
      margin-bottom: 18px;
      page-break-inside: avoid;
      background: var(--soft);
    }}
    h2 {{ margin: 0 0 8px 0; font-size: 22px; }}
    h3 {{ margin: 14px 0 8px 0; font-size: 16px; }}
    .meta {{ color: var(--sub); margin-bottom: 8px; }}
    .report-table {{
      width: 100%; border-collapse: collapse; background: white;
      border: 1px solid var(--line); font-size: 13px;
    }}
    .report-table th, .report-table td {{
      border: 1px solid var(--line); padding: 6px 8px; text-align: left;
      vertical-align: top;
    }}
    .report-table thead th {{ background: #eef2f7; }}
    .compact th {{ width: 220px; background: #eef2f7; }}
    .empty {{ color: var(--sub); font-style: italic; }}

    @media print {{
      .wrap {{ max-width: none; margin: 0; padding: 0; }}
      .card {{ border-radius: 0; border: 1px solid #aaa; margin-bottom: 10px; }}
      h1 {{ font-size: 22px; }}
      h2 {{ font-size: 18px; }}
      h3 {{ font-size: 14px; }}
      @page {{ size: A4 portrait; margin: 12mm; }}
    }}
  </style>
</head>
<body>
    <div class="wrap">
    <h1>🧬 流式抗体配方报告</h1>
    <div class="subtitle">打印建议：A4 纵向，边距默认。生成时间：{generated_at}</div>
    <section class="meta-card">
      <div><span>实验日期：</span><strong>{experiment_date or '-'}</strong></div>
      <div><span>操作者：</span><strong>{operator or '-'}</strong></div>
      <div><span>实验批次：</span><strong>{batch_id or '-'}</strong></div>
      <div><span>样本数量：</span><strong>{sample_n}</strong></div>
    </section>
    {''.join(sections)}
  </div>

</body>
</html>
"""
    return html


def html_to_pdf_bytes(html_content: str) -> Optional[bytes]:
    """可选PDF导出：依赖 weasyprint；若环境未安装则返回 None。"""
    try:
        from weasyprint import HTML  # type: ignore

        return HTML(string=html_content).write_pdf()
    except Exception:
        return None




