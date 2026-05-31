"""
兼容脚本：命令行方式运行计算。

说明：
- 核心逻辑统一维护在 staining_logic.py
- 本文件仅用于本地快速调试，避免与主逻辑重复
"""

import pandas as pd

from staining_logic import (
    adjust_fmo_generic,
    compute_staining,
    export_to_single_sheet,
    load_excel_staining_plan,
    validate_and_prepare_df,
)


if __name__ == "__main__":
    input_path = "staining_plan.xlsx"
    output_path = "staining_result.xlsx"
    sample_n = 51

    raw_df = load_excel_staining_plan(input_path)
    df, errors = validate_and_prepare_df(raw_df)

    if errors:
        raise ValueError("; ".join(errors))

    results = {}
    for ab_type in ["一抗", "二抗", "胞内抗体"]:
        df_sub = df[df["抗体类型"] == ab_type].copy()
        if not df_sub.empty:
            adjusted = adjust_fmo_generic(df, df_sub)
            results[ab_type] = compute_staining(adjusted, df_sub, sample_n)

    export_to_single_sheet(results, output_path)
    print(f"✅ 配制完成，结果保存于：{output_path}")

