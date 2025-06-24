import pandas as pd#复制于第三版，用于测试网页版本
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def load_excel_staining_plan(file_path):
    df = pd.read_excel(file_path, engine='openpyxl')
    df['是否作为FMO'] = df['是否作为FMO'].fillna('').apply(lambda x: str(x).strip() == '是')
    df['抗体类型'] = df['一抗/二抗/胞内抗体'].fillna('一抗').apply(str.strip)
    return df

def adjust_fmo_generic(df_all, df_current):
    fmo_marker_total = df_all[df_all['是否作为FMO']]['marker'].tolist()
    df_extra = pd.DataFrame({'marker': list(set(fmo_marker_total) - set(df_current['marker']))})
    df_extra['荧光染料'] = ''
    df_extra['稀释比例'] = ''
    df_extra['是否作为FMO'] = True
    df_extra['抗体类型'] = '非当前抗体类型FMO'
    return pd.concat([df_current, df_extra], ignore_index=True)

def compute_staining(df_fmo_ref, df_dye_only, sample_n, volume_per_well=50):
    fmo_markers = df_fmo_ref[df_fmo_ref['是否作为FMO']]
    total_wells = sample_n + len(fmo_markers)
    total_volume = total_wells * volume_per_well

    step1_results = []
    dye_mix_info = {}
    total_fmo_final_vol = 0
    step3_results = []

    # Step 1
    for _, row in df_dye_only.iterrows():
        if row['抗体类型'] == '自发荧光':
            continue  # 跳过自发荧光染料（如 YFP）

        marker = row['marker']
        dilution = int(str(row['稀释比例']).split(':')[1])
        is_fmo = row['是否作为FMO']
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

        step1_results.append({
            'marker': marker,
            '荧光染料': row['荧光染料'],
            '稀释比例': row['稀释比例'],
            '是否为FMO': '是' if is_fmo else '否',
            '所需抗体体积 (μL)': dye_vol,
            '加入FSB体积 (μL)': fsb_vol,
            '稀释后总体积 (μL)': final_volume
        })

        dye_mix_info[marker] = {
            'dilution': dilution,
            'final_vol': final_volume,
            'used_in_mix': 0,
            'is_fmo': is_fmo
        }

    # Step 2
    step2_fsb = round(total_volume - total_fmo_final_vol - sum(
        info['final_vol'] for info in dye_mix_info.values() if not info['is_fmo']
    ), 2)
    step2_dyes = [m for m, info in dye_mix_info.items() if not info['is_fmo']]

    # Step 3
    for marker in fmo_markers['marker']:
        ab_type_series = df_fmo_ref[df_fmo_ref['marker'] == marker]['抗体类型']
        ab_type = ab_type_series.values[0] if not ab_type_series.empty else ''

        if ab_type == '自发荧光':
            # YFP这类：不稀释，但FMO时需染所有其他染料
            all_dyes = [m for m in dye_mix_info if not dye_mix_info[m]['is_fmo']]
            fsb_vol = round(volume_per_well - len(all_dyes), 2)
            fmo_entry = {'FMO通道': marker, '加入主mix体积 (μL)': fsb_vol}
            for m in all_dyes:
                fmo_entry[f'加 {m} 染料 (μL)'] = 1
                dye_mix_info[m]['used_in_mix'] += 1
            step3_results.append(fmo_entry)
            continue

        if marker not in dye_mix_info:
            continue  # 非本抗体类型的FMO，不在本mix中配置

        # 正常抗体类型的FMO配置
        other_fmo = [
            m for m in fmo_markers['marker']
            if m != marker and m in dye_mix_info
        ]
        fsb_vol = round(volume_per_well - len(other_fmo), 2)
        fmo_entry = {'FMO通道': marker, '加入主mix体积 (μL)': fsb_vol}
        for m in other_fmo:
            fmo_entry[f'加 {m} 染料 (μL)'] = 1
            dye_mix_info[m]['used_in_mix'] += 1
        step3_results.append(fmo_entry)

    # Step 4
    master_mix_used = sum(row['加入主mix体积 (μL)'] for row in step3_results)
    master_mix_remaining = round(total_volume - total_fmo_final_vol - master_mix_used, 2)

    dye_remaining = []
    for m, info in dye_mix_info.items():
        if info['is_fmo']:
            remain = round(info['final_vol'] - info['used_in_mix'], 2)
            if remain > 0:
                dye_remaining.append({'marker': m, '加入体积 (μL)': remain})
    dye_remaining.append({'marker': '剩余主mix', '加入体积 (μL)': master_mix_remaining})

    return (
        pd.DataFrame(step1_results),
        step2_fsb,
        step2_dyes,
        pd.DataFrame(step3_results),
        pd.DataFrame(dye_remaining)
    )

def export_to_single_sheet(results_dict, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = 'FMO 配制方案'

    for label, (step1_df, step2_fsb, step2_dyes, step3_df, step4_df) in results_dict.items():
        ws.append([f'【{label}】Step 1: 染料稀释'])
        for row in dataframe_to_rows(step1_df, index=False, header=True):
            ws.append(row)
        ws.append([])

        ws.append(['Step 2: 主mix配置'])
        ws.append(['FSB体积 (μL)', step2_fsb])
        ws.append(['加入染料', ', '.join(step2_dyes)])
        ws.append([])

        ws.append(['Step 3: FMO配置'])
        for row in dataframe_to_rows(step3_df, index=False, header=True):
            ws.append(row)
        ws.append([])

        ws.append(['Step 4: 剩余验证'])
        for row in dataframe_to_rows(step4_df, index=False, header=True):
            ws.append(row)
        ws.append([])

    wb.save(output_path)
