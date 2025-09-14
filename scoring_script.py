import os

def clean_old_files():
    for file in os.listdir():
        if file.endswith(".xlsx") and "评分结果" in file:
            try:
                os.remove(file)
            except Exception as e:
                print(f"⚠️ 无法删除文件 {file}：{e}")
import pandas as pd
import numpy as np
from scoring_rules import MALE_RULES, FEMALE_RULES
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side

def process_scores(file_path):
    print(f"📥 正在读取文件：{file_path}")
    # ✅ 清理旧评分文件
    clean_old_files()

    raw_df = pd.read_excel(file_path, header=None)

    header_indices = raw_df[raw_df.apply(lambda row: row.astype(str).str.contains('性别').any(), axis=1)].index.tolist()
    print(f"🔍 识别到 {len(header_indices)} 个表头段落")

    all_results = []
    time_projects = ['1500米', '800米']
    all_projects = list(MALE_RULES.keys())
    if '仰卧起坐' in all_projects and '引体向上' in all_projects:
        all_projects.remove('仰卧起坐')
        insert_index = all_projects.index('引体向上') + 1
        all_projects.insert(insert_index, '仰卧起坐')

    for i, header_idx in enumerate(header_indices):
        end_idx = header_indices[i + 1] if i + 1 < len(header_indices) else len(raw_df)
        segment = raw_df.iloc[header_idx:end_idx].reset_index(drop=True)
        segment.columns = segment.iloc[0]
        df = segment[1:].reset_index(drop=True)

        required_cols = ['姓名', '性别', '班级']
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            print(f"⚠️ 段落缺少字段：{missing}，跳过")
            continue

        df = df[df['性别'].isin(['男', '女'])].copy()
        if df.empty:
            print("⚠️ 段落无有效性别数据，跳过")
            continue

        result = df.copy()
        remarks = []

        for proj in all_projects:
            result[f'{proj}_得分'] = ""

        for idx, row in df.iterrows():
            gender = row['性别']
            rule_dict = MALE_RULES if gender == '男' else FEMALE_RULES
            score_values = []
            missing_items = []

            for proj in rule_dict:
                col_name = f'{proj}_得分'
                val = row.get(proj)

                if pd.isna(val):
                    result.at[idx, col_name] = "无"
                    missing_items.append(proj)
                    continue

                try:
                    val = float(val)
                    if proj in time_projects:
                        val *= 60
                except:
                    result.at[idx, col_name] = "无"
                    missing_items.append(f"{proj}(非数值)")
                    continue

                matched = False
                for low, up, pts in rule_dict[proj]:
                    if low <= val <= up:
                        result.at[idx, col_name] = pts
                        score_values.append(pts)
                        matched = True
                        break

                if not matched:
                    result.at[idx, col_name] = "无"
                    missing_items.append(f"{proj}(超范围)")

            if score_values:
                total = sum(score_values)
                avg = round(total / len(score_values), 2)
                result.at[idx, '总分'] = total
                result.at[idx, '平均分'] = avg
            else:
                result.at[idx, '总分'] = "无"
                result.at[idx, '平均分'] = "无"

            remarks.append("缺：" + "、".join(missing_items) if missing_items else "")

        result['备注'] = remarks
        cols = [col for col in result.columns if col != '备注'] + ['备注']
        result = result[cols]
        all_results.append(result)

    if not all_results:
        print("❌ 没有有效数据段落，评分失败")
        return None

    final_result = pd.concat(all_results, ignore_index=True)

    # ✅ 统一总表列顺序
    standard_columns = [
        '序号', '班级', '学号', '性别', '姓名',
        '引体向上', '仰卧起坐', '1分钟跳绳', '立定跳远', '抛实心球', '100米', '1500米', '800米',
        '引体向上_得分', '仰卧起坐_得分', '1分钟跳绳_得分', '立定跳远_得分', '抛实心球_得分', '100米_得分', '1500米_得分', '800米_得分',
        '总分', '平均分', '备注'
    ]

    for col in standard_columns:
        if col not in final_result.columns:
            final_result[col] = ""

    final_result = final_result[standard_columns]

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    total_file = f"总表_评分结果_{timestamp}.xlsx"
    final_result.to_excel(total_file, index=False)
    print(f"✅ 总表已保存：{total_file}")

    # ✅ 美化总表
    wb = load_workbook(total_file)
    ws = wb.active
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
            if cell.row == 1:
                cell.font = Font(bold=True)

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2

    wb.save(total_file)

    # ✅ 分班输出
    grouped = final_result.groupby('班级')
    for class_name, class_df in grouped:
        for col in standard_columns:
            if col not in class_df.columns:
                class_df[col] = ""
        class_df = class_df[standard_columns]

        safe_name = "".join(c if c.isalnum() or c in "_-" else "_" for c in str(class_name))
        file_name = f"{safe_name}_评分结果_{timestamp}.xlsx"
        class_df.to_excel(file_name, index=False)

        wb = load_workbook(file_name)
        ws = wb.active

        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border
                if cell.row == 1:
                    cell.font = Font(bold=True)

        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = max_length + 2

        wb.save(file_name)
        print(f"✅ 分班表已保存：{file_name}")

    print("🎉 所有评分文件已生成完毕")
    return total_file

if __name__ == "__main__":
    total_file = process_scores("raw_scores.xlsx")
    print("生成的总表文件：", total_file)
