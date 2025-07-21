import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
# import glob # glob 不再需要，因为我们使用文件对话框
import os
# import tkinter as tk
# from tkinter import filedialog

# 设置matplotlib支持中文字体
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
matplotlib.rcParams['axes.unicode_minus'] = False


def process_age_data(age, df, version, iugs_col, show_cols):
    """处理单个年代数据并返回结果DataFrame，如果无法处理则返回None"""
    print(f"\n--- 正在处理年代: {age} ---")
    def match_age_func(val):
        if pd.isna(val):
            return False
        try:
            return str(val).strip() == str(age).strip() or float(val) == float(age)
        except ValueError:
            return str(val).strip() == str(age).strip()
        except Exception:
            return False

    match = df[df[version].apply(match_age_func)]
    show_df_subset = df[show_cols]
    result_df = None

    if not match.empty:
        for idx, row in match.iterrows():
            labels = ["宇", "界", "系", "统", "阶"]
            info = []
            for i, col_label in enumerate(labels):
                if col_label in df.columns:
                    info.append(f"{row.get(col_label, '')}{labels[i]}")
            iugs_value = row.get(iugs_col, "")
            print(f"直接匹配: {' '.join(info)} {iugs_value}")
            display_df = show_df_subset.loc[[idx]].copy()
            result_df = display_df
            break 
    else:
        try:
            age_num = float(age)
        except ValueError:
            print(f'年代"{age}"无法识别为数字，无法插值。')
            return None

        version_numeric = pd.to_numeric(df[version], errors='coerce')
        valid_idx = version_numeric.notna()
        version_numeric_filtered = version_numeric[valid_idx]
        df_filtered = df[valid_idx]

        if version_numeric_filtered.empty:
            print(f'{version}列没有可用于插值的数字数据')
            return None

        lower_idx_series = version_numeric_filtered[version_numeric_filtered <= age_num]
        lower_idx = lower_idx_series.idxmax() if not lower_idx_series.empty else None
        upper_idx_series = version_numeric_filtered[version_numeric_filtered >= age_num]
        upper_idx = upper_idx_series.idxmin() if not upper_idx_series.empty else None

        if lower_idx is None or upper_idx is None:
            print(f'无法找到合适的上下界进行插值 (age_num={age_num})')
            return None
        
        if lower_idx == upper_idx and df_filtered.loc[lower_idx, version] == age_num:
            match_row_df = df_filtered.loc[[lower_idx]]
            labels = ["宇", "界", "系", "统", "阶"]
            info = []
            for i, col_label in enumerate(labels):
                if col_label in df.columns:
                    info.append(f"{match_row_df.iloc[0].get(col_label, '')}{labels[i]}")
            iugs_value = match_row_df.iloc[0].get(iugs_col, "")
            print(f"精确数值匹配: {' '.join(info)} {iugs_value}")
            result_df = show_df_subset.loc[[match_row_df.index[0]]].copy()
            result_df.loc[match_row_df.index[0], version] = age_num
            return result_df 
        elif lower_idx == upper_idx:
             print(f'无法找到合适的上下界进行插值，输入值 {age_num} 可能超出 {version} 列的数值范围。')
             return None

        lower_row = df_filtered.loc[lower_idx]
        upper_row = df_filtered.loc[upper_idx]
        lower_x = float(lower_row[version])
        upper_x = float(upper_row[version])
        lower_y_val = pd.to_numeric(lower_row[iugs_col], errors='coerce')
        upper_y_val = pd.to_numeric(upper_row[iugs_col], errors='coerce')

        if pd.isna(lower_y_val) or pd.isna(upper_y_val):
            print('上下界的IUGS 2023.9数据缺失或非数值，无法插值')
            return None
        
        lower_y = float(lower_y_val)
        upper_y = float(upper_y_val)
        ratio = (age_num - lower_x) / (upper_x - lower_x) if upper_x != lower_x else 0
        interp_value = lower_y + (upper_y - lower_y) * ratio
        
        iugs_numeric = pd.to_numeric(df[iugs_col], errors='coerce')
        valid_iugs_idx = iugs_numeric.notna()
        iugs_numeric_filtered = iugs_numeric[valid_iugs_idx]
        df_iugs_filtered = df[valid_iugs_idx]

        upper_value_candidates = iugs_numeric_filtered[iugs_numeric_filtered > interp_value]
        upper_row_for_table = None

        if upper_value_candidates.empty:
            print(f'IUGS 2023.9列中没有大于插值值{interp_value:.4f}的数据')
            if not iugs_numeric_filtered.empty:
                closest_value_idx = (iugs_numeric_filtered - interp_value).abs().idxmin()
                upper_row_for_table = df_iugs_filtered.loc[[closest_value_idx]]
                upper_value_display = upper_row_for_table.iloc[0][iugs_col]
                print(f"尝试使用IUGS 2023.9中最接近的值: {upper_value_display}")
            else:
                return None
        else:
            upper_value_final = upper_value_candidates.min()
            upper_idx_final_series = df_iugs_filtered[iugs_numeric_filtered == upper_value_final].index
            if upper_idx_final_series.empty:
                 print(f"无法找到值为 {upper_value_final} 的原始索引。")
                 return None
            upper_idx_final = upper_idx_final_series[0]
            upper_row_for_table = df_iugs_filtered.loc[[upper_idx_final]]
            upper_value_display = upper_value_final

        if upper_row_for_table is None:
            print("无法确定用于插值显示的对应行。")
            return None

        result_df = show_df_subset.loc[[upper_row_for_table.index[0]]].copy()
        result_df.loc[upper_row_for_table.index[0], version] = age_num 
        result_df.loc[upper_row_for_table.index[0], iugs_col] = interp_value 
        
        labels = ["宇", "界", "系", "统", "阶"]
        info_row = df.loc[upper_row_for_table.index[0]] 
        info = []
        for i, col_label in enumerate(labels):
            if col_label in df.columns:
                info.append(f"{info_row.get(col_label, '')}{labels[i]}")
        print(f"插值结果: {' '.join(info)} IUGS插值:{interp_value:.4f} (原上界IUGS:{upper_value_display}, 输入年代:{age_num})")
    
    return result_df

def main():
    # df 是主数据表，从 wjy.xlsx 读取
    df = pd.read_excel('wjy.xlsx') 
    df.columns = df.columns.str.strip()
    valid_cols = [col for col in df.columns if col and '误差' not in col and not col.startswith('Unnamed:')]
    df = df[valid_cols]

    print('可选年代表版本（来自主数据表 wjy.xlsx）：')
    for col in df.columns:
        print(col, end=' | ')
    print('\n')
    version = input('输入当前年代表版本（主数据表中的列名）：').strip()
    if version not in df.columns:
        print(f'未找到年代表版本：{version}')
        return

    # 使用命令行输入文件路径替代tkinter文件对话框
    print("\n请输入包含年代数据的Excel文件路径：")
    selected_excel_file_path = input("文件路径: ").strip()
    
    if not selected_excel_file_path:
        print("未输入任何文件路径。")
        return
    
    print(f"\n您输入的文件路径是: {selected_excel_file_path}")

    # 检查文件是否存在
    if not os.path.exists(selected_excel_file_path):
        print(f"错误：文件 {selected_excel_file_path} 不存在。")
        return

    ages_from_file = []
    age_column_name = ""
    source_excel_df_for_read = None # 用于读取原始数据
    try:
        source_excel_df_for_read = pd.read_excel(selected_excel_file_path)
        print(f"\n已读取文件: {os.path.basename(selected_excel_file_path)}")
        print("文件中的列名：")
        for col_name in source_excel_df_for_read.columns:
            print(col_name, end=' | ')
        print('\n')
        
        age_column_name = input("请输入包含年代数据的那一列的名称：").strip()
        if age_column_name not in source_excel_df_for_read.columns:
            print(f"在 {os.path.basename(selected_excel_file_path)} 中未找到列：{age_column_name}")
            return
            
        raw_ages = source_excel_df_for_read[age_column_name].astype(str).tolist()
        ages_from_file = [age_val.strip() for age_val in raw_ages if age_val.strip()] 

    except FileNotFoundError:
        print(f"错误：文件 {selected_excel_file_path} 未找到。")
        return
    except Exception as e:
        print(f"读取或处理Excel文件时发生错误: {e}")
        return

    if not ages_from_file:
        print(f"在文件 {os.path.basename(selected_excel_file_path)} 的 '{age_column_name}' 列中未找到有效年代数据。")
        return

    iugs_col = 'IUGS 2023.9'
    show_cols = ['宇', '界', '系', '统', '阶', version, iugs_col]
    show_cols = [col for col in show_cols if col in df.columns] # 确保所有列都存在于主df中

    results_dfs = []
    processed_iugs_map = {} # 用于存储原始年代字符串到处理后IUGS值的映射

    for age_str in ages_from_file:
        result_df = process_age_data(age_str, df, version, iugs_col, show_cols)
        if result_df is not None:
            results_dfs.append(result_df)
            if iugs_col in result_df.columns and not result_df.empty:
                 # 假设 process_age_data 返回的 DataFrame 只有一行相关数据
                processed_iugs_map[age_str] = result_df.iloc[0][iugs_col]

    if not results_dfs:
        print("\n所有输入年代均未能成功处理，无法生成表格。")
    else:
        final_table_df = pd.concat(results_dfs).reset_index(drop=True)
        print("\n--- 合并处理结果 ---")
        print(final_table_df.to_string())

        # 绘制表格图像
        fig, ax = plt.subplots(figsize=(12, max(3, len(final_table_df) * 0.5 + 1))) 
        ax.axis('tight')
        ax.axis('off')
        table = ax.table(cellText=final_table_df.values, colLabels=final_table_df.columns, loc='center', cellLoc='center')
        table.auto_set_font_size(False)
        table.set_fontsize(10)
        table.scale(1.2, 1.2)

        # 增强表格颜色区分度
        for (row, col), cell in table.get_celld().items():
            if row == 0:  # 表头行
                cell.set_facecolor("#4472C4")  # 深蓝色表头
                cell.set_text_props(color='white', fontweight='bold')
            else:  # 数据行
                # 设置交替行颜色（斑马条纹）
                if row % 2 == 0:
                    cell.set_facecolor("#E6F0FF")  # 浅蓝色
                else:
                    cell.set_facecolor("#FFFFFF")  # 白色
                
                # 高亮特定列
                if col < len(final_table_df.columns):
                    col_name = final_table_df.columns[col]
                    
                    # 根据列类型设置不同颜色
                    if col_name == version:
                        cell.set_facecolor("#90EE90")  # 浅绿色
                    elif col_name == iugs_col:
                        cell.set_facecolor("#FFD700")  # 金色
                    elif col_name in ['宇', '界', '系', '统', '阶']:
                        # 为不同级别的地质时期设置颜色
                        geo_colors = {"宇": "#FFF2CC", "界": "#FCE4D6", "系": "#DDEBF7", "统": "#E2EFDA", "阶": "#D9D2E9"}
                        cell.set_facecolor(geo_colors.get(col_name, "#FFFFFF"))
        
        plt.title(f'年代数据处理结果 (源文件: {os.path.basename(selected_excel_file_path)}, 年代列: {age_column_name})', fontsize=14)
        plt.show()

    # 写回处理结果到原始Excel文件
    if source_excel_df_for_read is not None and age_column_name:
        try:
            # 重新读取原始文件以确保我们基于最新版本进行修改（尽管这里可能不需要，因为我们有source_excel_df_for_read）
            # 但为了安全，如果文件可能在程序运行时被外部修改，重新读取是好的
            # 为简化，我们直接使用之前读取的 source_excel_df_for_read
            df_to_write_back = source_excel_df_for_read.copy()

            # 创建新的IUGS数据列，与原始数据对齐
            new_iugs_column_data = []
            # 获取原始Excel文件中指定列的（处理过的）字符串值列表
            original_age_values_in_excel_col = df_to_write_back[age_column_name].astype(str).str.strip().tolist()

            for original_age_val in original_age_values_in_excel_col:
                if original_age_val in processed_iugs_map:
                    new_iugs_column_data.append(processed_iugs_map[original_age_val])
                else:
                    new_iugs_column_data.append(pd.NA) # 或 None, '', np.nan 根据需要
            
            # 确定插入位置和新列名
            insert_loc = df_to_write_back.columns.get_loc(age_column_name) + 1
            new_column_name = f"{iugs_col} (处理结果)"
            
            # 如果新列名已存在，则添加后缀以避免冲突
            temp_new_col_name = new_column_name
            counter = 1
            while temp_new_col_name in df_to_write_back.columns:
                temp_new_col_name = f"{new_column_name}_{counter}"
                counter += 1
            new_column_name = temp_new_col_name

            df_to_write_back.insert(insert_loc, new_column_name, new_iugs_column_data)
            
            # 写回文件
            df_to_write_back.to_excel(selected_excel_file_path, index=False)
            print(f"\n处理结果已成功写入到文件 '{os.path.basename(selected_excel_file_path)}' 的 '{new_column_name}' 列。")

        except Exception as e:
            print(f"将结果写回到Excel文件时发生错误: {e}")

if __name__ == '__main__':
    main()