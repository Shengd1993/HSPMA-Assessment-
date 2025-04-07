import pandas as pd
import numpy as np


def read_expert_data(file_paths):
    expert_scores = {}
    for i, file_path in enumerate(file_paths, start=1):
        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            expert_data = {}
            for sheet_name in sheet_names:
                df = excel_file.parse(sheet_name)
                df.set_index(df.columns[0], inplace=True)
                expert_data[sheet_name] = df
            expert_scores[f'expert_{i}'] = expert_data
            print(f"成功读取专家 {i} 的打分数据")
        except FileNotFoundError:
            print(f"错误：未找到专家 {i} 的文件 {file_path}。")
        except Exception as e:
            print(f"错误：读取专家 {i} 的文件 {file_path} 时发生未知错误：{e}")
    return expert_scores


def combine_matrices(expert_scores, sheet_name):
    matrices = [expert_scores[f'expert_{i}'][sheet_name] for i in range(1, 6)]
    combined_matrix = sum(matrices) / len(matrices)
    return combined_matrix


def ahp_weight(matrix):
    eigenvalues, eigenvectors = np.linalg.eig(matrix)
    max_eigenvalue = np.max(eigenvalues).real
    index = np.argmax(eigenvalues)
    eigenvector = eigenvectors[:, index].real
    weights = eigenvector / eigenvector.sum()
    n = matrix.shape[0]
    ri = {1: 0, 2: 0, 3: 0.58, 4: 0.90, 5: 1.12, 6: 1.24, 7: 1.32, 8: 1.41, 9: 1.45}
    ci = (max_eigenvalue - n) / (n - 1)
    cr = ci / ri[n]
    return weights, cr


def calculate_all_weights(expert_scores):
    # 计算一级指标权重
    combined_matrix_primary = combine_matrices(expert_scores, '一级指标判断矩阵')
    primary_weights, primary_cr = ahp_weight(combined_matrix_primary)
    primary_index_names = combined_matrix_primary.index.tolist()
    primary_weight_dict = dict(zip(primary_index_names, primary_weights))
    print("一级指标权重:")
    for index, weight in primary_weight_dict.items():
        print(f"{index}: {weight:.4f}")
    print(f"一级指标一致性比例 CR: {primary_cr:.4f}")

    # 计算二级指标权重
    secondary_weight_dict = {}
    for primary_index in primary_index_names:
        sheet_name = f'{primary_index}二级指标判断矩阵'
        combined_matrix_secondary = combine_matrices(expert_scores, sheet_name)
        secondary_weights, secondary_cr = ahp_weight(combined_matrix_secondary)
        secondary_index_names = combined_matrix_secondary.index.tolist()
        secondary_weight_dict[primary_index] = dict(zip(secondary_index_names, secondary_weights))
        print(f"\n{primary_index} 二级指标权重:")
        for index, weight in secondary_weight_dict[primary_index].items():
            print(f"{index}: {weight:.4f}")
        print(f"{primary_index} 二级指标一致性比例 CR: {secondary_cr:.4f}")

    # 计算总权重
    total_weight_dict = {}
    for primary_index, primary_weight in primary_weight_dict.items():
        for secondary_index, secondary_weight in secondary_weight_dict[primary_index].items():
            total_weight = primary_weight * secondary_weight
            total_weight_dict[(primary_index, secondary_index)] = total_weight
    print("\n总权重:")
    for (primary_index, secondary_index), total_weight in total_weight_dict.items():
        print(f"{primary_index} - {secondary_index}: {total_weight:.4f}")

    return primary_weight_dict, secondary_weight_dict, total_weight_dict


def save_weights_to_excel(primary_weight_dict, secondary_weight_dict, total_weight_dict):
    # 创建一个 Excel 写入对象
    writer = pd.ExcelWriter('weights_result.xlsx', engine='openpyxl')

    # 保存一级指标权重到 Excel
    primary_df = pd.DataFrame.from_dict(primary_weight_dict, orient='index', columns=['权重'])
    primary_df.index.name = '一级指标'
    primary_df.to_excel(writer, sheet_name='一级指标权重')

    # 保存二级指标权重到 Excel
    for primary_index, secondary_weights in secondary_weight_dict.items():
        secondary_df = pd.DataFrame.from_dict(secondary_weights, orient='index', columns=['权重'])
        secondary_df.index.name = '二级指标'
        secondary_df.to_excel(writer, sheet_name=f'{primary_index}二级指标权重')

    # 保存总权重到 Excel
    index = pd.MultiIndex.from_tuples(total_weight_dict.keys(), names=['一级指标', '二级指标'])
    total_df = pd.DataFrame(list(total_weight_dict.values()), index=index, columns=['总权重'])
    total_df.to_excel(writer, sheet_name='总权重')

    # 保存 Excel 文件
    writer.close()
    print("权重结果已保存到 weights_result.xlsx")


# 定义 5 个专家的 Excel 文件路径列表
file_paths = [
    'judgment_matrix_1.xlsx',
    'judgment_matrix_2.xlsx',
    'judgment_matrix_3.xlsx',
    'judgment_matrix_4.xlsx',
    'judgment_matrix_5.xlsx'
]

# 读取数据
expert_scores = read_expert_data(file_paths)

# 计算权重
primary_weight_dict, secondary_weight_dict, total_weight_dict = calculate_all_weights(expert_scores)

# 保存权重结果到 Excel
save_weights_to_excel(primary_weight_dict, secondary_weight_dict, total_weight_dict)
