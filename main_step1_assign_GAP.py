# main_step1_assign_GAP.py
import numpy as np
import pandas as pd
import os
import sys
from src import config
from src.utils import generate_mock_data, load_data_from_excel, visualize_and_save_allocation, save_results_to_txt, \
    load_vehicle_data, load_timewindow_data
from src.clustering import repair_solution, check_time_feasibility
# 导入我们新建的 GAP 求解器
from src.gap_solver import solve_gap_heuristic


def check_full_feasibility(labels, customers_matrix, depots, capacities, tw_data, vehicles_data):
    num_depots = len(depots)
    loads = np.zeros(num_depots)
    for i, label in enumerate(labels):
        if label != -1:
            loads[label] += customers_matrix[i, 2]

    for i in range(num_depots):
        if loads[i] > capacities[i]: return False

    for i, label in enumerate(labels):
        if label != -1:
            customer = customers_matrix[i]
            if not check_time_feasibility(customer, depots[label], label, tw_data, vehicles_data):
                return False
    return True


def main():
    # 0. 准备独立输出文件夹
    output_dir = os.path.join("results", "step1_GAP")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    report_file = os.path.join(output_dir, "gap_optimization_process.txt")
    if os.path.exists(report_file):
        os.remove(report_file)

    print(">>> [Phase 1 - GAP] 初始化数据加载...")

    # 1. 数据加载 (同之前)
    if config.USE_CUSTOM_DATA:
        print(f">>> 模式：使用 Excel 数据 ({config.CUSTOM_DATA_FILE})")
        try:
            df, fixed_depots, capacities = load_data_from_excel(config.CUSTOM_DATA_FILE)
        except Exception as e:
            print(f"错误: {e}")
            sys.exit(1)
    else:
        print(">>> 模式：使用随机生成数据")
        df = generate_mock_data(config.RANDOM_NUM_CUSTOMERS, config.RANDOM_MAP_SIZE)
        fixed_depots = config.RANDOM_DEPOTS
        capacities = config.RANDOM_CAPACITIES
        if len(capacities) != len(fixed_depots):
            capacities = np.full(len(fixed_depots), capacities[0])

    if 'id' not in df.columns:
        df['id'] = range(1, len(df) + 1)

    customers_matrix = df[['x', 'y', 'demand', 'id']].values
    coords = df[['x', 'y']].values
    num_depots = len(fixed_depots)

    print(f"客户数量: {len(df)}")
    print(f"配送中心数量: {num_depots}")
    print(f"容量限制数组: {capacities}")

    print(">>> 预加载车辆与时间窗数据 (用于可行性检查)...")
    vehicles_data = load_vehicle_data(num_depots)
    tw_data = load_timewindow_data(df, num_depots)

    # ==========================
    # Phase 1: 广义指派法 (GAP) 求解
    # ==========================
    print("\n>>> Phase 1: 运行基于后悔值的 GAP 启发式算法...")

    gap_labels, gap_dist = solve_gap_heuristic(
        customers_matrix, fixed_depots, capacities, tw_data, vehicles_data
    )

    # 打印 GAP 初步负载
    print("\n--- [Phase 1: GAP Result] 负载详情 ---")
    for i in range(num_depots):
        c_indices = np.where(gap_labels == i)[0]
        load = np.sum(customers_matrix[c_indices, 2]) if len(c_indices) > 0 else 0
        limit = capacities[i]
        status = "正常" if load <= limit else f"超载! (+{load - limit:.1f})"
        print(f"  Depot {i + 1}: 负载 {load}/{limit:.0f} [{status}]")

    # 保存 GAP 结果文本 (此处使用了一个简单粗暴的追加写方法，如果你愿意也可以复用 utils 的函数，只需修改路径)
    with open(report_file, 'a', encoding='utf-8') as f:
        f.write("========== Phase 1: 广义指派法 (GAP) 初步结果 ==========\n")
        f.write(f"初步总距离: {gap_dist:.2f} km\n\n")

    visualize_and_save_allocation(
        coords, fixed_depots, gap_labels,
        f"Phase 1: GAP Heuristic (Dist:{gap_dist:.1f})",
        "step1_GAP/step1_gap_raw.png"  # 存在子文件夹中
    )

    # ==========================
    # Phase 2: 最终确认与贪婪修复
    # ==========================
    print("\n>>> Phase 2: 最终约束检查 (容量 + 时间可行性)...")

    final_labels = gap_labels.copy()
    final_dist = gap_dist
    status_msg = "Unknown"

    print("    正在验证 GAP 结果的可行性...")
    is_fully_valid = check_full_feasibility(gap_labels, customers_matrix, fixed_depots, capacities, tw_data,
                                            vehicles_data)

    if is_fully_valid:
        print("    >>> GAP 结果已完美满足所有约束 (容量 & 时间)，无需修复！")
        status_msg = "GAP Perfect"
    else:
        print("    >>> GAP 结果存在约束违反 (可能是强行分配导致的超载)，启动智能修复...")

        repaired_labels, success, repaired_dist = repair_solution(
            customers_matrix, fixed_depots, gap_labels, capacities,
            tw_data=tw_data, vehicles_data=vehicles_data
        )

        if success:
            print("    >>> 修复成功！所有客户满足容量与时间约束。")
            final_labels = repaired_labels
            final_dist = repaired_dist
            status_msg = "GAP Repaired"
        else:
            print("    >>> 修复不完全 (系统可能极度拥挤)。")
            final_labels = repaired_labels
            final_dist = 0
            for i, cust in enumerate(customers_matrix):
                if final_labels[i] != -1:
                    final_dist += np.sqrt(np.sum((cust[:2] - fixed_depots[final_labels[i]]) ** 2))
            status_msg = "GAP Partially Invalid"

    # ==========================
    # 结果保存与展示
    # ==========================
    print("\n=== 最终分配详情 ===")
    for i in range(num_depots):
        c_indices = np.where(final_labels == i)[0]
        load = np.sum(customers_matrix[c_indices, 2]) if len(c_indices) > 0 else 0
        limit = capacities[i]
        status = "正常" if load <= limit else f"超载! (+{load - limit:.1f})"
        print(f"  Depot {i + 1}: 负载 {load}/{limit:.0f} [{status}]")

    unservable = np.sum(final_labels == -1)
    if unservable > 0:
        print(f"  [警告] 有 {unservable} 个客户无法被任何仓库服务！")

    # 写入最终结果文本
    with open(report_file, 'a', encoding='utf-8') as f:
        f.write(f"========== Phase 2: 最终修复结果 ({status_msg}) ==========\n")
        f.write(f"最终总距离: {final_dist:.2f} km\n")

    visualize_and_save_allocation(
        coords, fixed_depots, final_labels,
        f"Final Allocation: GAP ({status_msg}, Dist:{final_dist:.1f})",
        "step1_GAP/step2_gap_final.png"
    )

    # 核心：保存给 Phase 2 用的 CSV 文件
    # 注意这里保存的文件名必须和 ACO 预期的一样，但为了区分，我们存在特定的文件夹里
    df['label'] = final_labels
    output_csv = os.path.join(output_dir, "allocation_results_GAP.csv")
    df.to_csv(output_csv, index=False)

    print(f"\n>>> [重要] 分配结果已保存至: {output_csv}")
    print("    你可以修改 Phase 2 (路径规划) 代码去读取这个新的 CSV 文件。")
    print("\n程序结束。请查看 results/step1_GAP 文件夹。")


if __name__ == "__main__":
    main()