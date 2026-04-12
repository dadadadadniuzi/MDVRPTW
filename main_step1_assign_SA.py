# main_step1_assign_SA.py
import numpy as np
import pandas as pd
import os
import sys
from src import config
from src.utils import generate_mock_data, load_data_from_excel, visualize_and_save_allocation, load_vehicle_data, \
    load_timewindow_data
from src.clustering import repair_solution, check_time_feasibility
# 导入新建的扫描法
from src.sweep_solver import solve_sweep_heuristic


def check_full_feasibility(labels, customers_matrix, depots, capacities, tw_data, vehicles_data):
    """校验是否所有约束都满足"""
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
    output_dir = os.path.join("results", "step1_SA")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    report_file = os.path.join(output_dir, "sa_optimization_process.txt")
    if os.path.exists(report_file):
        os.remove(report_file)

    print(">>> [Phase 1 - Sweep] 初始化数据加载...")

    # 1. 数据加载
    if config.USE_CUSTOM_DATA:
        try:
            df, fixed_depots, capacities = load_data_from_excel(config.CUSTOM_DATA_FILE)
        except Exception as e:
            print(f"错误: {e}")
            sys.exit(1)
    else:
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

    print(">>> 预加载车辆与时间窗数据 (用于可行性检查)...")
    vehicles_data = load_vehicle_data(num_depots)
    tw_data = load_timewindow_data(df, num_depots)

    # ==========================
    # Phase 1: 扫描法 (Sweep) 求解
    # ==========================
    print("\n>>> Phase 1: 运行扫描法 (Sweep Algorithm)...")

    sa_labels, sa_dist = solve_sweep_heuristic(customers_matrix, fixed_depots, capacities)

    # 打印 SA 初步负载
    print("\n--- [Phase 1: Sweep Result] 负载详情 ---")
    for i in range(num_depots):
        c_indices = np.where(sa_labels == i)[0]
        load = np.sum(customers_matrix[c_indices, 2]) if len(c_indices) > 0 else 0
        limit = capacities[i]
        status = "正常" if load <= limit else f"超载! (+{load - limit:.1f})"
        print(f"  Depot {i + 1}: 负载 {load}/{limit:.0f} [{status}]")

    # 保存 SA 结果文本
    with open(report_file, 'a', encoding='utf-8') as f:
        f.write("========== Phase 1: 扫描法 (Sweep) 初步结果 ==========\n")
        f.write(f"初步总距离: {sa_dist:.2f} km\n\n")

    visualize_and_save_allocation(
        coords, fixed_depots, sa_labels,
        f"Phase 1: Sweep Algorithm (Dist:{sa_dist:.1f})",
        "step1_SA/step1_sa_raw.png"
    )

    # ==========================
    # Phase 2: 最终确认与贪婪修复
    # ==========================
    print("\n>>> Phase 2: 最终约束检查 (容量 + 时间可行性)...")

    final_labels = sa_labels.copy()
    final_dist = sa_dist
    status_msg = "Unknown"

    print("    正在验证 Sweep 结果的可行性...")
    is_fully_valid = check_full_feasibility(sa_labels, customers_matrix, fixed_depots, capacities, tw_data,
                                            vehicles_data)

    if is_fully_valid:
        print("    >>> Sweep 结果已满足所有约束 (容量 & 时间)，无需修复！")
        status_msg = "Sweep Perfect"
    else:
        print("    >>> Sweep 结果存在约束违反 (主要由于未考虑时间窗)，启动智能修复...")

        repaired_labels, success, repaired_dist = repair_solution(
            customers_matrix, fixed_depots, sa_labels, capacities,
            tw_data=tw_data, vehicles_data=vehicles_data
        )

        if success:
            print("    >>> 修复成功！所有客户满足容量与时间约束。")
            final_labels = repaired_labels
            final_dist = repaired_dist
            status_msg = "Sweep Repaired"
        else:
            print("    >>> 修复不完全 (系统可能极度拥挤)。")
            final_labels = repaired_labels
            final_dist = 0
            for i, cust in enumerate(customers_matrix):
                if final_labels[i] != -1:
                    final_dist += np.sqrt(np.sum((cust[:2] - fixed_depots[final_labels[i]]) ** 2))
            status_msg = "Sweep Partially Invalid"

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

    with open(report_file, 'a', encoding='utf-8') as f:
        f.write(f"========== Phase 2: 最终修复结果 ({status_msg}) ==========\n")
        f.write(f"最终总距离: {final_dist:.2f} km\n")

    visualize_and_save_allocation(
        coords, fixed_depots, final_labels,
        f"Final Allocation: Sweep ({status_msg}, Dist:{final_dist:.1f})",
        "step1_SA/step2_sa_final.png"
    )

    # 核心：保存 CSV，供 Phase 2 读取
    df['label'] = final_labels
    output_csv = os.path.join(output_dir, "allocation_results_SA.csv")
    df.to_csv(output_csv, index=False)

    print(f"\n>>> [重要] 分配结果已保存至: {output_csv}")
    print("\n程序结束。请查看 results/step1_SA 文件夹。")


if __name__ == "__main__":
    main()