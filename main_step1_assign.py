# main_step1_assign.py
import numpy as np
import pandas as pd
import os
import sys
from src import config
# 引入工具函数
from src.utils import generate_mock_data, load_data_from_excel, visualize_and_save_allocation, save_results_to_txt, \
    load_vehicle_data, load_timewindow_data
from src.dbo_algorithm import DBO
from src.clustering import assign_nearest_neighbor, assign_weighted_balanced, calculate_fitness_load_balance, \
    repair_solution, check_time_feasibility


# ==========================================
# 辅助函数：打印当前阶段的负载详情
# ==========================================
def print_stage_loads(stage_name, labels, customers_matrix, capacities):
    """
    打印指定阶段的每个仓库负载情况
    """
    print(f"\n--- [{stage_name}] 负载详情 ---")
    num_depots = len(capacities)

    # 遍历每个仓库
    for i in range(num_depots):
        # 找到属于该仓库的客户索引
        c_indices = np.where(labels == i)[0]

        # 计算总负载
        if len(c_indices) > 0:
            load = np.sum(customers_matrix[c_indices, 2])  # index 2 is demand
        else:
            load = 0

        limit = capacities[i]

        # 判断状态
        status = "正常"
        if load > limit:
            diff = load - limit
            status = f"超载! (+{diff:.1f})"

        # 打印
        print(f"  Depot {i + 1}: 负载 {load}/{limit:.0f} [{status}]")
    print("-----------------------------------")


def main():
    report_file = "optimization_process.txt"
    if os.path.exists(os.path.join("results", report_file)):
        os.remove(os.path.join("results", report_file))

    print(">>> [Phase 1] 初始化数据加载...")

    # ==========================
    # 1. 数据加载
    # ==========================
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

    # 确保有 ID 列
    if 'id' not in df.columns:
        df['id'] = range(1, len(df) + 1)

    customers_matrix = df[['x', 'y', 'demand', 'id']].values
    coords = df[['x', 'y']].values
    num_depots = len(fixed_depots)

    # 全局更新容量 (仅供参考)
    config.MAX_CAPACITY = np.mean(capacities)

    print(f"客户数量: {len(df)}")
    print(f"配送中心数量: {num_depots}")
    print(f"容量限制数组: {capacities}")

    # ==========================
    # 2. 预加载 Phase 2 数据 (用于时间可行性检查)
    # ==========================
    print(">>> [Phase 1] 预加载车辆与时间窗数据 (用于可行性检查)...")
    vehicles_data = load_vehicle_data(num_depots)
    tw_data = load_timewindow_data(df, num_depots)

    # ==========================
    # Phase 1: K-means (基准)
    # ==========================
    print("\n>>> Phase 1: 运行基础 K-means...")
    initial_weights = np.ones(num_depots)
    km_labels, km_dist = assign_nearest_neighbor(fixed_depots, customers_matrix)

    # 【新增】打印 Phase 1 负载
    print_stage_loads("Phase 1: K-means", km_labels, customers_matrix, capacities)

    # 【保存 Phase 1 结果与图片】
    save_results_to_txt(report_file, "Phase 1: 基础 K-means 结果", fixed_depots, km_labels, df, km_dist)
    visualize_and_save_allocation(
        coords, fixed_depots, km_labels,
        f"Phase 1: K-means (Dist:{km_dist:.1f})",
        "step1_kmeans.png"  # <--- 第一张图
    )



    # ==========================
    # Phase 2: DBO 优化 (中间态)
    # ==========================
    print("\n>>> Phase 2: DBO 优化 (寻找近似解)...")

    def fitness_wrapper(weights):
        # 仅优化容量
        return calculate_fitness_load_balance(weights, customers_matrix, fixed_depots, capacities)

    dbo = DBO(
        pop_size=config.DBO_POP_SIZE,
        dim=num_depots,
        lb=0.5, ub=3.0,
        max_iter=config.DBO_MAX_ITER,
        obj_func=fitness_wrapper,
        initial_guess=initial_weights
    )

    best_weights, _ = dbo.optimize()
    dbo_labels, dbo_dist = assign_weighted_balanced(fixed_depots, customers_matrix, best_weights)

    # 【新增】打印 Phase 2 负载
    print_stage_loads("Phase 2: DBO Result", dbo_labels, customers_matrix, capacities)

    # 【保存 Phase 2 结果与图片】(无论是否超载，作为来时路)
    save_results_to_txt(report_file, "Phase 2: DBO 优化初步结果", fixed_depots, dbo_labels, df, dbo_dist)
    visualize_and_save_allocation(
        coords, fixed_depots, dbo_labels,
        f"Phase 2: DBO Result (Dist:{dbo_dist:.1f})",
        "step2_dbo_raw.png"  # <--- 第二张图
    )

    # ==========================
    # Phase 3: 最终确认与修复
    # ==========================
    print("\n>>> Phase 3: 最终约束检查 (容量 + 时间可行性)...")

    final_labels = dbo_labels.copy()
    final_dist = dbo_dist
    status_msg = "Unknown"

    # 定义检查函数
    def check_full_feasibility(labels):
        # 容量检查
        loads = np.zeros(num_depots)
        for i, label in enumerate(labels):
            loads[label] += customers_matrix[i, 2]
        for i in range(num_depots):
            if loads[i] > capacities[i]:
                return False
        # 时间检查
        for i, label in enumerate(labels):
            customer = customers_matrix[i]
            if not check_time_feasibility(customer, fixed_depots[label], label, tw_data, vehicles_data):
                return False
        return True

    # 1. 检查 DBO 结果
    print("    正在验证 DBO 结果的可行性...")
    is_dbo_fully_valid = check_full_feasibility(dbo_labels)

    if is_dbo_fully_valid:
        print("    >>> DBO 结果已完美满足所有约束 (容量 & 时间)，无需修复！")
        status_msg = "DBO Perfect"
        # final_labels 保持不变
    else:
        print("    >>> DBO 结果存在约束违反，启动智能修复 (Smart Repair)...")

        repaired_labels, success, repaired_dist = repair_solution(
            customers_matrix,  # 包含 ID
            fixed_depots,
            dbo_labels,
            capacities,
            tw_data=tw_data,  # 传入时间窗
            vehicles_data=vehicles_data  # 传入车辆数据
        )

        if success:
            print("    >>> 修复成功！所有客户满足容量与时间约束。")
            final_labels = repaired_labels
            final_dist = repaired_dist
            status_msg = "Repaired Valid"
        else:
            print("    >>> 修复不完全 (可能存在不可服务客户或严重超载)。")
            final_labels = repaired_labels
            # 重新计算距离
            final_dist = 0
            for i, cust in enumerate(customers_matrix):
                if final_labels[i] != -1:
                    final_dist += np.sqrt(np.sum((cust[:2] - fixed_depots[final_labels[i]]) ** 2))
            status_msg = "Partially Invalid"

    # ==========================
    # 3. 保存最终结果 (包括图片和CSV)
    # ==========================

    # 保存文本
    save_results_to_txt(report_file, f"Phase 3: 最终分配结果 ({status_msg})", fixed_depots, final_labels, df,
                        final_dist)

    # 【保存 Phase 3 图片】(无论是否经过修复，这是最终结果)
    visualize_and_save_allocation(
        coords, fixed_depots, final_labels,
        f"Phase 3: Final Allocation ({status_msg}, Dist:{final_dist:.1f})",
        "step3_final_result.png"  # <--- 第三张图
    )

    # 保存 CSV (供 Phase 2 读取)
    df['label'] = final_labels
    output_csv = os.path.join("results", "allocation_results.csv")
    df.to_csv(output_csv, index=False)
    print(f"\n>>> [重要] 分配结果已保存至: {output_csv}")

    # 打印最终详情
    print("\n=== 最终分配详情 ===")
    for i in range(num_depots):
        c_indices = np.where(final_labels == i)[0]
        load = np.sum(customers_matrix[c_indices, 2])
        limit = capacities[i]

        status = "正常"
        if load > limit: status = f"超载! ({load - limit:.1f})"

        print(f"  Depot {i + 1}: 负载 {load}/{limit:.0f} [{status}]")

    print("\n程序结束。请查看 results/ 文件夹下的 3 张图片。")


if __name__ == "__main__":
    main()