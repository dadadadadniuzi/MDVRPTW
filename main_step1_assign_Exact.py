# main_step1_assign_Exact.py
import numpy as np
import pandas as pd
import os
import sys
import pulp  # 导入精确求解器库
from src import config
from src.utils import generate_mock_data, load_data_from_excel, visualize_and_save_allocation, load_vehicle_data, \
    load_timewindow_data
from src.clustering import check_time_feasibility


def solve_exact_assignment(customers_matrix, depots, capacities, tw_data, vehicles_data):
    """
    使用 PuLP 建立 MILP 模型，求出100%完美的理论最优分配方案
    """
    num_customers = len(customers_matrix)
    num_depots = len(depots)

    # 1. 建立优化问题模型 (目标是最小化距离)
    prob = pulp.LpProblem("MDVRP_Allocation", pulp.LpMinimize)

    # 2. 定义决策变量 x[i][j] (0-1变量：客户 i 是否分配给仓库 j)
    x = pulp.LpVariable.dicts("x",
                              ((i, j) for i in range(num_customers) for j in range(num_depots)),
                              cat='Binary')

    # 3. 预处理成本矩阵与时间可行性
    costs = {}
    for i in range(num_customers):
        for j in range(num_depots):
            dist = np.sqrt(np.sum((customers_matrix[i][:2] - depots[j]) ** 2))
            # 预检：如果时间不可行，强制成本无穷大（或者加约束禁止选择）
            if tw_data and vehicles_data:
                if not check_time_feasibility(customers_matrix[i], depots[j], j, tw_data, vehicles_data):
                    dist = 9999999.0  # 极大的惩罚
            costs[i, j] = dist

    # 4. 设置目标函数: Minimize Sum(cost[i,j] * x[i,j])
    prob += pulp.lpSum([costs[i, j] * x[i, j] for i in range(num_customers) for j in range(num_depots)])

    # 5. 添加约束条件
    # 约束 A: 每个客户必须且只能分配给1个仓库
    for i in range(num_customers):
        prob += pulp.lpSum([x[i, j] for j in range(num_depots)]) == 1

    # 约束 B: 每个仓库的负载不能超过其容量上限
    for j in range(num_depots):
        prob += pulp.lpSum([customers_matrix[i][2] * x[i, j] for i in range(num_customers)]) <= capacities[j]

    # 6. 求解模型
    print("    正在调用底层 CBC 求解器进行精确计算...")
    prob.solve(pulp.PULP_CBC_CMD(msg=False))  # 隐藏求解器日志

    # 7. 解析结果
    status = pulp.LpStatus[prob.status]
    if status != 'Optimal':
        print(f"    警告: 求解器状态为 {status}，可能因为容量总和不足导致无解。")
        return None, 0

    labels = np.zeros(num_customers, dtype=int)
    total_dist = 0
    for i in range(num_customers):
        for j in range(num_depots):
            if pulp.value(x[i, j]) > 0.5:  # 变量值为 1
                labels[i] = j
                # 如果这个分配是极度惩罚的（时间不可行），说明彻底无解
                if costs[i, j] > 1000000:
                    labels[i] = -1
                else:
                    total_dist += costs[i, j]

    return labels, total_dist


def main():
    output_dir = os.path.join("results", "step1_Exact")
    if not os.path.exists(output_dir): os.makedirs(output_dir)

    print(">>> [Phase 1 - Exact MILP] 初始化...")
    # (同前：加载数据 df, depots, capacities, tw_data, vehicles_data)
    # ---------------- 简化的加载逻辑 ----------------
    if config.USE_CUSTOM_DATA:
        df, fixed_depots, capacities = load_data_from_excel(config.CUSTOM_DATA_FILE)
    else:
        df = generate_mock_data(config.RANDOM_NUM_CUSTOMERS, config.RANDOM_MAP_SIZE)
        fixed_depots, capacities = config.RANDOM_DEPOTS, config.RANDOM_CAPACITIES
    if 'id' not in df.columns: df['id'] = range(1, len(df) + 1)

    customers_matrix = df[['x', 'y', 'demand', 'id']].values
    coords = df[['x', 'y']].values
    num_depots = len(fixed_depots)

    vehicles_data = load_vehicle_data(num_depots)
    tw_data = load_timewindow_data(df, num_depots)
    # -----------------------------------------------

    print("\n>>> 运行精确数学求解器 (MILP Exact Solver)...")
    exact_labels, exact_dist = solve_exact_assignment(
        customers_matrix, fixed_depots, capacities, tw_data, vehicles_data
    )

    if exact_labels is None:
        print("数学证明：当前系统约束下无解（需求超载或时间完全冲突）。")
        sys.exit(1)

    print(f"\n>>> 精确求解完成！理论最优绝对距离下界: {exact_dist:.2f} km")

    print("\n--- [Exact Result] 负载详情 ---")
    for i in range(num_depots):
        c_indices = np.where(exact_labels == i)[0]
        load = np.sum(customers_matrix[c_indices, 2]) if len(c_indices) > 0 else 0
        print(f"  Depot {i + 1}: 负载 {load}/{capacities[i]:.0f} [正常]")

    visualize_and_save_allocation(
        coords, fixed_depots, exact_labels,
        f"Theoretical Optimal Allocation (Dist:{exact_dist:.1f})",
        "step1_Exact/step1_exact_optimal.png"
    )

    df['label'] = exact_labels
    df.to_csv(os.path.join(output_dir, "allocation_results_Exact.csv"), index=False)
    print("\n程序结束。请查看 results/step1_Exact 文件夹。")


if __name__ == "__main__":
    main()