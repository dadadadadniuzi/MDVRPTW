# src/gap_solver.py
import numpy as np
from .utils import euclidean_distance
from .clustering import check_time_feasibility


def solve_gap_heuristic(customers_data, depots, capacities, tw_data=None, vehicles_data=None):
    """
    使用基于后悔值的贪婪启发式算法求解带时间窗的广义指派问题 (GAP)
    :param customers_data: [x, y, demand, id]
    :return: labels (每个客户分配的仓库索引)
    """
    num_customers = len(customers_data)
    num_depots = len(depots)

    # 1. 计算成本矩阵 (Cost Matrix)
    # cost[i][j] 表示把客户 i 分给仓库 j 的代价 (距离)
    # 如果时间窗绝对不可行，代价设为无穷大
    costs = np.zeros((num_customers, num_depots))

    for i in range(num_customers):
        for j in range(num_depots):
            dist = euclidean_distance(customers_data[i][:2], depots[j])

            # 时间窗预检：如果连专车都跑不完，直接判死刑
            is_feasible = True
            if tw_data and vehicles_data:
                is_feasible = check_time_feasibility(customers_data[i], depots[j], j, tw_data, vehicles_data)

            if is_feasible:
                costs[i][j] = dist
            else:
                costs[i][j] = float('inf')

    # 2. 初始化分配状态
    labels = np.full(num_customers, -1)  # -1 表示未分配
    current_loads = np.zeros(num_depots)
    demands = customers_data[:, 2]

    unassigned = set(range(num_customers))

    # 3. 迭代分配 (基于后悔值 Regret)
    while unassigned:
        # 计算所有未分配客户的"后悔值"
        regrets = {}
        valid_options_count = {}  # 记录该客户还有几个可行的仓库

        for i in unassigned:
            # 找出当前还能容纳该客户的仓库，且时间可行
            valid_costs = []
            for j in range(num_depots):
                if current_loads[j] + demands[i] <= capacities[j] and costs[i][j] != float('inf'):
                    valid_costs.append((costs[i][j], j))

            valid_options_count[i] = len(valid_costs)

            if len(valid_costs) == 0:
                # 这个客户哪个仓库都塞不下了 (超载或时间冲突)
                regrets[i] = float('inf')
            elif len(valid_costs) == 1:
                # 只剩唯一选择，必须马上分配，后悔值极大
                regrets[i] = 999999.0
            else:
                # 按成本从小到大排序
                valid_costs.sort(key=lambda x: x[0])
                # 后悔值 = 第二优的成本 - 最优的成本
                # 差距越大，说明如果不分给最优的，损失越惨重
                regret_val = valid_costs[1][0] - valid_costs[0][0]
                regrets[i] = regret_val

        # 找出后悔值最大的客户
        # 如果有多个无穷大(无解的)，随便挑一个先处理(后续交给repair)
        target_customer = max(regrets.keys(), key=lambda k: regrets[k])

        # 找出他最理想的仓库 (成本最低且能装下)
        best_depot = -1
        min_cost = float('inf')
        for j in range(num_depots):
            if current_loads[j] + demands[target_customer] <= capacities[j] and costs[target_customer][j] < min_cost:
                min_cost = costs[target_customer][j]
                best_depot = j

        if best_depot != -1:
            # 成功分配
            labels[target_customer] = best_depot
            current_loads[best_depot] += demands[target_customer]
            unassigned.remove(target_customer)
        else:
            # 实在分不进去了 (容量满了)
            # 强行分给他距离最近、时间可行的仓库 (允许它暂时超载，后续交由 Repair 处理)
            # 或者找个超载最少的
            best_depot_fallback = -1
            min_cost_fallback = float('inf')
            for j in range(num_depots):
                if costs[target_customer][j] < min_cost_fallback:
                    min_cost_fallback = costs[target_customer][j]
                    best_depot_fallback = j

            if best_depot_fallback != -1:
                labels[target_customer] = best_depot_fallback
                current_loads[best_depot_fallback] += demands[target_customer]
                unassigned.remove(target_customer)
            else:
                # 连时间可行的仓库都没有 (彻底无解)
                labels[target_customer] = -1
                unassigned.remove(target_customer)

    # 4. 计算总成本
    total_dist = 0
    for i in range(num_customers):
        if labels[i] != -1:
            total_dist += euclidean_distance(customers_data[i][:2], depots[labels[i]])

    return labels, total_dist