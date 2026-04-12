# src/sweep_solver.py
import numpy as np
from .utils import euclidean_distance


def solve_sweep_heuristic(customers_data, depots, capacities):
    """
    使用扫描法 (Sweep Algorithm) 求解多车场初始分配
    :param customers_data: [x, y, demand, id]
    :return: labels, total_distance
    """
    num_customers = len(customers_data)
    num_depots = len(depots)

    # 1. 计算全局中心点 (Center of Gravity)
    # 以所有仓库的几何中心作为扫描雷达的原点
    cx = np.mean(depots[:, 0])
    cy = np.mean(depots[:, 1])

    # 2. 计算每个仓库的极坐标角度 [-pi, pi]
    depot_angles = np.zeros(num_depots)
    for j in range(num_depots):
        depot_angles[j] = np.arctan2(depots[j][1] - cy, depots[j][0] - cx)

    # 3. 计算每个客户的极坐标角度，并排序 (这就是"扫描"的过程)
    customer_angles = []
    for i in range(num_customers):
        x, y = customers_data[i][:2]
        angle = np.arctan2(y - cy, x - cx)
        customer_angles.append((angle, i))

    # 按角度从小到大排序 (模拟雷达顺时针/逆时针扫描)
    customer_angles.sort(key=lambda item: item[0])

    # 4. 执行分配
    labels = np.full(num_customers, -1)
    current_loads = np.zeros(num_depots)
    demands = customers_data[:, 2]

    # 遍历扫描到的客户
    for angle, cust_idx in customer_angles:
        demand = demands[cust_idx]

        # 寻找在角度上最接近该客户，且容量足够的仓库
        best_depot = -1
        min_angle_diff = float('inf')

        for j in range(num_depots):
            if current_loads[j] + demand <= capacities[j]:
                # 计算角度差 (取圆周上的最短距离)
                diff = abs(angle - depot_angles[j])
                diff = min(diff, 2 * np.pi - diff)

                if diff < min_angle_diff:
                    min_angle_diff = diff
                    best_depot = j

        if best_depot != -1:
            labels[cust_idx] = best_depot
            current_loads[best_depot] += demand
        else:
            # 如果所有仓库都满了 (或者角度接近的满了，其他也满了)，暂时设为 -1
            # 在真实场景中，如果总容量>总需求，通常能分完。
            # 如果分不完，随便找个还能装的塞进去
            for j in range(num_depots):
                if current_loads[j] + demand <= capacities[j]:
                    labels[cust_idx] = j
                    current_loads[j] += demand
                    break

    # 5. 计算初始总成本 (直线距离)
    total_dist = 0
    for i in range(num_customers):
        if labels[i] != -1:
            total_dist += euclidean_distance(customers_data[i][:2], depots[labels[i]])

    return labels, total_dist