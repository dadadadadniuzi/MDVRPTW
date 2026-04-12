# src/clustering.py
# 用于解决多车场问题的分配函数
import numpy as np
from .utils import euclidean_distance

def assign_nearest_neighbor(depots, customers_data):
    """
    (Phase 1 使用) 纯距离分配，不考虑容量
    """
    labels = []
    total_dist = 0
    for customer in customers_data:
        cust_coord = customer[:2]
        dists = [euclidean_distance(cust_coord, center) for center in depots]
        min_idx = np.argmin(dists)
        labels.append(min_idx)
        total_dist += dists[min_idx]
    return np.array(labels), total_dist


def assign_weighted_balanced(depots, customers_data, weights):
    """
    (Phase 2 使用) 根据权重分配，有惩罚系数
    """
    labels = []
    total_dist_real = 0
    for customer in customers_data:
        cust_coord = customer[:2]
        # 加权距离：物理距离 * 权重因子
        weighted_dists = [euclidean_distance(cust_coord, d) * weights[i] for i, d in enumerate(depots)]
        best_idx = np.argmin(weighted_dists)
        labels.append(best_idx)
        total_dist_real += euclidean_distance(cust_coord, depots[best_idx])
    return np.array(labels), total_dist_real


def calculate_fitness_load_balance(weights, customers_data, depots, capacities):
    """
    适应度函数：限制最大容量capacity，设置了惩罚函数
    """
    labels, real_dist_sum = assign_weighted_balanced(depots, customers_data, weights)

    demands = customers_data[:, 2]
    depot_loads = np.zeros(len(depots))
    for i, label in enumerate(labels):
        depot_loads[label] += demands[i]

    penalty = 0
    # 策略：惩罚系数设为极大，并且使用平方惩罚，让算法对超载极其敏感
    penalty_factor = 1000000

    for i in range(len(depots)):
        limit = capacities[i]
        load = depot_loads[i]
        if load > limit:
            # 使用平方惩罚：超载越多，惩罚呈指数级增长
            penalty += ((load - limit) ** 2) * penalty_factor

    return real_dist_sum + penalty

# =========================================================
# 新增：时间可行性检查函数
# =========================================================
def check_time_feasibility(customer, depot, depot_idx, tw_data, vehicles_data):
    """
    检查客户是否能被该仓库服务 (极限可行性)
    :param customer: [x, y, demand]
    :param depot: [x, y]
    :param depot_idx: 仓库索引
    :param tw_data: 时间窗数据 {'depots':..., 'customers':...}
    :param vehicles_data: 车辆数据 {0: [...], 1: [...]}
    :return: True (可行), False (不可行)
    """
    cid = int(customer[3])  # 假设 customer 数据传入时带上了 ID (在 repair_solution 中处理)

    # 1. 获取时间窗
    d_tw = tw_data['depots'].get(depot_idx)
    c_tw = tw_data['customers'].get(cid)

    if not d_tw or not c_tw:
        return True  # 如果没有时间窗数据，默认可行

    # 2. 获取该仓库的最快车速
    v_list = vehicles_data.get(depot_idx, [])
    if not v_list:
        max_speed = 60.0  # 默认兜底
    else:
        max_speed = max([v['velocity'] for v in v_list])

    # 3. 计算极限时间
    dist_go = euclidean_distance(depot, customer[:2])
    dist_back = euclidean_distance(customer[:2], depot)

    travel_go = dist_go / max_speed
    travel_back = dist_back / max_speed

    # 到达时间 = max(仓库开门 + 行驶, 客户最早)
    arrival = max(d_tw['start'] + travel_go, c_tw['start'])

    # 离开时间 = 到达 + 服务
    departure = arrival + c_tw['service']

    # 回库时间
    return_time = departure + travel_back

    # 检查是否能在仓库关门前回来
    if return_time > d_tw['end']:
        return False

    return True

# =========================================================
# 贪婪修复算法 (The Repairman)
# =========================================================
def repair_solution(customers_data, depots, labels, capacities, tw_data=None, vehicles_data=None):
    """
    强制修复超载 + 时间不可行的方案
    :param customers_data: 包含 [x, y, demand, id] (注意: 必须包含ID)
    """
    #复制标签
    new_labels = labels.copy()
    demands = customers_data[:, 2]  # index 2 is demand
    num_depots = len(depots)
    #最大迭代次数
    max_moves = 2000
    moves = 0

    while moves < max_moves:
        # 1. 计算当前负载
        current_loads = np.zeros(num_depots)
        for i, label in enumerate(new_labels):
            current_loads[label] += demands[i]

        # 2. 找出“问题仓库” (超载 OR 包含不可服务客户)
        problematic_depots = set()

        # (A) 容量超载
        for i in range(num_depots):
            if current_loads[i] > capacities[i]:
                problematic_depots.add(i)

        # (B) 时间不可行 (Time Infeasible)
        # 我们需要遍历所有客户，检查是否有人在当前分配下“回不来”
        infeasible_customers = []  # [(cust_idx, current_depot_idx)]

        if tw_data and vehicles_data:
            for i, label in enumerate(new_labels):
                # customers_data[i] 必须包含 ID
                is_feasible = check_time_feasibility(customers_data[i], depots[label], label, tw_data, vehicles_data)
                if not is_feasible:
                    problematic_depots.add(label)
                    infeasible_customers.append(i)

        # 如果没有问题，成功！
        if len(problematic_depots) == 0:
            final_dist = 0
            for i, cust in enumerate(customers_data):
                center = depots[new_labels[i]]
                final_dist += euclidean_distance(cust[:2], center)
            return new_labels, True, final_dist

        # 3. 修复策略
        # 优先处理“时间不可行”的客户，因为这是硬伤
        if infeasible_customers:
            target_cust_idx = infeasible_customers[0]  # 取第一个处理
            src_idx = new_labels[target_cust_idx]

            # 尝试移到其他仓库
            best_dst = -1
            min_cost = float('inf')

            for dst_idx in range(num_depots):
                if dst_idx == src_idx: continue

                # 检查移动到 dst_idx 是否【时间可行】
                if tw_data and vehicles_data:
                    if not check_time_feasibility(customers_data[target_cust_idx], depots[dst_idx], dst_idx, tw_data,
                                                  vehicles_data):
                        continue  # 这个仓库也不行，跳过

                # 检查移动后，dst_idx 是否会【容量严重超载】
                # (允许轻微超载，后续再修容量，但不能太离谱)
                if current_loads[dst_idx] + demands[target_cust_idx] > capacities[dst_idx] * 1.2:
                    continue

                    # 计算代价
                dist = euclidean_distance(customers_data[target_cust_idx][:2], depots[dst_idx])
                if dist < min_cost:
                    min_cost = dist
                    best_dst = dst_idx

            if best_dst != -1:
                new_labels[target_cust_idx] = best_dst
                moves += 1
                # print(f"  [时间修复] 客户 {int(customers_data[target_cust_idx][3])} 从 Depot {src_idx+1} 移至 Depot {best_dst+1}")
                continue
            else:
                # 所有仓库都不可行 -> 标记为"Unservable" (设为 -1)
                print(f"警告：客户 {int(customers_data[target_cust_idx][3])} 无法被任何仓库在时间窗内服务！")
                new_labels[target_cust_idx] = -1
                moves += 1
                continue

        # 如果没有时间问题，处理容量超载 (同之前的逻辑)
        # 找一个超载最严重的仓库
        overloaded = [d for d in problematic_depots if current_loads[d] > capacities[d]]
        if overloaded:
            src_idx = overloaded[0]  # 简单取第一个

            # 找该仓库里最适合移走的客户
            best_move = None
            min_cost_increase = float('inf')

            src_cust_indices = np.where(new_labels == src_idx)[0]
            for cust_idx in src_cust_indices:
                # 尝试移到有空位的仓库
                for dst_idx in range(num_depots):
                    if dst_idx == src_idx: continue

                    if current_loads[dst_idx] + demands[cust_idx] <= capacities[dst_idx]:
                        # 【新增】必须检查移动过去后时间是否可行
                        if tw_data and vehicles_data:
                            if not check_time_feasibility(customers_data[cust_idx], depots[dst_idx], dst_idx, tw_data,
                                                          vehicles_data):
                                continue

                        curr_dist = euclidean_distance(customers_data[cust_idx][:2], depots[src_idx])
                        new_dist = euclidean_distance(customers_data[cust_idx][:2], depots[dst_idx])
                        increase = new_dist - curr_dist

                        if increase < min_cost_increase:
                            min_cost_increase = increase
                            best_move = (cust_idx, dst_idx)

            if best_move:
                c_idx, d_idx = best_move
                new_labels[c_idx] = d_idx
                moves += 1
            else:
                # 陷入死胡同 (容量满了且时间也卡死了)
                return new_labels, False, 0

    return new_labels, False, 0