# src/aco_vrptw.py
import numpy as np
import copy
from .utils import euclidean_distance
from . import config


class ACO_VRPTW:
    def __init__(self, depot_coord, customers_data, vehicle_list, time_windows, depot_tw, original_ids):
        self.depot = depot_coord
        self.customers = customers_data
        self.num_customers = len(customers_data)
        self.num_nodes = self.num_customers + 1

        # 车辆池模板
        self.vehicle_templates = vehicle_list

        self.tws = time_windows
        self.depot_tw = depot_tw
        self.orig_ids = original_ids

        # 初始化距离矩阵
        self.dist_mat = np.zeros((self.num_nodes, self.num_nodes))
        all_coords = np.vstack([self.depot, self.customers[:, :2]])
        for i in range(self.num_nodes):
            for j in range(self.num_nodes):
                self.dist_mat[i][j] = euclidean_distance(all_coords[i], all_coords[j])

        # 初始化信息素和启发因子
        self.pheromone = np.ones((self.num_nodes, self.num_nodes))
        self.heuristic = 1.0 / (self.dist_mat + 1e-10)

    def run(self):
        best_routes = None
        best_cost = float('inf')
        convergence = []  # 记录收敛过程

        for iter in range(config.ACO_ITERATIONS):
            ant_routes = []
            ant_costs = []

            for k in range(config.ACO_ANT_COUNT):
                routes, cost = self.construct_solution()
                ant_routes.append(routes)
                ant_costs.append(cost)

                if cost < best_cost:
                    best_cost = cost
                    best_routes = routes
                    convergence.append(best_cost)

            # 更新信息素
            self.update_pheromone(ant_routes, ant_costs)

            # 简单日志
            # if (iter + 1) % 10 == 0:
            #     print(f"    Iter {iter+1}: Best Cost {best_cost:.2f}")

        # 【修复点】必须返回 tuple，即使 best_routes 是 None
        return best_routes, best_cost, convergence

    def construct_solution(self):
        unvisited = set(range(1, self.num_nodes))
        routes = []
        total_cost = 0

        # 复制车辆库存，避免修改模板
        available_vehicles = copy.deepcopy(self.vehicle_templates)

        while unvisited:
            # 1. 选择一辆车 (简单策略：按顺序找有库存的)
            chosen_vehicle = None
            for v in available_vehicles:
                if v['count'] > 0:
                    chosen_vehicle = v
                    break

            if chosen_vehicle is None:
                # 车辆用光了，无法服务剩余客户 -> 解无效
                return [], float('inf')

            # 扣库存
            chosen_vehicle['count'] -= 1

            # 2. 路径构建
            curr_node = 0
            curr_load = 0
            # 车必须在仓库开门后出发
            curr_time = self.depot_tw['start']

            route_path = []
            route_dist = 0
            route_penalty = 0

            while True:
                candidates = []
                probs = []

                for next_node in unvisited:
                    demand = self.customers[next_node - 1][2]

                    # (A) 容量检查
                    if curr_load + demand > chosen_vehicle['capacity']:
                        continue

                    # (B) 时间窗检查
                    dist = self.dist_mat[curr_node][next_node]
                    travel_time = dist / chosen_vehicle['velocity']
                    arrival = curr_time + travel_time

                    # 客户时间窗
                    oid = self.orig_ids[next_node - 1]
                    tw = self.tws[oid]

                    # 硬约束：如果到达时间太晚，连客户门都关了 (严格模式)
                    # 或者我们可以允许迟到(软)，但在本阶段建议稍微严格一点，或者完全依赖惩罚
                    # 这里采用：必须在【仓库关门前】能回到仓库

                    start_service = max(arrival, tw['start'])
                    finish_service = start_service + tw['service']

                    # 回程检查
                    dist_back = self.dist_mat[next_node][0]
                    time_back = dist_back / chosen_vehicle['velocity']

                    if finish_service + time_back > self.depot_tw['end']:
                        continue  # 来不及回去了

                    candidates.append(next_node)

                    tau = self.pheromone[curr_node][next_node] ** config.ACO_ALPHA
                    eta = self.heuristic[curr_node][next_node] ** config.ACO_BETA
                    probs.append(tau * eta)

                if not candidates:
                    break  # 无法继续，回程

                # 轮盘赌
                probs = np.array(probs)
                if probs.sum() == 0:  # 极其罕见情况
                    break
                probs = probs / probs.sum()

                next_node = np.random.choice(candidates, p=probs)

                # 更新状态
                route_path.append(next_node - 1)  # 存局部索引
                unvisited.remove(next_node)
                curr_load += self.customers[next_node - 1][2]

                dist = self.dist_mat[curr_node][next_node]
                route_dist += dist

                arrival = curr_time + dist / chosen_vehicle['velocity']
                oid = self.orig_ids[next_node - 1]
                tw = self.tws[oid]

                # 软时间窗惩罚
                if arrival > tw['end']:
                    penalty = (arrival - tw['end']) * config.TIME_WINDOW_PENALTY
                    route_penalty += penalty

                curr_time = max(arrival, tw['start']) + tw['service']
                curr_node = next_node

            # 回仓库
            if len(route_path) > 0:
                dist_back = self.dist_mat[curr_node][0]
                route_dist += dist_back

                routes.append({
                    'path': route_path,
                    'vehicle': chosen_vehicle,
                    'distance': route_dist
                })

                # 成本 = 距离成本 + 软时间窗惩罚 + 固定启动成本(可选)
                # 假设每公里成本 = vehicle['cost']
                cost = route_dist * chosen_vehicle['cost'] + route_penalty
                total_cost += cost
            else:
                # 选了车但没跑任何点（可能是所有点都太远或时间不匹配），这辆车浪费了，回退？
                # 简单起见，不回退，算作一次空车尝试，继续循环
                pass

        return routes, total_cost

    def update_pheromone(self, ant_routes, ant_costs):
        self.pheromone *= (1 - config.ACO_RHO)

        for i in range(len(ant_routes)):
            routes = ant_routes[i]
            cost = ant_costs[i]

            if cost == float('inf') or cost == 0:
                continue

            delta = config.ACO_Q / cost

            for r_info in routes:
                # 还原路径: 0 -> node -> ... -> 0
                path_indices = r_info['path']
                curr = 0
                for node_idx in path_indices:
                    next_node = node_idx + 1  # 局部索引+1 = 矩阵索引
                    self.pheromone[curr][next_node] += delta
                    curr = next_node
                self.pheromone[curr][0] += delta