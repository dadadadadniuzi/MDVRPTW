# src/aco_improved.py
import numpy as np
import copy
import os
from .utils import euclidean_distance
from . import config
from .aco_vrptw import ACO_VRPTW


class ACO_Improved(ACO_VRPTW):
    def run(self):
        best_routes = None
        best_cost = float('inf')
        convergence = []

        # 定义各等级蚂蚁数量
        n_total = config.ACO_ANT_COUNT
        n_ball = int(n_total * 0.2)  # 滚球 (精英)
        n_brood = int(n_total * 0.4)  # 育雏 (标准)
        n_small = int(n_total * 0.2)  # 小蜣螂 (贪婪)
        # 剩下的都是偷窃 (随机)

        for iter in range(config.ACO_ITERATIONS):
            ant_routes = []
            ant_costs = []

            for k in range(n_total):
                # 1. 确定蚂蚁身份和参数
                role = 'brood'  # 默认为育雏
                alpha = config.ACO_ALPHA
                beta = config.ACO_BETA
                mutation_prob = 0.0  # 随机变异概率

                if k < n_ball:
                    role = 'ball'
                    alpha = 2.0  # 对信息素极度敏感 (跟随强者)
                    beta = 1.0
                elif k < n_ball + n_brood:
                    role = 'brood'
                    # 标准参数
                elif k < n_ball + n_brood + n_small:
                    role = 'small'
                    alpha = 0.1  # 几乎忽略信息素
                    beta = 4.0  # 极度依赖距离 (贪婪搜索)
                else:
                    role = 'thief'
                    mutation_prob = 0.2  # 20% 概率随机乱走

                # 2. 构建解
                routes, cost = self.construct_solution_improved(alpha, beta, mutation_prob)

                ant_routes.append(routes)
                ant_costs.append(cost)

                if cost < best_cost:
                    best_cost = cost
                    best_routes = routes
                    convergence.append(best_cost)

            # 3. 更新信息素 (差异化更新)
            # 滚球蚂蚁的路径权重更高
            self.update_pheromone_improved(ant_routes, ant_costs, n_ball)

            # 日志 (可选)
            # if (iter + 1) % 10 == 0:
            #     print(f"    [DBO-ACO] Iter {iter+1}: Best Cost {best_cost:.2f}")

        return best_routes, best_cost, convergence

    def construct_solution_improved(self, alpha, beta, mutation_prob):
        """
        支持自定义 alpha, beta 和 变异概率 的路径构建
        """
        unvisited = set(range(1, self.num_nodes))
        routes = []
        total_cost = 0

        # 复制车辆库存
        available_vehicles = copy.deepcopy(self.vehicle_templates)

        while unvisited:
            # 选车
            chosen_vehicle = None
            for v in available_vehicles:
                if v['count'] > 0:
                    chosen_vehicle = v
                    break

            if chosen_vehicle is None:
                return [], float('inf')

            chosen_vehicle['count'] -= 1

            # 路径初始化
            curr_node = 0
            curr_load = 0
            curr_time = self.depot_tw['start']

            route_path = []
            route_dist = 0
            route_penalty = 0

            while True:
                candidates = []
                probs = []

                # 筛选可行节点
                valid_nodes = []
                for next_node in unvisited:
                    demand = self.customers[next_node - 1][2]

                    # 容量检查
                    if curr_load + demand > chosen_vehicle['capacity']:
                        continue

                    # 时间窗检查
                    dist = self.dist_mat[curr_node][next_node]
                    arrival = curr_time + dist / chosen_vehicle['velocity']
                    oid = self.orig_ids[next_node - 1]
                    tw = self.tws[oid]

                    start_service = max(arrival, tw['start'])
                    finish_service = start_service + tw['service']

                    dist_back = self.dist_mat[next_node][0]
                    time_back = dist_back / chosen_vehicle['velocity']

                    if finish_service + time_back > self.depot_tw['end']:
                        continue

                    valid_nodes.append(next_node)

                if not valid_nodes:
                    break  # 无法继续

                # --- 核心改进：选择逻辑 ---
                # 如果是“偷窃蚂蚁”且触发了变异，直接随机选一个
                if np.random.random() < mutation_prob:
                    next_node = np.random.choice(valid_nodes)
                else:
                    # 标准轮盘赌，但使用自定义的 alpha/beta
                    for next_node in valid_nodes:
                        tau = self.pheromone[curr_node][next_node] ** alpha
                        eta = self.heuristic[curr_node][next_node] ** beta
                        probs.append(tau * eta)

                    probs = np.array(probs)
                    if probs.sum() == 0:
                        probs = np.ones(len(probs)) / len(probs)
                    else:
                        probs = probs / probs.sum()

                    next_node = np.random.choice(valid_nodes, p=probs)

                # 更新状态
                route_path.append(next_node - 1)
                unvisited.remove(next_node)
                curr_load += self.customers[next_node - 1][2]

                dist = self.dist_mat[curr_node][next_node]
                route_dist += dist

                arrival = curr_time + dist / chosen_vehicle['velocity']
                oid = self.orig_ids[next_node - 1]
                tw = self.tws[oid]

                if arrival > tw['end']:
                    route_penalty += (arrival - tw['end']) * config.TIME_WINDOW_PENALTY

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
                total_cost += route_dist * chosen_vehicle['cost'] + route_penalty

        return routes, total_cost

    def update_pheromone_improved(self, ant_routes, ant_costs, n_ball):
        # 1. 全局挥发
        self.pheromone *= (1 - config.ACO_RHO)

        # 2. 增强 (给予精英蚂蚁更高权重)
        for i in range(len(ant_routes)):
            routes = ant_routes[i]
            cost = ant_costs[i]

            if cost == float('inf') or cost == 0: continue

            # 基础增量
            delta = config.ACO_Q / cost

            # 如果是前 n_ball 只蚂蚁 (精英)，增量加倍
            if i < n_ball:
                delta *= 2.0

            for r_info in routes:
                path_indices = r_info['path']
                curr = 0
                for node_idx in path_indices:
                    next_node = node_idx + 1
                    self.pheromone[curr][next_node] += delta
                    curr = next_node
                self.pheromone[curr][0] += delta