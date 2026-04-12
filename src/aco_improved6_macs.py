import copy
import numpy as np

from . import config
from .aco_vrptw import ACO_VRPTW


class ACO_Improved6_MACS(ACO_VRPTW):
    def __init__(self, depot_coord, customers_data, vehicle_list, time_windows, depot_tw, original_ids):
        super().__init__(depot_coord, customers_data, vehicle_list, time_windows, depot_tw, original_ids)
        self.pheromone_cost = np.ones_like(self.pheromone)
        self.pheromone_vehicle = np.ones_like(self.pheromone)
        self.tau_min = 0.05
        self.tau_max = 6.0
        # improve6.1 tuning: keep MACS structure, but soften cross-colony mixing
        # and reduce vehicle-count dominance to avoid sacrificing distance.
        self.exchange_ratio = 0.06
        self.vehicle_count_penalty = 10.0
        self.cost_elite_boost = 2.6
        self.vehicle_elite_boost = 1.9

    def run(self, iterations=500, ant_count=70, seed=42):
        rng = np.random.default_rng(seed)
        best_routes = None
        best_cost = float("inf")
        convergence = []

        for it in range(iterations):
            cost_colony_solutions = []
            cost_colony_scores = []
            veh_colony_solutions = []
            veh_colony_scores = []

            alpha_cost = 1.0 + 0.6 * (it / max(iterations - 1, 1))
            beta_cost = 2.6 - 0.8 * (it / max(iterations - 1, 1))

            alpha_veh = 0.8
            beta_veh = 3.0

            for _ in range(ant_count):
                routes_c, pure_cost = self._construct_solution(
                    pheromone=self.pheromone_cost,
                    alpha=alpha_cost,
                    beta=beta_cost,
                    mutation_prob=0.08,
                    rng=rng,
                )
                cost_colony_solutions.append(routes_c)
                cost_colony_scores.append(pure_cost)

                if pure_cost < best_cost:
                    best_cost = pure_cost
                    best_routes = routes_c

                routes_v, pure_cost_v = self._construct_solution(
                    pheromone=self.pheromone_vehicle,
                    alpha=alpha_veh,
                    beta=beta_veh,
                    mutation_prob=0.18,
                    rng=rng,
                )
                veh_colony_solutions.append(routes_v)
                if routes_v is None or pure_cost_v == float("inf"):
                    veh_colony_scores.append(float("inf"))
                else:
                    veh_obj = pure_cost_v + self.vehicle_count_penalty * len(routes_v)
                    veh_colony_scores.append(veh_obj)
                    if pure_cost_v < best_cost:
                        best_cost = pure_cost_v
                        best_routes = routes_v

            rho = 0.22 - 0.12 * (it / max(iterations - 1, 1))
            self._update_pheromone(
                self.pheromone_cost, cost_colony_solutions, cost_colony_scores, rho, self.cost_elite_boost
            )
            self._update_pheromone(
                self.pheromone_vehicle, veh_colony_solutions, veh_colony_scores, rho, self.vehicle_elite_boost
            )
            self._exchange_pheromone()

            convergence.append(best_cost)

        return best_routes, best_cost, convergence

    def _construct_solution(self, pheromone, alpha, beta, mutation_prob, rng):
        unvisited = set(range(1, self.num_nodes))
        routes = []
        total_cost = 0.0
        available_vehicles = copy.deepcopy(self.vehicle_templates)

        while unvisited:
            chosen_vehicle = None
            for v in available_vehicles:
                if v["count"] > 0:
                    chosen_vehicle = v
                    break

            if chosen_vehicle is None:
                return None, float("inf")

            chosen_vehicle["count"] -= 1
            curr_node = 0
            curr_load = 0.0
            curr_time = self.depot_tw["start"]
            route_path = []
            route_dist = 0.0
            route_penalty = 0.0

            while True:
                valid_nodes = []
                scores = []

                for next_node in unvisited:
                    demand = self.customers[next_node - 1][2]
                    if curr_load + demand > chosen_vehicle["capacity"]:
                        continue

                    dist = self.dist_mat[curr_node][next_node]
                    arrival = curr_time + dist / chosen_vehicle["velocity"]
                    oid = self.orig_ids[next_node - 1]
                    tw = self.tws[oid]
                    service_start = max(arrival, tw["start"])
                    finish_service = service_start + tw["service"]
                    dist_back = self.dist_mat[next_node][0]
                    time_back = dist_back / chosen_vehicle["velocity"]
                    if finish_service + time_back > self.depot_tw["end"]:
                        continue

                    slack = max(tw["end"] - service_start, 0.1)
                    urgency = 1.0 / slack
                    cap_fit = demand / max(chosen_vehicle["capacity"] - curr_load, 1e-6)
                    eta = (1.0 / (dist + 1e-6)) * (1.0 + 0.55 * urgency) * (1.0 + 0.20 * cap_fit)
                    tau = pheromone[curr_node][next_node]
                    score = (tau ** alpha) * (eta ** beta)

                    valid_nodes.append(next_node)
                    scores.append(score)

                if not valid_nodes:
                    break

                if rng.random() < mutation_prob:
                    next_node = int(valid_nodes[int(rng.integers(0, len(valid_nodes)))])
                else:
                    probs = np.array(scores, dtype=float)
                    if probs.sum() <= 0:
                        probs = np.ones(len(valid_nodes), dtype=float) / len(valid_nodes)
                    else:
                        probs = probs / probs.sum()
                    next_node = int(rng.choice(valid_nodes, p=probs))

                route_path.append(next_node - 1)
                unvisited.remove(next_node)
                curr_load += self.customers[next_node - 1][2]

                dist = self.dist_mat[curr_node][next_node]
                route_dist += dist
                arrival = curr_time + dist / chosen_vehicle["velocity"]
                oid = self.orig_ids[next_node - 1]
                tw = self.tws[oid]
                if arrival > tw["end"]:
                    route_penalty += (arrival - tw["end"]) * config.TIME_WINDOW_PENALTY
                curr_time = max(arrival, tw["start"]) + tw["service"]
                curr_node = next_node

            if route_path:
                dist_back = self.dist_mat[curr_node][0]
                route_dist += dist_back
                routes.append({"path": route_path, "vehicle": chosen_vehicle, "distance": route_dist})
                total_cost += route_dist * chosen_vehicle["cost"] + route_penalty

        return routes, total_cost

    def _update_pheromone(self, pheromone, solutions, scores, rho, elite_boost):
        pheromone *= (1 - rho)
        ranked = [
            (sol, score)
            for sol, score in zip(solutions, scores)
            if sol is not None and score not in (0, float("inf"))
        ]
        ranked.sort(key=lambda x: x[1])

        for rank, (routes, score) in enumerate(ranked[:6]):
            delta = config.ACO_Q / score
            if rank == 0:
                delta *= elite_boost
            for r_info in routes:
                curr = 0
                for node_idx in r_info["path"]:
                    nxt = node_idx + 1
                    pheromone[curr][nxt] += delta
                    curr = nxt
                pheromone[curr][0] += delta

        np.clip(pheromone, self.tau_min, self.tau_max, out=pheromone)

    def _exchange_pheromone(self):
        mix_a = (1 - self.exchange_ratio) * self.pheromone_cost + self.exchange_ratio * self.pheromone_vehicle
        mix_b = (1 - self.exchange_ratio) * self.pheromone_vehicle + self.exchange_ratio * self.pheromone_cost
        self.pheromone_cost = np.clip(mix_a, self.tau_min, self.tau_max)
        self.pheromone_vehicle = np.clip(mix_b, self.tau_min, self.tau_max)
