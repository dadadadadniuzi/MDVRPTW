import copy
import numpy as np

from . import config
from .aco_vrptw import ACO_VRPTW


class ACO_Improved2(ACO_VRPTW):
    def __init__(self, depot_coord, customers_data, vehicle_list, time_windows, depot_tw, original_ids):
        super().__init__(depot_coord, customers_data, vehicle_list, time_windows, depot_tw, original_ids)
        self.tau_max = 5.0
        self.tau_min = 0.05

    def run(self):
        best_routes = None
        best_cost = float("inf")
        convergence = []

        for iteration in range(config.ACO_ITERATIONS):
            alpha, beta, rho = self._adaptive_params(iteration)
            ant_routes = []
            ant_costs = []

            for _ in range(config.ACO_ANT_COUNT):
                routes, cost = self.construct_solution_improved(alpha, beta)
                ant_routes.append(routes)
                ant_costs.append(cost)

                if cost < best_cost:
                    best_cost = cost
                    best_routes = routes

            convergence.append(best_cost)
            self.update_pheromone_improved(ant_routes, ant_costs, best_cost, rho)

        return best_routes, best_cost, convergence

    def _adaptive_params(self, iteration):
        progress = iteration / max(config.ACO_ITERATIONS - 1, 1)
        alpha = 0.8 + 1.4 * progress
        beta = 3.8 - 1.6 * progress
        rho = 0.30 - 0.20 * progress
        return alpha, beta, rho

    def construct_solution_improved(self, alpha, beta):
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
                return [], float("inf")

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
                    urgency_factor = 1.0 / slack
                    capacity_fit = demand / max(chosen_vehicle["capacity"] - curr_load, 1e-6)
                    distance_factor = 1.0 / (dist + 1e-6)

                    tau = self.pheromone[curr_node][next_node] ** alpha
                    eta = (
                        (distance_factor ** beta)
                        * (1.0 + 0.6 * urgency_factor)
                        * (1.0 + 0.3 * capacity_fit)
                    )

                    valid_nodes.append(next_node)
                    scores.append(tau * eta)

                if not valid_nodes:
                    break

                probs = np.array(scores, dtype=float)
                if probs.sum() == 0:
                    probs = np.ones(len(valid_nodes), dtype=float) / len(valid_nodes)
                else:
                    probs = probs / probs.sum()

                next_node = np.random.choice(valid_nodes, p=probs)
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
                routes.append(
                    {
                        "path": route_path,
                        "vehicle": chosen_vehicle,
                        "distance": route_dist,
                    }
                )
                total_cost += route_dist * chosen_vehicle["cost"] + route_penalty

        return routes, total_cost

    def update_pheromone_improved(self, ant_routes, ant_costs, best_cost, rho):
        self.pheromone *= (1 - rho)

        valid = [
            (routes, cost)
            for routes, cost in zip(ant_routes, ant_costs)
            if cost not in (0, float("inf")) and routes
        ]
        valid.sort(key=lambda item: item[1])

        for rank, (routes, cost) in enumerate(valid[:5]):
            delta = config.ACO_Q / cost
            if rank == 0:
                delta *= 2.0

            for r_info in routes:
                curr = 0
                for node_idx in r_info["path"]:
                    next_node = node_idx + 1
                    self.pheromone[curr][next_node] += delta
                    curr = next_node
                self.pheromone[curr][0] += delta

        if best_cost < float("inf"):
            best_delta = config.ACO_Q / best_cost
            self.pheromone += best_delta * 0.02

        self.pheromone = np.clip(self.pheromone, self.tau_min, self.tau_max)
