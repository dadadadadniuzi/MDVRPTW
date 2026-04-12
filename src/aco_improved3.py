import numpy as np

from . import config
from .aco_improved import ACO_Improved


class ACO_Improved3(ACO_Improved):
    def __init__(self, depot_coord, customers_data, vehicle_list, time_windows, depot_tw, original_ids):
        super().__init__(depot_coord, customers_data, vehicle_list, time_windows, depot_tw, original_ids)
        self.tau_min = 0.05
        self.tau_max = 6.0

    def run(self):
        best_routes = None
        best_cost = float("inf")
        convergence = []

        n_total = config.ACO_ANT_COUNT
        n_ball = int(n_total * 0.2)
        n_brood = int(n_total * 0.4)
        n_small = int(n_total * 0.2)

        for iteration in range(config.ACO_ITERATIONS):
            ant_routes = []
            ant_costs = []
            rho = self._adaptive_rho(iteration)

            for k in range(n_total):
                alpha = config.ACO_ALPHA
                beta = config.ACO_BETA
                mutation_prob = 0.0

                if k < n_ball:
                    alpha = 2.0
                    beta = 1.0
                elif k < n_ball + n_brood:
                    alpha = config.ACO_ALPHA
                    beta = config.ACO_BETA
                elif k < n_ball + n_brood + n_small:
                    alpha = 0.1
                    beta = 4.0
                else:
                    mutation_prob = 0.2

                routes, cost = self.construct_solution_improved(alpha, beta, mutation_prob)
                ant_routes.append(routes)
                ant_costs.append(cost)

                if cost < best_cost:
                    best_cost = cost
                    best_routes = routes

            convergence.append(best_cost)
            self.update_pheromone_improved3(ant_routes, ant_costs, n_ball, best_cost, best_routes, rho)

        return best_routes, best_cost, convergence

    def _adaptive_rho(self, iteration):
        progress = iteration / max(config.ACO_ITERATIONS - 1, 1)
        return 0.25 - 0.15 * progress

    def update_pheromone_improved3(self, ant_routes, ant_costs, n_ball, best_cost, best_routes, rho):
        self.pheromone *= (1 - rho)

        ranked = [
            (idx, routes, cost)
            for idx, (routes, cost) in enumerate(zip(ant_routes, ant_costs))
            if routes and cost not in (0, float("inf"))
        ]
        ranked.sort(key=lambda item: item[2])

        for rank, (idx, routes, cost) in enumerate(ranked[:5]):
            delta = config.ACO_Q / cost
            if idx < n_ball:
                delta *= 2.0
            if rank == 0:
                delta *= 1.5

            for r_info in routes:
                curr = 0
                for node_idx in r_info["path"]:
                    next_node = node_idx + 1
                    self.pheromone[curr][next_node] += delta
                    curr = next_node
                self.pheromone[curr][0] += delta

        if best_routes and best_cost not in (0, float("inf")):
            elite_delta = (config.ACO_Q / best_cost) * 1.2
            for r_info in best_routes:
                curr = 0
                for node_idx in r_info["path"]:
                    next_node = node_idx + 1
                    self.pheromone[curr][next_node] += elite_delta
                    curr = next_node
                self.pheromone[curr][0] += elite_delta

        self.pheromone = np.clip(self.pheromone, self.tau_min, self.tau_max)
