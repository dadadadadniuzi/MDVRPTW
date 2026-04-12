import copy
import numpy as np

from .aco_vrptw import ACO_VRPTW


class ACO_Improved5(ACO_VRPTW):
    def __init__(self, depot_coord, customers_data, vehicle_list, time_windows, depot_tw, original_ids, params=None):
        super().__init__(depot_coord, customers_data, vehicle_list, time_windows, depot_tw, original_ids)
        defaults = {
            "alpha": 1.0,
            "beta": 2.0,
            "rho": 0.10,
            "q": 100.0,
            "mutation_prob": 0.20,
            "time_penalty": 100.0,
            "elite_boost": 2.0,
            "tau_min": 0.05,
            "tau_max": 6.0,
            "n_ball_ratio": 0.20,
            "n_brood_ratio": 0.40,
            "n_small_ratio": 0.20,
        }
        self.params = defaults if params is None else {**defaults, **params}

    def run(self, iterations=400, ant_count=60, seed=42):
        rng = np.random.default_rng(seed)
        best_routes = None
        best_cost = float("inf")
        convergence = []

        n_total = ant_count
        n_ball = max(1, int(n_total * self.params["n_ball_ratio"]))
        n_brood = max(1, int(n_total * self.params["n_brood_ratio"]))
        n_small = max(1, int(n_total * self.params["n_small_ratio"]))

        for _ in range(iterations):
            ant_routes = []
            ant_costs = []

            for k in range(n_total):
                alpha = self.params["alpha"]
                beta = self.params["beta"]
                mutation_prob = 0.0

                if k < n_ball:
                    alpha = self.params["alpha"] + 0.8
                    beta = max(1.0, self.params["beta"] - 0.8)
                elif k < n_ball + n_brood:
                    alpha = self.params["alpha"]
                    beta = self.params["beta"]
                elif k < n_ball + n_brood + n_small:
                    alpha = 0.1
                    beta = self.params["beta"] + 1.8
                else:
                    mutation_prob = self.params["mutation_prob"]

                routes, cost = self.construct_solution(alpha, beta, mutation_prob, rng)
                ant_routes.append(routes)
                ant_costs.append(cost)

                if cost < best_cost:
                    best_cost = cost
                    best_routes = routes

            self.update_pheromone(ant_routes, ant_costs, n_ball, best_routes, best_cost)
            convergence.append(best_cost)

        return best_routes, best_cost, convergence

    def construct_solution(self, alpha, beta, mutation_prob, rng):
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
                probs = []

                for next_node in unvisited:
                    demand = self.customers[next_node - 1][2]
                    if curr_load + demand > chosen_vehicle["capacity"]:
                        continue

                    dist = self.dist_mat[curr_node][next_node]
                    arrival = curr_time + dist / chosen_vehicle["velocity"]
                    oid = self.orig_ids[next_node - 1]
                    tw = self.tws[oid]
                    start_service = max(arrival, tw["start"])
                    finish_service = start_service + tw["service"]

                    dist_back = self.dist_mat[next_node][0]
                    time_back = dist_back / chosen_vehicle["velocity"]
                    if finish_service + time_back > self.depot_tw["end"]:
                        continue

                    slack = max(tw["end"] - start_service, 0.1)
                    urgency = 1.0 / slack
                    capacity_fit = demand / max(chosen_vehicle["capacity"] - curr_load, 1e-6)
                    eta = (1.0 / (dist + 1e-6)) * (1.0 + 0.5 * urgency) * (1.0 + 0.25 * capacity_fit)
                    tau = self.pheromone[curr_node][next_node]
                    prob = (tau ** alpha) * (eta ** beta)

                    valid_nodes.append(next_node)
                    probs.append(prob)

                if not valid_nodes:
                    break

                if rng.random() < mutation_prob:
                    next_node = int(valid_nodes[int(rng.integers(0, len(valid_nodes)))])
                else:
                    probs = np.array(probs, dtype=float)
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
                    route_penalty += (arrival - tw["end"]) * self.params["time_penalty"]

                curr_time = max(arrival, tw["start"]) + tw["service"]
                curr_node = next_node

            if route_path:
                dist_back = self.dist_mat[curr_node][0]
                route_dist += dist_back
                routes.append({"path": route_path, "vehicle": chosen_vehicle, "distance": route_dist})
                total_cost += route_dist * chosen_vehicle["cost"] + route_penalty

        return routes, total_cost

    def update_pheromone(self, ant_routes, ant_costs, n_ball, best_routes, best_cost):
        self.pheromone *= (1 - self.params["rho"])

        ranked = [
            (idx, routes, cost)
            for idx, (routes, cost) in enumerate(zip(ant_routes, ant_costs))
            if routes and cost not in (0, float("inf"))
        ]
        ranked.sort(key=lambda item: item[2])

        for rank, (idx, routes, cost) in enumerate(ranked[:5]):
            delta = self.params["q"] / cost
            if idx < n_ball:
                delta *= 2.0
            if rank == 0:
                delta *= self.params["elite_boost"]
            for r_info in routes:
                curr = 0
                for node_idx in r_info["path"]:
                    next_node = node_idx + 1
                    self.pheromone[curr][next_node] += delta
                    curr = next_node
                self.pheromone[curr][0] += delta

        if best_routes and best_cost not in (0, float("inf")):
            elite_delta = (self.params["q"] / best_cost) * 0.8
            for r_info in best_routes:
                curr = 0
                for node_idx in r_info["path"]:
                    next_node = node_idx + 1
                    self.pheromone[curr][next_node] += elite_delta
                    curr = next_node
                self.pheromone[curr][0] += elite_delta

        self.pheromone = np.clip(self.pheromone, self.params["tau_min"], self.params["tau_max"])


def tune_params_by_dbo(problem_args, seed=42):
    bounds = [
        ("alpha", 0.8, 1.6),
        ("beta", 1.6, 3.0),
        ("rho", 0.06, 0.16),
        ("mutation_prob", 0.10, 0.28),
        ("time_penalty", 70.0, 140.0),
        ("elite_boost", 1.4, 2.4),
    ]

    pop_size = 8
    max_iter = 5
    dim = len(bounds)
    rng = np.random.default_rng(seed)
    lb = np.array([b[1] for b in bounds], dtype=float)
    ub = np.array([b[2] for b in bounds], dtype=float)
    pop = rng.uniform(lb, ub, size=(pop_size, dim))
    fit = np.full(pop_size, float("inf"))
    cache = {}

    base = {
        "q": 100.0,
        "tau_min": 0.05,
        "tau_max": 6.0,
        "n_ball_ratio": 0.20,
        "n_brood_ratio": 0.40,
        "n_small_ratio": 0.20,
    }

    def decode(vec):
        p = dict(base)
        for i, (name, _, _) in enumerate(bounds):
            p[name] = float(vec[i])
        return p

    def objective(vec, run_seed):
        key = tuple(np.round(vec, 3))
        if key in cache:
            return cache[key]

        params = decode(vec)
        # Two short runs to reduce noise mismatch between tuning and final run.
        costs = []
        for offset in (0, 1):
            solver = ACO_Improved5(*problem_args, params=params)
            _, cost, _ = solver.run(iterations=70, ant_count=22, seed=run_seed + offset)
            costs.append(cost)
        value = float(np.mean(costs))
        cache[key] = value
        return value

    for i in range(pop_size):
        fit[i] = objective(pop[i], seed + i * 13)

    gbest_idx = int(np.argmin(fit))
    gbest = pop[gbest_idx].copy()
    gbest_fit = float(fit[gbest_idx])

    n_ball = max(1, int(pop_size * 0.2))
    n_brood = max(1, int(pop_size * 0.4))
    n_small = max(1, int(pop_size * 0.2))

    for it in range(max_iter):
        order = np.argsort(fit)
        pop = pop[order]
        fit = fit[order]
        worst = pop[-1].copy()

        for i in range(pop_size):
            cand = pop[i].copy()

            if i < n_ball:
                cand = cand + 0.3 * rng.random(dim) * (gbest - cand) - 0.08 * rng.random(dim) * (worst - cand)
            elif i < n_ball + n_brood:
                sigma = 0.18 * (1 - it / max(max_iter - 1, 1)) + 0.02
                cand = gbest + rng.normal(0, sigma, size=dim)
            elif i < n_ball + n_brood + n_small:
                cand = cand + 0.18 * rng.random(dim) * (gbest - cand) + 0.05 * (rng.random(dim) - cand)
            else:
                peer = pop[int(rng.integers(0, pop_size))]
                cand = cand + rng.normal(0, 0.10, size=dim) + 0.12 * (peer - cand)

            cand = np.clip(cand, lb, ub)
            cand_fit = objective(cand, seed + 1000 + it * 37 + i)
            if cand_fit < fit[i]:
                pop[i] = cand
                fit[i] = cand_fit
                if cand_fit < gbest_fit:
                    gbest = cand.copy()
                    gbest_fit = float(cand_fit)

    return decode(gbest), gbest_fit
