import copy
import numpy as np

from . import config
from .aco_vrptw import ACO_VRPTW


class ACO_Improved6_2_DBO(ACO_VRPTW):
    def __init__(self, depot_coord, customers_data, vehicle_list, time_windows, depot_tw, original_ids, params=None):
        super().__init__(depot_coord, customers_data, vehicle_list, time_windows, depot_tw, original_ids)
        defaults = {
            "exchange_ratio": 0.06,
            "vehicle_count_penalty": 9.0,
            "cost_elite_boost": 2.7,
            "vehicle_elite_boost": 1.8,
            "rho_min": 0.08,
            "rho_max": 0.24,
            "tau_min": 0.05,
            "tau_max": 6.5,
            "mut_cost": 0.06,
            "mut_vehicle": 0.16,
            "n_ball_ratio": 0.20,
            "n_brood_ratio": 0.40,
            "n_small_ratio": 0.20,
            "adapt_interval": 40,
        }
        self.params = defaults if params is None else {**defaults, **params}
        self.pheromone_cost = np.ones_like(self.pheromone)
        self.pheromone_vehicle = np.ones_like(self.pheromone)

        # Inner DBO control vector: [alpha_cost, beta_cost, alpha_veh, beta_veh, exchange_ratio]
        self.ctrl_lb = np.array([1.0, 1.4, 0.4, 2.2, 0.02], dtype=float)
        self.ctrl_ub = np.array([2.4, 3.4, 1.5, 4.3, 0.18], dtype=float)
        self.ctrl_pop = None
        self.ctrl_fit = None
        self.ctrl_best = np.array([1.6, 2.2, 0.8, 3.1, self.params["exchange_ratio"]], dtype=float)
        self.ctrl_best_fit = float("inf")

    def run(self, iterations=550, ant_count=80, seed=42):
        rng = np.random.default_rng(seed)
        best_routes = None
        best_cost = float("inf")
        convergence = []

        self._init_inner_dbo(rng)
        n_total = ant_count
        n_ball = max(1, int(n_total * self.params["n_ball_ratio"]))
        n_brood = max(1, int(n_total * self.params["n_brood_ratio"]))
        n_small = max(1, int(n_total * self.params["n_small_ratio"]))

        for it in range(iterations):
            if it % int(self.params["adapt_interval"]) == 0:
                self._inner_dbo_step(rng)

            alpha_cost, beta_cost, alpha_veh, beta_veh, ex_ratio = self.ctrl_best
            ant_cost_routes, ant_cost_scores = [], []
            ant_veh_routes, ant_veh_scores = [], []

            for _ in range(n_total):
                # Cost colony
                routes_c, pure_cost = self._construct_solution(
                    pheromone=self.pheromone_cost,
                    alpha=alpha_cost,
                    beta=beta_cost,
                    mutation_prob=self.params["mut_cost"],
                    rng=rng,
                )
                ant_cost_routes.append(routes_c)
                ant_cost_scores.append(pure_cost)
                if pure_cost < best_cost:
                    best_cost = pure_cost
                    best_routes = routes_c

                # Vehicle colony
                routes_v, pure_cost_v = self._construct_solution(
                    pheromone=self.pheromone_vehicle,
                    alpha=alpha_veh,
                    beta=beta_veh,
                    mutation_prob=self.params["mut_vehicle"],
                    rng=rng,
                )
                ant_veh_routes.append(routes_v)
                if routes_v is None or pure_cost_v == float("inf"):
                    ant_veh_scores.append(float("inf"))
                else:
                    veh_obj = pure_cost_v + self.params["vehicle_count_penalty"] * len(routes_v)
                    ant_veh_scores.append(veh_obj)
                    if pure_cost_v < best_cost:
                        best_cost = pure_cost_v
                        best_routes = routes_v

            progress = it / max(iterations - 1, 1)
            rho = self.params["rho_max"] - (self.params["rho_max"] - self.params["rho_min"]) * progress

            self._update_pheromone(
                self.pheromone_cost,
                ant_cost_routes,
                ant_cost_scores,
                rho,
                self.params["cost_elite_boost"],
                n_ball,
            )
            self._update_pheromone(
                self.pheromone_vehicle,
                ant_veh_routes,
                ant_veh_scores,
                rho,
                self.params["vehicle_elite_boost"],
                n_ball,
            )
            self._exchange_pheromone(float(ex_ratio))
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
                valid_nodes, scores = [], []
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
                    eta = (1.0 / (dist + 1e-6)) * (1.0 + 0.55 * urgency) * (1.0 + 0.22 * cap_fit)
                    tau = pheromone[curr_node][next_node]
                    scores.append((tau ** alpha) * (eta ** beta))
                    valid_nodes.append(next_node)

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

    def _update_pheromone(self, pheromone, solutions, scores, rho, elite_boost, n_ball):
        pheromone *= (1 - rho)
        ranked = [
            (idx, sol, score)
            for idx, (sol, score) in enumerate(zip(solutions, scores))
            if sol is not None and score not in (0, float("inf"))
        ]
        ranked.sort(key=lambda x: x[2])

        for rank, (idx, routes, score) in enumerate(ranked[:6]):
            delta = config.ACO_Q / score
            if idx < n_ball:
                delta *= 2.0
            if rank == 0:
                delta *= elite_boost
            for r_info in routes:
                curr = 0
                for node_idx in r_info["path"]:
                    nxt = node_idx + 1
                    pheromone[curr][nxt] += delta
                    curr = nxt
                pheromone[curr][0] += delta

        np.clip(pheromone, self.params["tau_min"], self.params["tau_max"], out=pheromone)

    def _exchange_pheromone(self, ratio):
        mix_a = (1 - ratio) * self.pheromone_cost + ratio * self.pheromone_vehicle
        mix_b = (1 - ratio) * self.pheromone_vehicle + ratio * self.pheromone_cost
        self.pheromone_cost = np.clip(mix_a, self.params["tau_min"], self.params["tau_max"])
        self.pheromone_vehicle = np.clip(mix_b, self.params["tau_min"], self.params["tau_max"])

    def _init_inner_dbo(self, rng):
        pop_size = 6
        self.ctrl_pop = rng.uniform(self.ctrl_lb, self.ctrl_ub, size=(pop_size, len(self.ctrl_lb)))
        self.ctrl_fit = np.full(pop_size, float("inf"))

    def _inner_dbo_step(self, rng):
        # Evaluate control candidates with a small probing budget.
        for i in range(len(self.ctrl_pop)):
            score = self._probe_control(self.ctrl_pop[i], rng)
            self.ctrl_fit[i] = score
            if score < self.ctrl_best_fit:
                self.ctrl_best_fit = score
                self.ctrl_best = self.ctrl_pop[i].copy()

        order = np.argsort(self.ctrl_fit)
        self.ctrl_pop = self.ctrl_pop[order]
        self.ctrl_fit = self.ctrl_fit[order]
        worst = self.ctrl_pop[-1].copy()

        n_ball = max(1, int(len(self.ctrl_pop) * 0.3))
        n_brood = max(1, int(len(self.ctrl_pop) * 0.3))

        for i in range(len(self.ctrl_pop)):
            cand = self.ctrl_pop[i].copy()
            if i < n_ball:
                cand = cand + 0.32 * rng.random(len(cand)) * (self.ctrl_best - cand) - 0.08 * rng.random(len(cand)) * (worst - cand)
            elif i < n_ball + n_brood:
                cand = self.ctrl_best + rng.normal(0, 0.12, size=len(cand))
            else:
                peer = self.ctrl_pop[int(rng.integers(0, len(self.ctrl_pop)))]
                cand = cand + rng.normal(0, 0.08, size=len(cand)) + 0.10 * (peer - cand)
            cand = np.clip(cand, self.ctrl_lb, self.ctrl_ub)

            cand_score = self._probe_control(cand, rng)
            if cand_score < self.ctrl_fit[i]:
                self.ctrl_pop[i] = cand
                self.ctrl_fit[i] = cand_score
                if cand_score < self.ctrl_best_fit:
                    self.ctrl_best_fit = cand_score
                    self.ctrl_best = cand.copy()

    def _probe_control(self, ctrl_vec, rng):
        alpha_cost, beta_cost, alpha_veh, beta_veh, _ = ctrl_vec.tolist()
        # Probe with very small budget: 2 + 2 ants, single-step score aggregation.
        score_sum = 0.0
        for _ in range(2):
            _, c = self._construct_solution(self.pheromone_cost, alpha_cost, beta_cost, self.params["mut_cost"], rng)
            score_sum += c if c != float("inf") else 1e9
        for _ in range(2):
            routes, c = self._construct_solution(self.pheromone_vehicle, alpha_veh, beta_veh, self.params["mut_vehicle"], rng)
            if c == float("inf") or routes is None:
                score_sum += 1e9
            else:
                score_sum += c + self.params["vehicle_count_penalty"] * len(routes)
        return score_sum / 4.0


def tune_improve6_2_params_with_dbo(problem_args, seed=42):
    bounds = [
        ("exchange_ratio", 0.03, 0.12),
        ("vehicle_count_penalty", 6.0, 14.0),
        ("cost_elite_boost", 2.0, 3.2),
        ("vehicle_elite_boost", 1.4, 2.2),
        ("rho_min", 0.06, 0.12),
        ("rho_max", 0.18, 0.28),
        ("mut_cost", 0.03, 0.10),
        ("mut_vehicle", 0.12, 0.24),
    ]
    pop_size = 8
    max_iter = 5
    rng = np.random.default_rng(seed)
    dim = len(bounds)
    lb = np.array([b[1] for b in bounds], dtype=float)
    ub = np.array([b[2] for b in bounds], dtype=float)
    pop = rng.uniform(lb, ub, size=(pop_size, dim))
    fit = np.full(pop_size, float("inf"))
    cache = {}

    base = {
        "tau_min": 0.05,
        "tau_max": 6.5,
        "n_ball_ratio": 0.20,
        "n_brood_ratio": 0.40,
        "n_small_ratio": 0.20,
        "adapt_interval": 40,
    }

    def decode(vec):
        p = dict(base)
        for i, (name, _, _) in enumerate(bounds):
            p[name] = float(vec[i])
        if p["rho_min"] >= p["rho_max"]:
            p["rho_min"] = max(0.05, p["rho_max"] - 0.08)
        return p

    def objective(vec, run_seed):
        key = tuple(np.round(vec, 3))
        if key in cache:
            return cache[key]
        params = decode(vec)
        solver = ACO_Improved6_2_DBO(*problem_args, params=params)
        _, cost, _ = solver.run(iterations=120, ant_count=36, seed=run_seed)
        cache[key] = float(cost)
        return float(cost)

    for i in range(pop_size):
        fit[i] = objective(pop[i], seed + i * 31)

    best_idx = int(np.argmin(fit))
    gbest = pop[best_idx].copy()
    gbest_fit = float(fit[best_idx])
    n_ball = max(1, int(pop_size * 0.25))
    n_brood = max(1, int(pop_size * 0.35))

    for it in range(max_iter):
        order = np.argsort(fit)
        pop = pop[order]
        fit = fit[order]
        worst = pop[-1].copy()

        for i in range(pop_size):
            cand = pop[i].copy()
            if i < n_ball:
                cand = cand + 0.30 * rng.random(dim) * (gbest - cand) - 0.10 * rng.random(dim) * (worst - cand)
            elif i < n_ball + n_brood:
                sigma = 0.18 * (1 - it / max(max_iter - 1, 1)) + 0.03
                cand = gbest + rng.normal(0, sigma, size=dim)
            else:
                peer = pop[int(rng.integers(0, pop_size))]
                cand = cand + rng.normal(0, 0.09, size=dim) + 0.10 * (peer - cand)

            cand = np.clip(cand, lb, ub)
            cand_fit = objective(cand, seed + 900 + it * pop_size + i)
            if cand_fit < fit[i]:
                pop[i] = cand
                fit[i] = cand_fit
                if cand_fit < gbest_fit:
                    gbest = cand.copy()
                    gbest_fit = cand_fit

    return decode(gbest), gbest_fit
