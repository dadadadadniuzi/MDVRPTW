import copy
import math
import numpy as np

from .utils import euclidean_distance


def _build_dist_matrix(depot_coord, customers_data):
    num_customers = len(customers_data)
    num_nodes = num_customers + 1
    dist_mat = np.zeros((num_nodes, num_nodes), dtype=float)
    all_coords = np.vstack([depot_coord, customers_data[:, :2]])
    for i in range(num_nodes):
        for j in range(num_nodes):
            dist_mat[i][j] = euclidean_distance(all_coords[i], all_coords[j])
    return dist_mat


def _construct_solution(
    dist_mat,
    pheromone,
    customers_data,
    vehicle_templates,
    time_windows,
    depot_tw,
    original_ids,
    alpha,
    beta,
    mutation_prob,
    time_penalty,
    rng,
):
    num_nodes = len(customers_data) + 1
    unvisited = set(range(1, num_nodes))
    routes = []
    total_cost = 0.0
    available_vehicles = copy.deepcopy(vehicle_templates)

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
        curr_time = depot_tw["start"]
        route_path = []
        route_dist = 0.0
        route_penalty = 0.0

        while True:
            valid_nodes = []
            scores = []

            for next_node in unvisited:
                demand = customers_data[next_node - 1][2]
                if curr_load + demand > chosen_vehicle["capacity"]:
                    continue

                dist = dist_mat[curr_node][next_node]
                arrival = curr_time + dist / chosen_vehicle["velocity"]
                oid = original_ids[next_node - 1]
                tw = time_windows[oid]

                service_start = max(arrival, tw["start"])
                finish_service = service_start + tw["service"]
                dist_back = dist_mat[next_node][0]
                time_back = dist_back / chosen_vehicle["velocity"]
                if finish_service + time_back > depot_tw["end"]:
                    continue

                slack = max(tw["end"] - service_start, 0.1)
                urgency = 1.0 / slack
                capacity_fit = demand / max(chosen_vehicle["capacity"] - curr_load, 1e-6)
                eta = (1.0 / (dist + 1e-6)) * (1.0 + 0.5 * urgency) * (1.0 + 0.2 * capacity_fit)
                tau = pheromone[curr_node][next_node]
                score = (tau ** alpha) * (eta ** beta)

                valid_nodes.append(next_node)
                scores.append(score)

            if not valid_nodes:
                break

            if rng.random() < mutation_prob:
                chosen_idx = int(rng.integers(0, len(valid_nodes)))
                next_node = valid_nodes[chosen_idx]
            else:
                probs = np.array(scores, dtype=float)
                if probs.sum() <= 0:
                    probs = np.ones(len(valid_nodes), dtype=float) / len(valid_nodes)
                else:
                    probs = probs / probs.sum()
                next_node = int(rng.choice(valid_nodes, p=probs))

            route_path.append(next_node - 1)
            unvisited.remove(next_node)
            curr_load += customers_data[next_node - 1][2]
            dist = dist_mat[curr_node][next_node]
            route_dist += dist

            arrival = curr_time + dist / chosen_vehicle["velocity"]
            oid = original_ids[next_node - 1]
            tw = time_windows[oid]
            if arrival > tw["end"]:
                route_penalty += (arrival - tw["end"]) * time_penalty

            curr_time = max(arrival, tw["start"]) + tw["service"]
            curr_node = next_node

        if route_path:
            dist_back = dist_mat[curr_node][0]
            route_dist += dist_back
            routes.append({"path": route_path, "vehicle": chosen_vehicle, "distance": route_dist})
            total_cost += route_dist * chosen_vehicle["cost"] + route_penalty

    return routes, total_cost


def solve_with_params(
    depot_coord,
    customers_data,
    vehicle_list,
    time_windows,
    depot_tw,
    original_ids,
    params,
    iterations,
    ant_count,
    seed,
):
    rng = np.random.default_rng(seed)
    dist_mat = _build_dist_matrix(depot_coord, customers_data)
    num_nodes = len(customers_data) + 1
    pheromone = np.ones((num_nodes, num_nodes), dtype=float)
    best_routes = None
    best_cost = float("inf")

    alpha = params["alpha"]
    beta = params["beta"]
    rho_min = params["rho_min"]
    rho_max = params["rho_max"]
    tau_min = params["tau_min"]
    tau_max = params["tau_max"]
    mutation_prob = params["mutation_prob"]
    time_penalty = params["time_penalty"]
    q = params["q"]
    elite_boost = params["elite_boost"]
    n_ball_ratio = params["n_ball_ratio"]
    n_small_ratio = params["n_small_ratio"]

    n_ball = max(1, int(ant_count * n_ball_ratio))
    n_small = max(1, int(ant_count * n_small_ratio))
    n_brood = max(1, int(ant_count * 0.4))

    for it in range(iterations):
        ant_routes = []
        ant_costs = []
        progress = it / max(iterations - 1, 1)
        rho = rho_max - (rho_max - rho_min) * progress

        for k in range(ant_count):
            a = alpha
            b = beta
            mut = 0.0

            if k < n_ball:
                a = alpha + 0.8
                b = max(1.0, beta - 0.8)
            elif k < n_ball + n_brood:
                a = alpha
                b = beta
            elif k < n_ball + n_brood + n_small:
                a = max(0.1, alpha - 0.9)
                b = beta + 1.5
            else:
                mut = mutation_prob

            routes, cost = _construct_solution(
                dist_mat,
                pheromone,
                customers_data,
                vehicle_list,
                time_windows,
                depot_tw,
                original_ids,
                a,
                b,
                mut,
                time_penalty,
                rng,
            )

            ant_routes.append(routes)
            ant_costs.append(cost)
            if cost < best_cost:
                best_cost = cost
                best_routes = routes

        pheromone *= (1 - rho)
        ranked = [
            (idx, rts, cst)
            for idx, (rts, cst) in enumerate(zip(ant_routes, ant_costs))
            if rts and cst not in (0, float("inf"))
        ]
        ranked.sort(key=lambda item: item[2])

        for rank, (idx, routes, cost) in enumerate(ranked[:5]):
            delta = q / cost
            if idx < n_ball:
                delta *= 2.0
            if rank == 0:
                delta *= elite_boost
            for r_info in routes:
                curr = 0
                for node_idx in r_info["path"]:
                    next_node = node_idx + 1
                    pheromone[curr][next_node] += delta
                    curr = next_node
                pheromone[curr][0] += delta

        if best_routes and best_cost < float("inf"):
            elite_delta = (q / best_cost) * 0.8
            for r_info in best_routes:
                curr = 0
                for node_idx in r_info["path"]:
                    next_node = node_idx + 1
                    pheromone[curr][next_node] += elite_delta
                    curr = next_node
                pheromone[curr][0] += elite_delta

        pheromone = np.clip(pheromone, tau_min, tau_max)

    return best_routes, best_cost


def tune_params_with_dbo(
    depot_coord,
    customers_data,
    vehicle_list,
    time_windows,
    depot_tw,
    original_ids,
    seed=42,
):
    bounds = [
        ("alpha", 0.8, 2.2),
        ("beta", 1.2, 4.5),
        ("rho_min", 0.05, 0.18),
        ("rho_max", 0.18, 0.35),
        ("tau_min", 0.01, 0.12),
        ("tau_max", 2.5, 8.0),
        ("mutation_prob", 0.05, 0.30),
        ("time_penalty", 60.0, 180.0),
        ("q", 60.0, 180.0),
        ("elite_boost", 1.2, 2.5),
        ("n_ball_ratio", 0.10, 0.35),
        ("n_small_ratio", 0.10, 0.35),
    ]
    dim = len(bounds)
    pop_size = 6
    max_iter = 4
    rng = np.random.default_rng(seed)

    lb = np.array([b[1] for b in bounds], dtype=float)
    ub = np.array([b[2] for b in bounds], dtype=float)
    pop = rng.uniform(lb, ub, size=(pop_size, dim))
    fitness = np.full(pop_size, float("inf"))
    cache = {}

    def decode(vec):
        d = {}
        for i, (name, _, _) in enumerate(bounds):
            d[name] = float(vec[i])
        if d["rho_min"] >= d["rho_max"]:
            d["rho_min"], d["rho_max"] = min(d["rho_min"], d["rho_max"] - 0.02), max(d["rho_max"], d["rho_min"] + 0.02)
        d["tau_max"] = max(d["tau_max"], d["tau_min"] + 0.5)
        ratio_sum = d["n_ball_ratio"] + d["n_small_ratio"]
        if ratio_sum > 0.75:
            d["n_ball_ratio"] *= 0.75 / ratio_sum
            d["n_small_ratio"] *= 0.75 / ratio_sum
        return d

    def obj(vec, eval_seed):
        key = tuple(np.round(vec, 3))
        if key in cache:
            return cache[key]
        params = decode(vec)
        _, cost = solve_with_params(
            depot_coord,
            customers_data,
            vehicle_list,
            time_windows,
            depot_tw,
            original_ids,
            params,
            iterations=25,
            ant_count=14,
            seed=eval_seed,
        )
        cache[key] = cost
        return cost

    for i in range(pop_size):
        fitness[i] = obj(pop[i], seed + i)

    best_idx = int(np.argmin(fitness))
    gbest = pop[best_idx].copy()
    gbest_fit = float(fitness[best_idx])

    n_ball = max(1, int(pop_size * 0.2))
    n_brood = max(1, int(pop_size * 0.4))
    n_small = max(1, int(pop_size * 0.2))

    for it in range(max_iter):
        order = np.argsort(fitness)
        pop = pop[order]
        fitness = fitness[order]
        worst = pop[-1].copy()

        for i in range(pop_size):
            cur = pop[i].copy()
            if i < n_ball:
                r1 = rng.random(dim)
                r2 = rng.random(dim)
                cur = cur + 0.35 * r1 * (gbest - cur) - 0.10 * r2 * (worst - cur)
            elif i < n_ball + n_brood:
                sigma = 0.22 * (1 - it / max(max_iter - 1, 1)) + 0.03
                cur = gbest + rng.normal(0, sigma, size=dim)
            elif i < n_ball + n_brood + n_small:
                cur = cur + 0.20 * rng.random(dim) * (gbest - cur) + 0.05 * (rng.random(dim) - cur)
            else:
                peer = pop[int(rng.integers(0, pop_size))]
                cur = cur + rng.normal(0, 0.12, size=dim) + 0.12 * (peer - cur)

            cur = np.clip(cur, lb, ub)
            fit = obj(cur, seed + 97 + it * pop_size + i)
            if fit < fitness[i]:
                pop[i] = cur
                fitness[i] = fit
                if fit < gbest_fit:
                    gbest = cur.copy()
                    gbest_fit = float(fit)

    return decode(gbest), gbest_fit
