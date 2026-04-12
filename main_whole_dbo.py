import copy
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from src import config
from src.utils import generate_mock_data, load_data_from_excel, load_timewindow_data


def load_vehicle_pool(num_depots):
    if config.USE_CUSTOM_VEHICLES and os.path.exists(config.CUSTOM_DATA_FILE):
        df = pd.read_excel(config.CUSTOM_DATA_FILE, sheet_name="Vehicles")
        df.columns = df.columns.str.lower().str.strip()

        vehicles = []
        for _, row in df.iterrows():
            depot_idx = int(row["depot_id"]) - 1
            if not (0 <= depot_idx < num_depots):
                continue

            for instance_id in range(1, int(row["count"]) + 1):
                vehicles.append(
                    {
                        "depot_idx": depot_idx,
                        "type_id": int(row["type_id"]),
                        "vehicle_no": instance_id,
                        "capacity": float(row["capacity"]),
                        "velocity": float(row["velocity"]),
                        "fixed_cost": float(row.get("fixed_cost", 0.0)),
                        "var_cost": float(row.get("var_cost", 1.0)),
                    }
                )
        return vehicles

    vehicles = []
    for depot_idx in range(num_depots):
        vehicle_type_id = 1
        for v_conf in config.RANDOM_VEHICLE_TYPES:
            count = np.random.randint(v_conf["count"][0], v_conf["count"][1] + 1)
            for instance_id in range(1, count + 1):
                vehicles.append(
                    {
                        "depot_idx": depot_idx,
                        "type_id": vehicle_type_id,
                        "vehicle_no": instance_id,
                        "capacity": float(v_conf["cap"]),
                        "velocity": float(v_conf["speed"]),
                        "fixed_cost": 0.0,
                        "var_cost": float(v_conf["cost"]),
                    }
                )
            vehicle_type_id += 1
    return vehicles


def load_problem_data():
    if config.USE_CUSTOM_DATA:
        try:
            customers_df, depots, depot_capacities = load_data_from_excel(config.CUSTOM_DATA_FILE)
        except Exception as exc:
            print(f"Error loading Excel data: {exc}")
            sys.exit(1)
    else:
        customers_df = generate_mock_data(config.RANDOM_NUM_CUSTOMERS, config.RANDOM_MAP_SIZE)
        depots = config.RANDOM_DEPOTS
        depot_capacities = np.full(len(depots), 1e9)

    if "id" not in customers_df.columns:
        customers_df["id"] = range(1, len(customers_df) + 1)

    tw_data = load_timewindow_data(customers_df, len(depots))
    vehicle_pool = load_vehicle_pool(len(depots))
    return customers_df, np.array(depots, dtype=float), np.array(depot_capacities, dtype=float), vehicle_pool, tw_data


class WholeNetworkDBO:
    def __init__(self, customers_df, depots, depot_capacities, vehicle_pool, tw_data):
        self.customers_df = customers_df.copy()
        self.depots = np.array(depots, dtype=float)
        self.depot_capacities = np.array(depot_capacities, dtype=float)
        self.vehicle_pool = vehicle_pool
        self.tw_data = tw_data

        self.customer_ids = self.customers_df["id"].astype(int).to_numpy()
        self.customer_coords = self.customers_df[["x", "y"]].to_numpy(dtype=float)
        self.demands = self.customers_df["demand"].to_numpy(dtype=float)
        self.num_customers = len(self.customer_ids)
        self.num_depots = len(self.depots)

        self.customer_tw = {
            int(cid): tw_data["customers"][int(cid)] for cid in self.customer_ids
        }
        self.customer_dist = self._build_customer_dist()
        self.depot_customer_dist = self._build_depot_customer_dist()

    def _build_customer_dist(self):
        dist = np.zeros((self.num_customers, self.num_customers), dtype=float)
        for i in range(self.num_customers):
            for j in range(self.num_customers):
                dist[i, j] = np.linalg.norm(self.customer_coords[i] - self.customer_coords[j])
        return dist

    def _build_depot_customer_dist(self):
        dist = np.zeros((self.num_depots, self.num_customers), dtype=float)
        for d_idx in range(self.num_depots):
            for c_idx in range(self.num_customers):
                dist[d_idx, c_idx] = np.linalg.norm(self.depots[d_idx] - self.customer_coords[c_idx])
        return dist

    def optimize(self):
        pop_size = config.DBO_POP_SIZE
        max_iter = config.DBO_MAX_ITER
        dim = self.num_customers

        population = np.random.rand(pop_size, dim)
        fitness = np.full(pop_size, float("inf"))
        decoded_cache = [None] * pop_size

        gbest = None
        gbest_fit = float("inf")
        gbest_routes = None
        curve = []

        for i in range(pop_size):
            routes, cost = self.decode_solution(population[i])
            fitness[i] = cost
            decoded_cache[i] = routes
            if cost < gbest_fit:
                gbest_fit = cost
                gbest = population[i].copy()
                gbest_routes = routes

        for iteration in range(max_iter):
            order = np.argsort(fitness)
            population = population[order]
            fitness = fitness[order]
            decoded_cache = [decoded_cache[idx] for idx in order]

            worst = population[-1].copy()
            n_ball = max(1, int(pop_size * 0.2))
            n_brood = max(1, int(pop_size * 0.3))
            n_small = max(1, int(pop_size * 0.3))

            for i in range(pop_size):
                candidate = population[i].copy()

                if i < n_ball:
                    candidate = self._ball_rolling_update(candidate, gbest, worst)
                elif i < n_ball + n_brood:
                    candidate = self._brood_update(candidate, gbest, iteration, max_iter)
                elif i < n_ball + n_brood + n_small:
                    candidate = self._small_update(candidate, gbest)
                else:
                    candidate = self._thief_update(candidate, population[np.random.randint(pop_size)])

                candidate = np.clip(candidate, 0.0, 1.0)
                routes, cost = self.decode_solution(candidate)

                if cost < fitness[i]:
                    population[i] = candidate
                    fitness[i] = cost
                    decoded_cache[i] = routes

                    if cost < gbest_fit:
                        gbest_fit = cost
                        gbest = candidate.copy()
                        gbest_routes = routes

            curve.append(gbest_fit)
            if (iteration + 1) % 20 == 0:
                print(f"Iteration {iteration + 1}/{max_iter}, Best Cost: {gbest_fit:.2f}")

        return gbest, gbest_routes, gbest_fit, curve

    def _ball_rolling_update(self, candidate, gbest, worst):
        r1 = np.random.rand(self.num_customers)
        r2 = np.random.rand(self.num_customers)
        updated = candidate + 0.35 * r1 * (gbest - candidate) - 0.10 * r2 * (worst - candidate)
        if np.random.rand() < 0.2:
            updated = self._swap_mutation(updated, swaps=2)
        return updated

    def _brood_update(self, candidate, gbest, iteration, max_iter):
        sigma = 0.20 * (1 - iteration / max_iter) + 0.02
        updated = gbest + np.random.normal(0, sigma, self.num_customers)
        if np.random.rand() < 0.3:
            updated = 0.5 * updated + 0.5 * candidate
        return updated

    def _small_update(self, candidate, gbest):
        peer = np.random.rand(self.num_customers)
        updated = candidate + 0.20 * np.random.rand(self.num_customers) * (gbest - candidate) + 0.05 * (peer - candidate)
        if np.random.rand() < 0.4:
            updated = self._swap_mutation(updated, swaps=1)
        return updated

    def _thief_update(self, candidate, peer):
        updated = candidate + np.random.normal(0, 0.12, self.num_customers) + 0.15 * (peer - candidate)
        updated = self._swap_mutation(updated, swaps=2)
        return updated

    def _swap_mutation(self, vector, swaps=1):
        mutated = vector.copy()
        for _ in range(swaps):
            i, j = np.random.choice(self.num_customers, 2, replace=False)
            mutated[i], mutated[j] = mutated[j], mutated[i]
        return mutated

    def decode_solution(self, priority_vector):
        order = np.argsort(priority_vector)
        available_vehicles = copy.deepcopy(self.vehicle_pool)
        routes = []
        depot_loads = np.zeros(self.num_depots, dtype=float)

        for c_idx in order:
            assigned = False
            best_option = None

            for route_idx, route in enumerate(routes):
                feasible, delta_dist, new_state = self._try_append_to_route(route, c_idx, depot_loads)
                if feasible:
                    delta_cost = delta_dist * route["vehicle"]["var_cost"]
                    score = delta_cost + 0.01 * new_state["finish_time"]
                    if best_option is None or score < best_option["score"]:
                        best_option = {
                            "kind": "append",
                            "route_idx": route_idx,
                            "score": score,
                            "delta_cost": delta_cost,
                            "new_state": new_state,
                        }

            for vehicle_idx, vehicle in enumerate(available_vehicles):
                feasible, start_route = self._try_start_route(vehicle, c_idx, depot_loads)
                if feasible:
                    score = start_route["cost"] + 0.01 * start_route["finish_time"]
                    if best_option is None or score < best_option["score"]:
                        best_option = {
                            "kind": "new",
                            "vehicle_idx": vehicle_idx,
                            "score": score,
                            "new_route": start_route,
                        }

            if best_option is None:
                return None, float("inf")

            if best_option["kind"] == "append":
                route = routes[best_option["route_idx"]]
                state = best_option["new_state"]
                route["path"].append(c_idx)
                route["timeline"].append(state["timeline_item"])
                route["load"] += self.demands[c_idx]
                route["distance"] = state["distance"]
                route["cost"] += best_option["delta_cost"]
                route["current_time"] = state["current_time"]
                route["finish_time"] = state["finish_time"]
                route["last_customer"] = c_idx
            else:
                route = best_option["new_route"]
                routes.append(route)
                depot_loads[route["depot_idx"]] += self.demands[c_idx]
                available_vehicles.pop(best_option["vehicle_idx"])
                assigned = True

            if not assigned and best_option["kind"] == "append":
                depot_loads[route["depot_idx"]] += self.demands[c_idx]

        total_cost = sum(route["cost"] for route in routes)
        return routes, total_cost

    def _try_start_route(self, vehicle, c_idx, depot_loads):
        depot_idx = vehicle["depot_idx"]
        depot_tw = self.tw_data["depots"][depot_idx]
        tw = self.customer_tw[int(self.customer_ids[c_idx])]

        if self.demands[c_idx] > vehicle["capacity"]:
            return False, None
        if depot_loads[depot_idx] + self.demands[c_idx] > self.depot_capacities[depot_idx]:
            return False, None

        travel = self.depot_customer_dist[depot_idx, c_idx] / vehicle["velocity"]
        arrival = depot_tw["start"] + travel
        service_start = max(arrival, tw["start"])
        if service_start > tw["end"]:
            return False, None

        current_time = service_start + tw["service"]
        finish_time = current_time + self.depot_customer_dist[depot_idx, c_idx] / vehicle["velocity"]
        if finish_time > depot_tw["end"]:
            return False, None

        distance = 2 * self.depot_customer_dist[depot_idx, c_idx]
        # Keep whole-network Total_Cost comparable to the decomposition tables:
        # use route distance cost only, without adding fixed vehicle startup cost.
        cost = distance * vehicle["var_cost"]

        route = {
            "depot_idx": depot_idx,
            "vehicle": vehicle,
            "path": [c_idx],
            "load": self.demands[c_idx],
            "distance": distance,
            "cost": cost,
            "current_time": current_time,
            "finish_time": round(finish_time, 2),
            "last_customer": c_idx,
            "timeline": [
                {
                    "customer_id": int(self.customer_ids[c_idx]),
                    "arrival_time": round(arrival, 2),
                    "service_start": round(service_start, 2),
                    "finish_service": round(current_time, 2),
                }
            ],
        }
        return True, route

    def _try_append_to_route(self, route, c_idx, depot_loads):
        depot_idx = route["depot_idx"]
        vehicle = route["vehicle"]
        depot_tw = self.tw_data["depots"][depot_idx]
        tw = self.customer_tw[int(self.customer_ids[c_idx])]
        last_customer = route["last_customer"]

        if c_idx in route["path"]:
            return False, None, None
        if route["load"] + self.demands[c_idx] > vehicle["capacity"]:
            return False, None, None
        if depot_loads[depot_idx] + self.demands[c_idx] > self.depot_capacities[depot_idx]:
            return False, None, None

        travel = self.customer_dist[last_customer, c_idx] / vehicle["velocity"]
        arrival = route["current_time"] + travel
        service_start = max(arrival, tw["start"])
        if service_start > tw["end"]:
            return False, None, None

        current_time = service_start + tw["service"]
        finish_time = current_time + self.depot_customer_dist[depot_idx, c_idx] / vehicle["velocity"]
        if finish_time > depot_tw["end"]:
            return False, None, None

        previous_return = self.depot_customer_dist[depot_idx, last_customer]
        new_return = self.depot_customer_dist[depot_idx, c_idx]
        added_dist = self.customer_dist[last_customer, c_idx] + new_return - previous_return
        new_distance = route["distance"] + added_dist

        new_state = {
            "distance": new_distance,
            "current_time": current_time,
            "finish_time": round(finish_time, 2),
            "timeline_item": {
                "customer_id": int(self.customer_ids[c_idx]),
                "arrival_time": round(arrival, 2),
                "service_start": round(service_start, 2),
                "finish_service": round(current_time, 2),
            },
        }
        return True, added_dist, new_state


def plot_depot_routes(output_dir, depots, customers_df, routes):
    for depot_idx in range(len(depots)):
        depot_routes = [route for route in routes if route["depot_idx"] == depot_idx]
        if not depot_routes:
            continue

        plt.figure(figsize=(10, 8))
        plt.scatter(customers_df["x"], customers_df["y"], c="lightgray", s=30, alpha=0.5, label="All Customers", zorder=2)
        plt.scatter(
            depots[depot_idx][0],
            depots[depot_idx][1],
            c="red",
            marker="s",
            s=220,
            edgecolors="black",
            label=f"Depot {depot_idx + 1}",
            zorder=10,
        )

        colors = plt.cm.tab10(np.linspace(0, 1, max(1, len(depot_routes))))
        for color, route in zip(colors, depot_routes):
            coords = [depots[depot_idx]]
            for c_idx in route["path"]:
                coords.append(customers_df.loc[c_idx, ["x", "y"]].to_numpy(dtype=float))
            coords.append(depots[depot_idx])
            coords = np.array(coords, dtype=float)
            plt.plot(coords[:, 0], coords[:, 1], c=color, linewidth=2.0, alpha=0.9, zorder=5)
            plt.scatter(coords[1:-1, 0], coords[1:-1, 1], c=[color], s=45, zorder=6)

        plt.title(f"Whole DBO Routes - Depot {depot_idx + 1}")
        plt.xlabel("X (km)")
        plt.ylabel("Y (km)")
        plt.grid(True, linestyle="--", alpha=0.4)
        plt.legend()
        plt.tight_layout()
        plt.savefig(os.path.join(output_dir, f"whole_dbo_route_depot_{depot_idx + 1}.png"), dpi=300)
        plt.close()


def plot_convergence(output_dir, convergence):
    plt.figure(figsize=(9, 5))
    plt.plot(range(1, len(convergence) + 1), convergence, color="#d62728", linewidth=2)
    plt.title("Whole DBO Convergence Curve")
    plt.xlabel("Iteration")
    plt.ylabel("Best Cost")
    plt.grid(True, linestyle="--", alpha=0.4)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "whole_dbo_convergence.png"), dpi=300)
    plt.close()


def save_route_results(output_dir, customers_df, routes, best_cost, convergence):
    records = []
    for route_idx, route in enumerate(routes, start=1):
        depot_idx = route["depot_idx"]
        customer_ids = [int(customers_df.loc[c_idx, "id"]) for c_idx in route["path"]]
        timeline_str = " | ".join(f"{item['customer_id']}@{item['service_start']}" for item in route["timeline"])
        vehicle = route["vehicle"]

        records.append(
            {
                "Route_ID": route_idx,
                "Depot_ID": depot_idx + 1,
                "Vehicle_Type": vehicle["type_id"],
                "Vehicle_No": vehicle["vehicle_no"],
                "Capacity": vehicle["capacity"],
                "Load": round(route["load"], 2),
                "Distance": round(route["distance"], 2),
                "Fixed_Cost": round(vehicle["fixed_cost"], 2),
                "Variable_Cost": round(route["distance"] * vehicle["var_cost"], 2),
                "Total_Cost": round(route["cost"], 2),
                "Finish_Time": route["finish_time"],
                "Route": f"Depot {depot_idx + 1} -> {' -> '.join(map(str, customer_ids))} -> Depot {depot_idx + 1}",
                "Timeline": timeline_str,
            }
        )

    df_routes = pd.DataFrame(records)
    df_routes.to_csv(os.path.join(output_dir, "whole_dbo_routes.csv"), index=False)

    depot_summary = (
        df_routes.groupby("Depot_ID")[["Load", "Distance", "Total_Cost"]]
        .sum()
        .reset_index()
        .rename(columns={"Load": "Total_Load"})
    )
    depot_summary.to_csv(os.path.join(output_dir, "whole_dbo_depot_summary.csv"), index=False)

    with open(os.path.join(output_dir, "whole_dbo_summary.txt"), "w", encoding="utf-8") as file:
        file.write("Whole DBO Summary\n")
        file.write(f"Total routes: {len(routes)}\n")
        file.write(f"Best total cost: {best_cost:.2f}\n")
        file.write(f"Total distance: {df_routes['Distance'].sum():.2f}\n")
        file.write(f"Final convergence value: {convergence[-1]:.2f}\n")


def main():
    output_dir = os.path.join("results", "whole_dbo")
    os.makedirs(output_dir, exist_ok=True)
    np.random.seed(42)

    print(">>> Loading whole-network MDVRPTW data for DBO ...")
    customers_df, depots, depot_capacities, vehicle_pool, tw_data = load_problem_data()

    print(f"Customers: {len(customers_df)}")
    print(f"Depots: {len(depots)}")
    print(f"Vehicle instances: {len(vehicle_pool)}")

    solver = WholeNetworkDBO(customers_df, depots, depot_capacities, vehicle_pool, tw_data)
    _, best_routes, best_cost, convergence = solver.optimize()

    if best_routes is None:
        print("No feasible whole-network DBO solution was found.")
        sys.exit(1)

    print("\n=== Whole DBO Result ===")
    total_distance = sum(route["distance"] for route in best_routes)
    for route_idx, route in enumerate(best_routes, start=1):
        vehicle = route["vehicle"]
        customer_ids = [int(customers_df.loc[c_idx, "id"]) for c_idx in route["path"]]
        print(
            f"Route {route_idx}: Depot {route['depot_idx'] + 1}, "
            f"Vehicle {vehicle['type_id']}-{vehicle['vehicle_no']}, "
            f"Load {route['load']:.1f}/{vehicle['capacity']:.1f}, "
            f"Cost {route['cost']:.2f}, "
            f"Path: Depot -> {' -> '.join(map(str, customer_ids))} -> Depot"
        )

    print(f"Total distance: {total_distance:.2f}")
    print(f"Best total cost: {best_cost:.2f}")

    save_route_results(output_dir, customers_df, best_routes, best_cost, convergence)
    plot_depot_routes(output_dir, depots, customers_df, best_routes)
    plot_convergence(output_dir, convergence)

    print(f"\nResults saved to: {output_dir}")


if __name__ == "__main__":
    main()
