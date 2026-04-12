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


class WholeNetworkACO:
    def __init__(self, customers_df, depots, vehicle_pool, tw_data):
        self.customers_df = customers_df.copy()
        self.depots = np.array(depots, dtype=float)
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
        self.pheromone = np.ones((self.num_customers + 1, self.num_customers + 1), dtype=float)
        self.heuristic = self._build_heuristic()

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

    def _build_heuristic(self):
        heuristic = np.ones((self.num_customers + 1, self.num_customers + 1), dtype=float)

        for c_idx in range(self.num_customers):
            start_dist = np.min(self.depot_customer_dist[:, c_idx])
            heuristic[0, c_idx + 1] = 1.0 / (start_dist + 1e-6)
            heuristic[c_idx + 1, 0] = heuristic[0, c_idx + 1]

        for i in range(self.num_customers):
            for j in range(self.num_customers):
                heuristic[i + 1, j + 1] = 1.0 / (self.customer_dist[i, j] + 1e-6)

        return heuristic

    def run(self):
        best_solution = None
        best_cost = float("inf")
        convergence = []

        for iteration in range(config.ACO_ITERATIONS):
            iteration_solutions = []
            iteration_costs = []

            for _ in range(config.ACO_ANT_COUNT):
                solution, total_cost = self.construct_solution()
                iteration_solutions.append(solution)
                iteration_costs.append(total_cost)

                if total_cost < best_cost:
                    best_cost = total_cost
                    best_solution = solution

            convergence.append(best_cost)
            self.update_pheromone(iteration_solutions, iteration_costs)

            if (iteration + 1) % 20 == 0:
                print(f"Iteration {iteration + 1}/{config.ACO_ITERATIONS}, Best Cost: {best_cost:.2f}")

        return best_solution, best_cost, convergence

    def construct_solution(self):
        unvisited = set(range(self.num_customers))
        remaining_vehicles = copy.deepcopy(self.vehicle_pool)
        routes = []
        total_cost = 0.0
        idle_rounds = 0

        while unvisited:
            feasible_vehicle_indices = [
                idx
                for idx, vehicle in enumerate(remaining_vehicles)
                if self._vehicle_has_feasible_start(vehicle, unvisited)
            ]
            if not feasible_vehicle_indices:
                return None, float("inf")

            vehicle_idx = self._select_vehicle(feasible_vehicle_indices, remaining_vehicles, unvisited)
            vehicle = remaining_vehicles.pop(vehicle_idx)

            route = self._build_route(vehicle, unvisited)
            if route is None:
                idle_rounds += 1
                if idle_rounds > len(self.vehicle_pool):
                    return None, float("inf")
                continue

            idle_rounds = 0
            routes.append(route)
            total_cost += route["cost"]

            for c_idx in route["path"]:
                unvisited.remove(c_idx)

        return routes, total_cost

    def _vehicle_has_feasible_start(self, vehicle, unvisited):
        depot_idx = vehicle["depot_idx"]
        depot_tw = self.tw_data["depots"][depot_idx]
        start_time = depot_tw["start"]
        end_time = depot_tw["end"]

        for c_idx in unvisited:
            if self.demands[c_idx] > vehicle["capacity"]:
                continue

            travel = self.depot_customer_dist[depot_idx, c_idx] / vehicle["velocity"]
            arrival = start_time + travel
            tw = self.customer_tw[int(self.customer_ids[c_idx])]
            service_start = max(arrival, tw["start"])
            if service_start > tw["end"]:
                continue

            finish_service = service_start + tw["service"]
            return_time = finish_service + self.depot_customer_dist[depot_idx, c_idx] / vehicle["velocity"]
            if return_time <= end_time:
                return True

        return False

    def _select_vehicle(self, feasible_vehicle_indices, vehicles, unvisited):
        scored = []
        for idx in feasible_vehicle_indices:
            vehicle = vehicles[idx]
            nearest = min(
                self.depot_customer_dist[vehicle["depot_idx"], c_idx]
                for c_idx in unvisited
                if self.demands[c_idx] <= vehicle["capacity"]
            )
            score = nearest * vehicle["var_cost"] + 0.2 * vehicle["fixed_cost"]
            scored.append((score, idx))

        scored.sort(key=lambda item: item[0])
        shortlist = scored[: min(3, len(scored))]
        return shortlist[np.random.randint(0, len(shortlist))][1]

    def _build_route(self, vehicle, unvisited):
        depot_idx = vehicle["depot_idx"]
        depot_tw = self.tw_data["depots"][depot_idx]
        route_unvisited = set(unvisited)

        curr_customer = None
        curr_time = depot_tw["start"]
        curr_load = 0.0
        path = []
        arrival_records = []
        route_distance = 0.0

        while True:
            candidates = []
            candidate_scores = []

            for c_idx in route_unvisited:
                demand = self.demands[c_idx]
                if curr_load + demand > vehicle["capacity"]:
                    continue

                move_dist = self._distance_from_state(depot_idx, curr_customer, c_idx)
                travel_time = move_dist / vehicle["velocity"]
                arrival = curr_time + travel_time
                tw = self.customer_tw[int(self.customer_ids[c_idx])]
                service_start = max(arrival, tw["start"])

                if service_start > tw["end"]:
                    continue

                finish_service = service_start + tw["service"]
                return_time = finish_service + self.depot_customer_dist[depot_idx, c_idx] / vehicle["velocity"]
                if return_time > depot_tw["end"]:
                    continue

                urgency = 1.0 / max(tw["end"] - arrival, 0.1)
                last_index = 0 if curr_customer is None else curr_customer + 1
                pheromone = self.pheromone[last_index, c_idx + 1] ** config.ACO_ALPHA
                heuristic = self.heuristic[last_index, c_idx + 1] ** config.ACO_BETA
                score = pheromone * heuristic * urgency

                candidates.append(
                    {
                        "customer_idx": c_idx,
                        "move_dist": move_dist,
                        "arrival": arrival,
                        "service_start": service_start,
                        "finish_service": finish_service,
                    }
                )
                candidate_scores.append(score)

            if not candidates:
                break

            probs = np.array(candidate_scores, dtype=float)
            if probs.sum() == 0:
                probs = np.ones(len(candidates), dtype=float) / len(candidates)
            else:
                probs = probs / probs.sum()

            chosen = candidates[np.random.choice(len(candidates), p=probs)]
            c_idx = chosen["customer_idx"]

            path.append(c_idx)
            route_unvisited.remove(c_idx)
            arrival_records.append(
                {
                    "customer_id": int(self.customer_ids[c_idx]),
                    "arrival_time": round(chosen["arrival"], 2),
                    "service_start": round(chosen["service_start"], 2),
                    "finish_service": round(chosen["finish_service"], 2),
                }
            )
            route_distance += chosen["move_dist"]
            curr_time = chosen["finish_service"]
            curr_load += self.demands[c_idx]
            curr_customer = c_idx

        if not path:
            return None

        return_dist = self.depot_customer_dist[depot_idx, curr_customer]
        route_distance += return_dist
        # Keep whole-network Total_Cost comparable to the decomposition tables:
        # use route distance cost only, without adding fixed vehicle startup cost.
        total_cost = route_distance * vehicle["var_cost"]

        return {
            "depot_idx": depot_idx,
            "vehicle": vehicle,
            "path": path,
            "distance": route_distance,
            "load": curr_load,
            "cost": total_cost,
            "timeline": arrival_records,
            "finish_time": round(curr_time + return_dist / vehicle["velocity"], 2),
        }

    def _distance_from_state(self, depot_idx, curr_customer, next_customer):
        if curr_customer is None:
            return self.depot_customer_dist[depot_idx, next_customer]
        return self.customer_dist[curr_customer, next_customer]

    def update_pheromone(self, solutions, costs):
        self.pheromone *= (1 - config.ACO_RHO)

        ranked = sorted(
            [
                (solution, cost)
                for solution, cost in zip(solutions, costs)
                if solution is not None and np.isfinite(cost) and cost > 0
            ],
            key=lambda item: item[1],
        )

        elite_count = min(5, len(ranked))
        for rank, (solution, cost) in enumerate(ranked[:elite_count]):
            delta = config.ACO_Q / cost
            if rank == 0:
                delta *= 2.0

            for route in solution:
                previous = 0
                for c_idx in route["path"]:
                    next_idx = c_idx + 1
                    self.pheromone[previous, next_idx] += delta
                    self.pheromone[next_idx, previous] += delta
                    previous = next_idx
                self.pheromone[previous, 0] += delta
                self.pheromone[0, previous] += delta


def load_vehicle_pool(num_depots):
    if config.USE_CUSTOM_VEHICLES and os.path.exists(config.CUSTOM_DATA_FILE):
        df = pd.read_excel(config.CUSTOM_DATA_FILE, sheet_name="Vehicles")
        df.columns = df.columns.str.lower().str.strip()

        vehicles = []
        for _, row in df.iterrows():
            depot_idx = int(row["depot_id"]) - 1
            if not (0 <= depot_idx < num_depots):
                continue

            count = int(row["count"])
            for instance_id in range(1, count + 1):
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


def plot_depot_routes(output_dir, depots, customers_df, routes):
    for depot_idx in range(len(depots)):
        depot_routes = [route for route in routes if route["depot_idx"] == depot_idx]
        if not depot_routes:
            continue

        plt.figure(figsize=(10, 8))
        plt.scatter(
            customers_df["x"],
            customers_df["y"],
            c="lightgray",
            s=30,
            alpha=0.5,
            label="All Customers",
            zorder=2,
        )
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

        plt.title(f"Whole ACO Routes - Depot {depot_idx + 1}")
        plt.xlabel("X (km)")
        plt.ylabel("Y (km)")
        plt.grid(True, linestyle="--", alpha=0.4)
        plt.legend()
        plt.tight_layout()
        plt.savefig(os.path.join(output_dir, f"whole_route_depot_{depot_idx + 1}.png"), dpi=300)
        plt.close()


def plot_convergence(output_dir, convergence):
    plt.figure(figsize=(9, 5))
    plt.plot(range(1, len(convergence) + 1), convergence, color="#1f77b4", linewidth=2)
    plt.title("Whole ACO Convergence Curve")
    plt.xlabel("Iteration")
    plt.ylabel("Best Cost")
    plt.grid(True, linestyle="--", alpha=0.4)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "whole_aco_convergence.png"), dpi=300)
    plt.close()


def save_route_results(output_dir, depots, customers_df, routes, best_cost, convergence):
    records = []
    for route_idx, route in enumerate(routes, start=1):
        depot_idx = route["depot_idx"]
        customer_ids = [int(customers_df.loc[c_idx, "id"]) for c_idx in route["path"]]
        route_str = " -> ".join(map(str, customer_ids))
        timeline_str = " | ".join(
            f"{item['customer_id']}@{item['service_start']}"
            for item in route["timeline"]
        )

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
                "Route": f"Depot {depot_idx + 1} -> {route_str} -> Depot {depot_idx + 1}",
                "Timeline": timeline_str,
            }
        )

    df_routes = pd.DataFrame(records)
    df_routes.to_csv(os.path.join(output_dir, "whole_aco_routes.csv"), index=False)

    depot_summary = (
        df_routes.groupby("Depot_ID")[["Load", "Distance", "Total_Cost"]]
        .sum()
        .reset_index()
        .rename(columns={"Load": "Total_Load"})
    )
    depot_summary.to_csv(os.path.join(output_dir, "whole_aco_depot_summary.csv"), index=False)

    with open(os.path.join(output_dir, "whole_aco_summary.txt"), "w", encoding="utf-8") as file:
        file.write("Whole ACO Summary\n")
        file.write(f"Total routes: {len(routes)}\n")
        file.write(f"Best total cost: {best_cost:.2f}\n")
        file.write(f"Total distance: {df_routes['Distance'].sum():.2f}\n")
        file.write(f"Final convergence value: {convergence[-1]:.2f}\n")


def load_problem_data():
    if config.USE_CUSTOM_DATA:
        try:
            customers_df, depots, _ = load_data_from_excel(config.CUSTOM_DATA_FILE)
        except Exception as exc:
            print(f"Error loading Excel data: {exc}")
            sys.exit(1)
    else:
        customers_df = generate_mock_data(config.RANDOM_NUM_CUSTOMERS, config.RANDOM_MAP_SIZE)
        depots = config.RANDOM_DEPOTS

    if "id" not in customers_df.columns:
        customers_df["id"] = range(1, len(customers_df) + 1)

    tw_data = load_timewindow_data(customers_df, len(depots))
    vehicle_pool = load_vehicle_pool(len(depots))
    return customers_df, np.array(depots, dtype=float), vehicle_pool, tw_data


def main():
    output_dir = os.path.join("results", "whole_aco")
    os.makedirs(output_dir, exist_ok=True)
    np.random.seed(42)

    print(">>> Loading whole-network MDVRPTW data ...")
    customers_df, depots, vehicle_pool, tw_data = load_problem_data()

    print(f"Customers: {len(customers_df)}")
    print(f"Depots: {len(depots)}")
    print(f"Vehicle instances: {len(vehicle_pool)}")

    solver = WholeNetworkACO(customers_df, depots, vehicle_pool, tw_data)
    best_routes, best_cost, convergence = solver.run()

    if best_routes is None:
        print("No feasible whole-network solution was found.")
        sys.exit(1)

    print("\n=== Whole ACO Result ===")
    total_distance = sum(route["distance"] for route in best_routes)
    for route_idx, route in enumerate(best_routes, start=1):
        vehicle = route["vehicle"]
        customer_ids = [int(customers_df.loc[c_idx, "id"]) for c_idx in route["path"]]
        route_str = " -> ".join(map(str, customer_ids))
        print(
            f"Route {route_idx}: Depot {route['depot_idx'] + 1}, "
            f"Vehicle {vehicle['type_id']}-{vehicle['vehicle_no']}, "
            f"Load {route['load']:.1f}/{vehicle['capacity']:.1f}, "
            f"Cost {route['cost']:.2f}, "
            f"Path: Depot -> {route_str} -> Depot"
        )

    print(f"Total distance: {total_distance:.2f}")
    print(f"Best total cost: {best_cost:.2f}")

    save_route_results(output_dir, depots, customers_df, best_routes, best_cost, convergence)
    plot_depot_routes(output_dir, depots, customers_df, best_routes)
    plot_convergence(output_dir, convergence)

    print(f"\nResults saved to: {output_dir}")


if __name__ == "__main__":
    main()
