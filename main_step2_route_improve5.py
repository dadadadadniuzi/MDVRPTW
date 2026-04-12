import os
import sys

import matplotlib
matplotlib.use("Agg")
import numpy as np
import pandas as pd

from src import config
from src.aco_improved5 import ACO_Improved5, tune_params_by_dbo
from src.utils import load_data_from_excel, load_vehicle_data, load_timewindow_data, plot_vehicle_routes


def get_phase1_results():
    csv_path = os.path.join("results", "allocation_results.csv")
    if not os.path.exists(csv_path):
        print(f"Error: cannot find {csv_path}")
        sys.exit(1)

    print(f">>> Loading phase-1 allocation result: {csv_path}")
    df = pd.read_csv(csv_path)
    if config.USE_CUSTOM_DATA:
        _, fixed_depots, _ = load_data_from_excel(config.CUSTOM_DATA_FILE)
    else:
        fixed_depots = config.RANDOM_DEPOTS
    return df, fixed_depots


def main():
    output_dir = os.path.join("results", "improved5_aco")
    os.makedirs(output_dir, exist_ok=True)
    np.random.seed(42)

    df_customers, depots = get_phase1_results()
    num_depots = len(depots)

    print("\n>>> Loading phase-2 data (vehicles & time windows)...")
    vehicles_data = load_vehicle_data(num_depots)
    tw_data = load_timewindow_data(df_customers, num_depots)

    total_system_cost = 0.0
    all_routes_records = []
    tuned_rows = []

    print("\n" + "=" * 62)
    print("Start Phase 2: Improve5 (DBO-tuned ACO) for decomposed VRPTW")
    print("=" * 62)

    for depot_idx in range(num_depots):
        print(f"\n>>> Optimizing Depot {depot_idx + 1} ...")
        cluster_df = df_customers[df_customers["label"] == depot_idx]
        if len(cluster_df) == 0:
            print("  No customers in this depot, skip.")
            continue

        cluster_data = cluster_df[["x", "y", "demand"]].values
        original_ids = cluster_df["id"].values
        depot_coord = depots[depot_idx]
        v_list = vehicles_data[depot_idx]
        depot_tw = tw_data["depots"][depot_idx]

        print(f"  Customers: {len(cluster_df)}")
        print(f"  Vehicles: {sum(v['count'] for v in v_list)}")
        print("  Running DBO to tune ACO parameters ...")

        problem_args = (
            depot_coord,
            cluster_data,
            v_list,
            tw_data["customers"],
            depot_tw,
            original_ids,
        )
        best_params, eval_cost = tune_params_by_dbo(problem_args, seed=100 + depot_idx * 97)
        print(f"  DBO tuned eval cost: {eval_cost:.2f}")

        solver = ACO_Improved5(*problem_args, params=best_params)
        best_routes, best_cost, curve = solver.run(
            iterations=config.ACO_ITERATIONS,
            ant_count=config.ACO_ANT_COUNT,
            seed=2026 + depot_idx,
        )
        print(f"  Convergence tail: {curve[-5:]}")

        if best_routes is None:
            print("  [Failed] No feasible solution found.")
            continue

        print(f"  [Done] Best cost: {best_cost:.2f}")
        total_system_cost += best_cost

        tuned_row = {"Depot_ID": depot_idx + 1, "Tuned_Eval_Cost": round(eval_cost, 4)}
        tuned_row.update({k: round(v, 6) for k, v in best_params.items()})
        tuned_rows.append(tuned_row)

        simple_routes_for_plot = []
        for i, r_info in enumerate(best_routes):
            route_indices = r_info["path"]
            v_info = r_info["vehicle"]
            dist = r_info["distance"]

            real_ids = [original_ids[idx] for idx in route_indices]
            route_str = " -> ".join(map(str, real_ids))
            print(f"    Vehicle {i + 1} (Type {v_info['type_id']}): Depot -> {route_str} -> Depot")

            all_routes_records.append(
                {
                    "Depot_ID": depot_idx + 1,
                    "Vehicle_ID": i + 1,
                    "Vehicle_Type": v_info["type_id"],
                    "Capacity": v_info["capacity"],
                    "Route": f"0 -> {route_str} -> 0",
                    "Distance": round(dist, 2),
                    "Cost": round(dist * v_info["cost"], 2),
                }
            )
            simple_routes_for_plot.append(route_indices)

        filename = f"improved5_route_depot_{depot_idx + 1}.png"
        plot_vehicle_routes(
            depot_coord,
            cluster_data,
            simple_routes_for_plot,
            title=f"Improved5 ACO - Depot {depot_idx + 1} (Cost: {best_cost:.1f})",
            filename=f"improved5_aco/{filename}",
        )

    print("\n" + "=" * 62)
    print(f"Total system cost: {total_system_cost:.2f}")

    if all_routes_records:
        routes_csv = os.path.join(output_dir, "improved5_final_routes.csv")
        pd.DataFrame(all_routes_records).to_csv(routes_csv, index=False)
        print(f"[Saved] final route file: {routes_csv}")

    if tuned_rows:
        params_csv = os.path.join(output_dir, "improved5_tuned_params.csv")
        pd.DataFrame(tuned_rows).to_csv(params_csv, index=False)
        print(f"[Saved] tuned parameter file: {params_csv}")

    print("=" * 62)


if __name__ == "__main__":
    main()
