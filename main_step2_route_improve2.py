import os
import sys

import matplotlib
matplotlib.use("Agg")
import pandas as pd

from src import config
from src.aco_improved2 import ACO_Improved2
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
    output_dir = os.path.join("results", "improved2_aco")
    os.makedirs(output_dir, exist_ok=True)

    df_customers, depots = get_phase1_results()
    num_depots = len(depots)

    print("\n>>> Loading phase-2 data (vehicles & time windows)...")
    vehicles_data = load_vehicle_data(num_depots)
    tw_data = load_timewindow_data(df_customers, num_depots)

    total_system_cost = 0.0
    all_routes_records = []

    print("\n" + "=" * 60)
    print("Start Phase 2: Improved2 ACO for decomposed VRPTW")
    print("=" * 60)

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

        aco = ACO_Improved2(
            depot_coord=depot_coord,
            customers_data=cluster_data,
            vehicle_list=v_list,
            time_windows=tw_data["customers"],
            depot_tw=depot_tw,
            original_ids=original_ids,
        )

        best_routes, best_cost, curve = aco.run()
        print(f"  Convergence tail: {curve[-5:]}")

        if best_routes is None:
            print("  [Failed] No feasible solution found.")
            continue

        print(f"  [Done] Best cost: {best_cost:.2f}")
        total_system_cost += best_cost
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

        filename = f"improved2_route_depot_{depot_idx + 1}.png"
        plot_vehicle_routes(
            depot_coord,
            cluster_data,
            simple_routes_for_plot,
            title=f"Improved2 ACO - Depot {depot_idx + 1} (Cost: {best_cost:.1f})",
            filename=f"improved2_aco/{filename}",
        )

    print("\n" + "=" * 60)
    print(f"Total system cost: {total_system_cost:.2f}")

    if all_routes_records:
        csv_file = os.path.join(output_dir, "improved2_final_routes.csv")
        pd.DataFrame(all_routes_records).to_csv(csv_file, index=False)
        print(f"[Saved] final route file: {csv_file}")

    print("=" * 60)


if __name__ == "__main__":
    main()
