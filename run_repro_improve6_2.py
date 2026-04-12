import os
import pandas as pd
import numpy as np

from src import config
from src.utils import load_data_from_excel, load_vehicle_data, load_timewindow_data
from src.aco_improved6_2_dbo import ACO_Improved6_2_DBO, tune_improve6_2_params_with_dbo


def get_phase1_results():
    csv_path = os.path.join("results", "allocation_results.csv")
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"cannot find {csv_path}")
    df = pd.read_csv(csv_path)
    if config.USE_CUSTOM_DATA:
        _, fixed_depots, _ = load_data_from_excel(config.CUSTOM_DATA_FILE)
    else:
        fixed_depots = config.RANDOM_DEPOTS
    return df, fixed_depots


def run_one_repeat(rep_idx, seed, df_customers, depots, vehicles_data, tw_data):
    total_cost = 0.0
    total_distance = 0.0
    route_count = 0
    num_depots = len(depots)

    for depot_idx in range(num_depots):
        cluster_df = df_customers[df_customers["label"] == depot_idx]
        if len(cluster_df) == 0:
            continue

        cluster_data = cluster_df[["x", "y", "demand"]].values
        original_ids = cluster_df["id"].values
        depot_coord = depots[depot_idx]
        v_list = vehicles_data[depot_idx]
        depot_tw = tw_data["depots"][depot_idx]
        problem_args = (
            depot_coord,
            cluster_data,
            v_list,
            tw_data["customers"],
            depot_tw,
            original_ids,
        )

        tune_seed = int(seed + depot_idx * 101)
        run_seed = int(seed + 2026 + depot_idx * 17)
        best_params, _ = tune_improve6_2_params_with_dbo(problem_args, seed=tune_seed)
        solver = ACO_Improved6_2_DBO(*problem_args, params=best_params)
        best_routes, best_cost, _ = solver.run(
            iterations=max(config.ACO_ITERATIONS, 550),
            ant_count=max(config.ACO_ANT_COUNT, 80),
            seed=run_seed,
        )
        if best_routes is None:
            return {
                "Method": "Improved6.2 DBO-MACS-ACO",
                "Repeat": rep_idx,
                "Seed": seed,
                "Total_Cost": float("inf"),
                "Total_Distance": float("inf"),
                "Route_Count": 0,
            }

        total_cost += float(best_cost)
        route_count += len(best_routes)
        for route in best_routes:
            total_distance += float(route["distance"])

    return {
        "Method": "Improved6.2 DBO-MACS-ACO",
        "Repeat": rep_idx,
        "Seed": seed,
        "Total_Cost": total_cost,
        "Total_Distance": total_distance,
        "Route_Count": route_count,
    }


def main():
    out_dir = os.path.join("results", "repro_compare")
    os.makedirs(out_dir, exist_ok=True)

    seeds = [101, 202, 303, 404, 505]
    df_customers, depots = get_phase1_results()
    num_depots = len(depots)
    vehicles_data = load_vehicle_data(num_depots)
    tw_data = load_timewindow_data(df_customers, num_depots)

    rows = []
    for i, seed in enumerate(seeds, start=1):
        print(f"[Repeat {i}/5] seed={seed} ...")
        row = run_one_repeat(i, seed, df_customers, depots, vehicles_data, tw_data)
        rows.append(row)
        print(
            f"  cost={row['Total_Cost']:.6f}, distance={row['Total_Distance']:.6f}, routes={row['Route_Count']}"
        )

    df_raw = pd.DataFrame(rows)
    raw_csv = os.path.join(out_dir, "improve6_2_repeats_raw.csv")
    df_raw.to_csv(raw_csv, index=False)

    grp = df_raw.groupby("Method", as_index=False).agg(
        Mean_Cost=("Total_Cost", "mean"),
        Var_Cost=("Total_Cost", "var"),
        Std_Cost=("Total_Cost", "std"),
        Mean_Distance=("Total_Distance", "mean"),
        Var_Distance=("Total_Distance", "var"),
        Std_Distance=("Total_Distance", "std"),
        Mean_Routes=("Route_Count", "mean"),
    )
    summary_csv = os.path.join(out_dir, "improve6_2_repeats_summary.csv")
    grp.to_csv(summary_csv, index=False)

    summary_txt = os.path.join(out_dir, "improve6_2_repeats_summary.txt")
    with open(summary_txt, "w", encoding="utf-8") as f:
        f.write(grp.to_string(index=False))
        f.write("\n")

    print(f"[Saved] {raw_csv}")
    print(f"[Saved] {summary_csv}")
    print(f"[Saved] {summary_txt}")


if __name__ == "__main__":
    main()
