# main_step2_route.py
import numpy as np
import pandas as pd
import os
import sys
from src import config
from src.utils import load_data_from_excel, load_vehicle_data, load_timewindow_data, plot_vehicle_routes
from src.aco_vrptw import ACO_VRPTW


def get_phase1_results():
    csv_path = os.path.join("results", "allocation_results.csv")
    if not os.path.exists(csv_path):
        print(f"错误：找不到文件 {csv_path}")
        sys.exit(1)
    print(f">>> 正在读取 Phase 1 分配结果: {csv_path}")
    df = pd.read_csv(csv_path)
    if config.USE_CUSTOM_DATA:
        _, fixed_depots, _ = load_data_from_excel(config.CUSTOM_DATA_FILE)
    else:
        fixed_depots = config.RANDOM_DEPOTS
    return df, fixed_depots


def main():
    # 1. 获取分配结果
    df_customers, depots = get_phase1_results()
    num_depots = len(depots)

    # 2. 加载 Phase 2 数据
    print("\n>>> 加载 Phase 2 数据 (车辆 & 时间窗)...")
    vehicles_data = load_vehicle_data(num_depots)
    tw_data = load_timewindow_data(df_customers, num_depots)

    total_system_cost = 0
    all_routes_records = []  # 用于保存最终 CSV

    print("\n" + "=" * 50)
    print("开始 Phase 2: 多车场 VRPTW 路径规划 (ACO)")
    print("=" * 50)

    # 3. 对每个配送中心单独求解
    for depot_idx in range(num_depots):
        print(f"\n>>> 正在优化配送中心 ID-{depot_idx + 1} ...")

        # 提取该中心的客户
        cluster_df = df_customers[df_customers['label'] == depot_idx]
        if len(cluster_df) == 0:
            print("  该中心无客户，跳过。")
            continue

        cluster_data = cluster_df[['x', 'y', 'demand']].values
        original_ids = cluster_df['id'].values
        depot_coord = depots[depot_idx]

        v_list = vehicles_data[depot_idx]
        depot_tw = tw_data['depots'][depot_idx]

        print(f"  客户数: {len(cluster_df)}")
        print(f"  仓库时间窗: {depot_tw['start']} - {depot_tw['end']}")

        # 运行 ACO
        aco = ACO_VRPTW(
            depot_coord=depot_coord,
            customers_data=cluster_data,
            vehicle_list=v_list,
            time_windows=tw_data['customers'],
            depot_tw=depot_tw,
            original_ids=original_ids
        )

        best_routes, best_cost, curve= aco.run()


        # 这里简单打印一下最后几次迭代
        print(f"  收敛过程 (最后5代): {curve[-5:]}")

        if best_routes is None:
            print("  [失败] 未找到可行解 (可能车辆不足或时间窗冲突)。")
            continue

        print(f"  [完成] 最优总成本: {best_cost:.2f}")
        total_system_cost += best_cost

        # 收集结果
        simple_routes_for_plot = []

        for i, r_info in enumerate(best_routes):
            route_indices = r_info['path']
            v_info = r_info['vehicle']
            dist = r_info['distance']

            # 转换ID
            real_ids = [original_ids[idx] for idx in route_indices]
            route_str = " -> ".join(map(str, real_ids))

            print(f"    车辆 {i + 1} (Type {v_info['type_id']}): Depot -> {route_str} -> Depot")

            # 记录到列表，准备存 CSV
            all_routes_records.append({
                'Depot_ID': depot_idx + 1,
                'Vehicle_ID': i + 1,
                'Vehicle_Type': v_info['type_id'],
                'Capacity': v_info['capacity'],
                'Route': f"0 -> {route_str} -> 0",  # 0代表仓库
                'Distance': round(dist, 2),
                'Cost': round(dist * v_info['cost'], 2)  # 假设只算距离变动成本
            })

            simple_routes_for_plot.append(route_indices)

        # 保存图片 (文件名: route_depot_1.png)
        plot_vehicle_routes(
            depot_coord, cluster_data, simple_routes_for_plot,
            title=f"Depot {depot_idx + 1} Routes (Cost: {best_cost:.1f})",
            filename=f"route_depot_{depot_idx + 1}.png"
        )

    # 4. 保存 CSV
    print("\n" + "=" * 50)
    print(f"系统总行驶成本: {total_system_cost:.2f}")

    if all_routes_records:
        csv_file = os.path.join("results", "final_routes.csv")
        df_res = pd.DataFrame(all_routes_records)
        df_res.to_csv(csv_file, index=False)
        print(f"[成功] 最终路径方案已保存至: {csv_file}")
    else:
        print("[警告] 未生成任何有效路径。")

    print("=" * 50)


if __name__ == "__main__":
    main()