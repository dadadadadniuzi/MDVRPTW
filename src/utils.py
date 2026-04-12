# src/utils.py
# 数据处理工具
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.cm as cm
import os
from . import config

# ==========================================
# 设置 Matplotlib 中文字体
# ==========================================
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签 (Windows系统通用)
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号

# --- 1. 随机数据生成 ---
def generate_mock_data(num_customers, map_size):
    np.random.seed(42)
    data = {
        'id': range(1, num_customers + 1),
        'x': np.random.uniform(5, map_size - 5, num_customers),
        'y': np.random.uniform(5, map_size - 5, num_customers),
        'demand': np.random.randint(1, 25, num_customers)
    }
    return pd.DataFrame(data)


# --- 2. Excel 数据加载 ---
def load_data_from_excel(filepath):
    """
    读取 Excel 文件
    返回:
    1. df_customers: 客户数据
    2. depots_coords: 中心坐标数组
    3. capacities:   中心容量数组
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"找不到文件: {filepath}")

    print(f"正在读取数据文件: {filepath} ...")

    # 1. 读取客户
    try:
        df_customers = pd.read_excel(filepath, sheet_name='Customers')
        df_customers.columns = df_customers.columns.str.lower().str.strip()  # 去除空格并转小写
    except Exception as e:
        raise ValueError(f"读取 Customers 表失败: {e}")

    # 2. 读取配送中心
    try:
        df_depots = pd.read_excel(filepath, sheet_name='Depots')
        df_depots.columns = df_depots.columns.str.lower().str.strip()

        depots_coords = df_depots[['x', 'y']].values

        # --- 核心修改：读取独立容量 ---
        if 'capacity' in df_depots.columns:
            capacities = df_depots['capacity'].values.astype(float)  # 转为数组
        else:
            print("警告: Depots 表中未找到 'capacity' 列，将使用 config.DEFAULT_CAPACITY")
            # 如果没填，创建一个长度等于仓库数的数组，填默认值
            capacities = np.full(len(df_depots), config.DEFAULT_CAPACITY)

    except Exception as e:
        raise ValueError(f"读取 Depots 表失败: {e}")

    return df_customers, depots_coords, capacities


def load_vehicle_data(num_depots):
    """
    加载异构车辆数据
    返回: dict {depot_idx: [VehicleType1, VehicleType2, ...]}
    VehicleType = {'type_id': 1, 'count': 2, 'capacity': 100, ...}
    """
    vehicles = {}

    if config.USE_CUSTOM_VEHICLES and os.path.exists(config.CUSTOM_DATA_FILE):
        try:
            df = pd.read_excel(config.CUSTOM_DATA_FILE, sheet_name='Vehicles')
            print(">>> 成功加载 Excel 车辆数据 (异构车队)")
            for i in range(num_depots):
                # 假设 Excel 中 depot_id 从 1 开始
                depot_df = df[df['depot_id'] == (i + 1)]
                if len(depot_df) == 0:
                    print(f"警告: Depot {i + 1} 没有车辆配置，使用默认随机。")
                    vehicles[i] = _generate_random_vehicles()
                else:
                    v_list = []
                    for _, row in depot_df.iterrows():
                        v_list.append({
                            'type_id': row.get('type_id', 0),
                            'count': int(row['count']),
                            'capacity': float(row['capacity']),
                            'velocity': float(row['velocity']),
                            'cost': float(row.get('var_cost', 1.0))
                        })
                    vehicles[i] = v_list
            return vehicles
        except Exception as e:
            print(f"读取 Vehicles 表失败 ({e})，转为随机。")

    # 随机生成
    print(">>> 使用随机异构车辆配置")
    for i in range(num_depots):
        vehicles[i] = _generate_random_vehicles()
    return vehicles


def _generate_random_vehicles():
    """辅助函数：随机生成一个仓库的车队"""
    v_list = []
    # 从配置的 3 种车型中随机选
    type_id = 1
    for v_conf in config.RANDOM_VEHICLE_TYPES:
        # 随机数量
        count = np.random.randint(v_conf['count'][0], v_conf['count'][1] + 1)
        if count > 0:
            v_list.append({
                'type_id': type_id,
                'count': count,
                'capacity': v_conf['cap'],
                'velocity': v_conf['speed'],
                'cost': v_conf['cost']
            })
            type_id += 1
    # 兜底：如果一种车都没随到，至少给辆小车
    if not v_list:
        v_list.append({'type_id': 1, 'count': 3, 'capacity': 100, 'velocity': 60, 'cost': 1.0})
    return v_list


def load_timewindow_data(customers_df, num_depots):
    """
    加载复杂时间窗
    返回: dict
      'depots': {depot_idx: {'start': 8, 'end': 20}},
      'customers': {cust_id: {'start': 9, 'end': 11, 'service': 0.5}}
    """
    tw_data = {'depots': {}, 'customers': {}}

    # 1. Excel 读取模式
    if config.USE_CUSTOM_TIMEWINDOWS and os.path.exists(config.CUSTOM_DATA_FILE):
        try:
            df = pd.read_excel(config.CUSTOM_DATA_FILE, sheet_name='TimeWindows')
            print(">>> 成功加载 Excel 时间窗数据")

            # 读取仓库时间
            depot_rows = df[df['category'] == 'depot']
            for _, row in depot_rows.iterrows():
                did = int(row['id']) - 1  # 假设 Excel ID 1..N -> 0..N-1
                if 0 <= did < num_depots:
                    tw_data['depots'][did] = {
                        'start': float(row['start_time']),
                        'end': float(row['end_time'])
                    }

            # 读取客户时间
            cust_rows = df[df['category'] == 'customer']
            for _, row in cust_rows.iterrows():
                cid = int(row['id'])
                tw_data['customers'][cid] = {
                    'start': float(row['start_time']),
                    'end': float(row['end_time']),
                    'service': float(row['service_time'])
                }
            return tw_data
        except Exception as e:
            print(f"读取 TimeWindows 表失败 ({e})，转为随机。")

        # 2. 随机生成模式 (修改这里)
    if not config.USE_CUSTOM_TIMEWINDOWS:  # 如果没开启Excel或读取失败
        print(">>> 使用随机生成时间窗 (智能匹配模式)")

        # (A) 生成仓库时间窗
        for i in range(num_depots):
            pattern = np.random.choice(config.RANDOM_DEPOT_TW_PATTERNS)
            tw_data['depots'][i] = {
                'start': pattern['start'],
                'end': pattern['end'],
                'name': pattern['name']
            }
            print(f"  Depot {i + 1}: {pattern['name']} ({pattern['start']}-{pattern['end']})")

        # (B) 生成客户时间窗 (【核心修改】)
        # 必须根据客户所属的 Depot (label) 来生成

        for _, row in customers_df.iterrows():
            cid = int(row['id'])
            label = int(row.get('label', 0))  # 获取 Phase 1 分配的仓库

            # 获取该仓库的营业时间
            if label in tw_data['depots']:
                d_start = tw_data['depots'][label]['start']
                d_end = tw_data['depots'][label]['end']
            else:
                # 兜底：如果是全随机分配没 label，就用早6晚10
                d_start, d_end = 6.0, 22.0

            # 确保客户时间窗在仓库营业期内
            # 留出一点回程缓冲 (比如最后1小时不接客)
            valid_end_limit = d_end - 1.0
            valid_start_limit = d_start

            if valid_end_limit <= valid_start_limit + 2:
                # 仓库时间太短了(极少见)，强制全天
                valid_start_limit, valid_end_limit = 6.0, 20.0

            # 随机生成
            start = np.random.uniform(valid_start_limit, valid_end_limit - 2)
            duration = np.random.uniform(2, 4)
            end = min(start + duration, valid_end_limit)

            tw_data['customers'][cid] = {
                'start': round(start, 2),
                'end': round(end, 2),
                'service': 0.5
            }

    return tw_data


# --- 4. 其他辅助函数  ---
def euclidean_distance(p1, p2):
    return np.sqrt(np.sum((p1 - p2) ** 2))


def save_results_to_txt(filename, method_name, centers, labels, customers_df, total_sse):
    if not os.path.exists('results'):
        os.makedirs('results')
    filepath = os.path.join('results', filename)

    with open(filepath, 'a', encoding='utf-8') as f:
        f.write(f"========== {method_name} 计算结果 ==========\n")
        for i in range(len(centers)):
            cluster_indices = np.where(labels == i)[0]
            cluster_customers = customers_df.iloc[cluster_indices]
            customer_ids = cluster_customers['id'].tolist()
            customer_ids.sort()

            current_center = centers[i]
            dist_sum = 0
            for _, row in cluster_customers.iterrows():
                cust_coord = np.array([row['x'], row['y']])
                dist = euclidean_distance(current_center, cust_coord)
                dist_sum += dist

            # Depots 命名
            f.write(f"（{i + 1}）配送中心 ID-{i + 1}，包含 {len(customer_ids)} 个目标客户，")
            f.write(f"其编号分别为：{', '.join(map(str, customer_ids))}。\n")
            f.write(f"      中心坐标: ({current_center[0]:.1f}, {current_center[1]:.1f})\n")
            f.write(f"      距离总和: {dist_sum:.2f} km\n\n")

        f.write(f"{method_name} 总距离: {total_sse:.2f}\n")
        f.write("------------------------------------------------------\n\n")


def visualize_and_save_allocation(customers, centroids, labels, title, filename):
    if not os.path.exists('results'):
        os.makedirs('results')

    plt.figure(figsize=(10, 8))
    num_clusters = len(centroids)
    colors = cm.rainbow(np.linspace(0, 1, num_clusters))

    for i in range(num_clusters):
        cluster_points = customers[labels == i]
        center = centroids[i]
        color = colors[i]

        plt.scatter(cluster_points[:, 0], cluster_points[:, 1], c=[color], s=40, alpha=0.8, label=f'Depot {i + 1}')
        plt.scatter(center[0], center[1], c=[color], marker='*', s=400, edgecolors='black', linewidths=1.5, zorder=10)
        plt.text(center[0], center[1] + 1, f'D-{i + 1}', fontsize=12, fontweight='bold', ha='center', color='black')

        for point in cluster_points:
            plt.plot([point[0], center[0]], [point[1], center[1]], c=color, alpha=0.3, linewidth=1.0)

    plt.title(title, fontsize=14)
    plt.xlabel('X (km)')
    plt.ylabel('Y (km)')
    plt.grid(True, linestyle='--', alpha=0.5)
    plt.savefig(os.path.join('results', filename), dpi=300)
    plt.close()


def plot_vehicle_routes(depot_coord, customers, routes, title, filename=None):
    """
    画出单个配送中心的车辆路径，并保存
    :param filename: 如果不为 None，则保存图片到该路径
    """
    plt.figure(figsize=(10, 8))
    # 画客户 (只取 x, y)
    plt.scatter(customers[:, 0], customers[:, 1], c='blue', s=40, label='Customers', zorder=5)

    # 标客户ID (可选，为了看清路径)
    # for i in range(len(customers)):
    #     plt.text(customers[i, 0], customers[i, 1]+0.5, str(i+1), fontsize=8)

    # 画仓库
    plt.scatter(depot_coord[0], depot_coord[1], c='red', marker='s', s=300, label='Depot', zorder=10,
                edgecolors='black')

    # 画路径
    colors = plt.cm.rainbow(np.linspace(0, 1, len(routes)))

    for route, color in zip(routes, colors):
        path_coords = []
        path_coords.append(np.array(depot_coord))

        for c_idx in route:
            cust_point = customers[c_idx][:2]
            path_coords.append(cust_point)

        path_coords.append(np.array(depot_coord))
        path_coords = np.array(path_coords)

        plt.plot(path_coords[:, 0], path_coords[:, 1], c=color, linewidth=2.0, alpha=0.8)

        # 标箭头
        if len(path_coords) > 1:
            mid_idx = len(path_coords) // 2
            p1 = path_coords[mid_idx - 1]
            p2 = path_coords[mid_idx]
            plt.arrow(p1[0], p1[1], (p2[0] - p1[0]) * 0.6, (p2[1] - p1[1]) * 0.6,
                      head_width=1.5, head_length=1.5, fc=color, ec=color, length_includes_head=True, zorder=8)

    plt.title(title, fontsize=14)
    plt.xlabel('X (km)')
    plt.ylabel('Y (km)')
    plt.grid(True, linestyle='--', alpha=0.5)
    plt.legend()

    if filename:
        save_path = os.path.join('results', filename)
        plt.savefig(save_path, dpi=300)
        print(f"  [图片保存] {save_path}")

    # plt.show() # 如果不需要弹出窗口，可以注释掉
    plt.close()  # 关闭画布释放内存