# src/config.py
# 控制数据中心
import numpy as np
import os

# ===========================
# 1. 第一阶段解决多车场问题阶段控制开关(当开关为 True 时生效)
# ===========================
USE_CUSTOM_DATA = True  # True: 读取Excel数据; False: 使用下方随机生成配置

# ===========================
# 2.1 自定义数据配置
# ===========================
# Excel 文件路径 (建议放在 data 目录下)
# 确保你安装了 openpyxl: pip install openpyxl
CUSTOM_DATA_FILE = os.path.join('data', 'mydata.xlsx')

# ===========================
# 2.2 第二阶段车辆路径阶段阶段数据开关(当开关为 True 时生效)
# ===========================

# 车辆数据来源: True=Excel, False=随机
USE_CUSTOM_VEHICLES = True

# 时间窗数据来源: True=Excel, False=随机
USE_CUSTOM_TIMEWINDOWS = True


# 全局默认容量 (如果 Depots 表里没有 capacity 列，则用这个)
# 默认每个配送中心的最大送货量，用于负载平衡
DEFAULT_CAPACITY = 115

# ===========================
# 3. 第一阶段解决多车场问题随机生成配置
# ===========================
RANDOM_NUM_CUSTOMERS = 50
RANDOM_MAP_SIZE = 100

# 【这里修改随机模式下的配送中心】
# 你可以随意增加行数(改变数量)或修改坐标(改变位置)
RANDOM_DEPOTS = np.array([
    [20, 80],  # Depot 1
    [80, 80],  # Depot 2
    [50, 50],  # Depot 3
    [20, 20],  # Depot 4
    [80, 20]   # Depot 5 (例如我现在改成了5个中心)
])

# 随机模式下的独立容量 (必须与上面的行数对应)
# 例如：Depot1=70, Depot2=80, Depot3=60, Depot4=90
RANDOM_CAPACITIES = np.array([70, 80, 60, 90])

# ===========================
# 4. 第一阶段解决多车场问题相关算法参数
# ===========================
DBO_POP_SIZE = 50
DBO_MAX_ITER = 200


# 种群比例 (参考原始论文)
PERC_BALL_ROLLING = 0.2  # 滚球蜣螂比例
PERC_BROOD = 0.4         # 产卵(育雏)比例
PERC_SMALL = 0.2         # 小蜣螂比例
PERC_THIEF = 0.2         # 偷窃蜣螂比例

# ===========================
# 5. 第二阶段车辆路径算法随机模式参数 (当上面开关为 False 时生效)
# ===========================
# 随机模式下的车辆配置 (异构车队)
# 格式: [ (概率, 数量范围, 载重, 速度, 成本) ]
RANDOM_VEHICLE_TYPES = [
    {'prob': 0.5, 'count': (2, 4), 'cap': 100, 'speed': 60, 'cost': 1.0}, # 小车
    {'prob': 0.3, 'count': (1, 3), 'cap': 150, 'speed': 50, 'cost': 1.2}, # 中车
    {'prob': 0.2, 'count': (0, 2), 'cap': 200, 'speed': 40, 'cost': 1.5}  # 大车
]

# 随机模式下的仓库时间窗类型 (多种模式)
RANDOM_DEPOT_TW_PATTERNS = [
    {'name': '朝九晚五', 'start': 9.0, 'end': 17.0},
    {'name': '早班',     'start': 6.0, 'end': 18.0},
    {'name': '晚班',     'start': 13.0, 'end': 22.0},
    {'name': '全天',     'start': 0.0, 'end': 24.0}
]

# 随机客户时间窗
CUSTOMER_TW_LENGTH_MIN = 2  # 至少2小时窗口
CUSTOMER_TW_LENGTH_MAX = 5
# ===========================
# 6. 第二阶段蚁群算法 (ACO) 参数
# ===========================
ACO_ANT_COUNT = 60      # 蚂蚁数量
ACO_ITERATIONS = 400     # 迭代次数
ACO_ALPHA = 1.0         # 信息素重要程度
ACO_BETA = 2.0          # 启发函数重要程度 (距离倒数)
ACO_RHO = 0.1           # 信息素挥发率
ACO_Q = 100.0           # 信息素强度

# 软时间窗惩罚系数
# 成本 = 距离 + PENALTY * (迟到时间)
TIME_WINDOW_PENALTY = 100