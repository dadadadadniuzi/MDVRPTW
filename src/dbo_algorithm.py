# src/dbo_algorithm.py
# 蜣螂算法实现核心
import numpy as np


class DBO:
    def __init__(self, pop_size, dim, lb, ub, max_iter, obj_func, initial_guess=None):
        """
        :param initial_guess: (可选) 一维数组，用于初始化种群中的第一个个体
        """
        self.pop_size = pop_size
        self.dim = dim
        self.lb = lb
        self.ub = ub
        self.max_iter = max_iter
        self.obj_func = obj_func

        # 1. 随机初始化种群
        self.pop = np.random.uniform(lb, ub, (pop_size, dim))

        # 2. 【核心修改】如果有初始猜测(K-means结果)，注入到种群第一个位置
        if initial_guess is not None:
            # 确保 initial_guess 在 lb 和 ub 之间
            guess = np.clip(initial_guess, lb, ub)
            self.pop[0] = guess
            # 也可以多注入几个，在初始值附近微调，增加搜索密度
            # self.pop[1] = guess + np.random.normal(0, 0.05, dim)

        self.fitness = np.zeros(pop_size)

        self.gbest_pos = np.zeros(dim)
        self.gbest_fit = float('inf')
        self.loss_curve = []

    def optimize(self):
        # 计算初始适应度
        for i in range(self.pop_size):
            self.fitness[i] = self.obj_func(self.pop[i])
            if self.fitness[i] < self.gbest_fit:
                self.gbest_fit = self.fitness[i]
                self.gbest_pos = self.pop[i].copy()

        # 定义各类蜣螂数量
        n_ball = int(self.pop_size * 0.2)
        n_brood = int(self.pop_size * 0.4)
        n_small = int(self.pop_size * 0.2)

        for t in range(self.max_iter):
            # --- 算法核心逻辑 (滚球、育雏、偷窃) ---
            # (此处代码与之前相同，为节省篇幅省略，保持原有逻辑即可)
            # ... (请保留之前的 optimize 内部完整逻辑) ...

            # 为了完整性，这里简写核心更新部分，你需要确保之前的逻辑在里面
            # 1. Ball Rolling
            for i in range(n_ball):
                # ... (同前)
                r1 = np.random.random()
                if r1 < 0.9:
                    self.pop[i] += 0.3 * (self.gbest_pos - self.pop[i])  # 简化示意
                else:
                    self.pop[i] += np.random.normal(0, 1, self.dim)
                self.pop[i] = np.clip(self.pop[i], self.lb, self.ub)

            # 2. Brood / 3. Small / 4. Thief (同前...)
            # ... (必须保留完整的DBO逻辑) ...

            # 每次迭代结束，更新适应度
            for i in range(self.pop_size):
                fit = self.obj_func(self.pop[i])
                self.fitness[i] = fit
                if fit < self.gbest_fit:
                    self.gbest_fit = fit
                    self.gbest_pos = self.pop[i].copy()

            self.loss_curve.append(self.gbest_fit)
            # 打印日志
            if (t + 1) % 10 == 0:
                print(f"Iteration {t + 1}/{self.max_iter}, Best Fitness: {self.gbest_fit:.2f}")

        return self.gbest_pos, self.loss_curve