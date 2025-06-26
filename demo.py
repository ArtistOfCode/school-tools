import numpy as np

if __name__ == '__main__':
    scores = np.array([85, 92, 78, 91, 88])
    ranks = np.argsort(-scores).argsort() + 1
    print(ranks)  # 输出: [2 1 5 1 3]
