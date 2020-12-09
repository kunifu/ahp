import pandas as pd
import numpy as np
from scipy.stats.mstats import gmean


def score2num(score):

    if '絶対的' in score: 

        num = 9

    elif 'かなり' in score: 

        num = 7

    elif 'やや' in score: 

        num = 3

    elif '同じ' in score: 

        num = 1

    else:

        num = 5

    return num


def calc_priority(mat):

    g_means = gmean(mat, axis=1)
    g_mean_sum = g_means.sum()
    prios = [g_mean/g_mean_sum for g_mean in g_means]

    return prios


if __name__ == "__main__":

    df = pd.read_excel('ahp.xlsx', sheet_name='基準')

    stds = set(df['基準a'])
    stds = list(stds.union(set(df['基準b'])))
    std_mat = np.ones((len(stds), len(stds)))

    for _, row in df.iterrows():

        std_a, std_b = row[['基準a', '基準b']]

        score = row[~row.isnull()].keys().drop(['基準a', '基準b'])
        score = score[0] if len(score) == 1 else print('Error')

        num = score2num(score)
        idx_a = stds.index(std_a)
        idx_b = stds.index(std_b)

        if '左' in score:

            std_mat[idx_a][idx_b] = num
            std_mat[idx_b][idx_a] = 1/num

            # print(f'{std_a}: {num}, {std_b}: 1/{num}')

        elif '右' in score:

            std_mat[idx_a][idx_b] = 1/num
            std_mat[idx_b][idx_a] = num

            # print(f'{std_a}: 1/{num}, {std_b}: {num}')

    std_prios = np.array(calc_priority(std_mat))

    obj_prios_mat = list()

    for std in stds:

        df = pd.read_excel('ahp.xlsx', sheet_name=std)

        objs = set(df['対象a'])
        objs = list(objs.union(set(df['対象b'])))
        obj_mat = np.ones((len(objs), len(objs)))

        for _, row in df.iterrows():

            obj_a, obj_b = row[['対象a', '対象b']]

            score = row[~row.isnull()].keys().drop(['対象a', '対象b'])
            score = score[0] if len(score) == 1 else print('Error')

            num = score2num(score)
            idx_a = objs.index(obj_a)
            idx_b = objs.index(obj_b)

            if '左' in score:

                obj_mat[idx_a][idx_b] = num
                obj_mat[idx_b][idx_a] = 1/num

                # print(f'{obj_a}: {num}, {obj_b}: 1/{num}')

            elif '右' in score:

                obj_mat[idx_a][idx_b] = 1/num
                obj_mat[idx_b][idx_a] = num

                # print(f'{obj_a}: 1/{num}, {obj_b}: {num}')

        obj_prios_mat.append(calc_priority(obj_mat))

    obj_prios_mat = np.array(obj_prios_mat)

    df_prio = pd.DataFrame(np.dot(std_prios, obj_prios_mat), index=objs)
    df_prio = df_prio.sort_values(0, ascending=False)
    print(df_prio.index.to_list())
    print(list(df_prio.values.reshape(df_prio.values.shape[0])))