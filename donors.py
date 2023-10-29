import os

import numpy as np
import pandas as pd
import xlsxwriter
from pathlib import Path
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

# Директория, где лежат рабочие файлы
os.chdir(r"C:\Users\AVShestakov\Доноры")

# Загружаем файлы в переменные `file`
file = 'RRU1800_old.xlsx'
file2 = 'RRU1800_new.xlsx'

# Загрузить лист в DataFrame по его имени: df
df_RRU_old_hua = pd.read_excel(file, sheet_name='hua')
df_RRU_old_eri = pd.read_excel(file, sheet_name='eri')

df_RRU_new_hua = pd.read_excel(file2, sheet_name='hua')
df_RRU_new_eri = pd.read_excel(file2, sheet_name='eri')

result_hua = (df_RRU_old_hua.merge(df_RRU_new_hua,
                                  on='siteid',
                                  how='outer',
                                  suffixes=['', '_new'],
                                  indicator=True))

result_hua['delta_RRU1800'] = result_hua['sum_RRU_1800_new'] - result_hua['sum_RRU_1800']

result_eri = (df_RRU_old_eri.merge(df_RRU_new_eri,
                                  on='siteid',
                                  how='outer',
                                  suffixes=['', '_new'],
                                  indicator=True))
result_eri['delta_RRU1800'] = result_eri['cells_1800_new'] - result_eri['cells_1800']

writer = pd.ExcelWriter('donors.xlsx', engine='xlsxwriter')

# Записываем DataFrame в файл
result_hua.to_excel(writer, 'hua', index=False)
result_eri.to_excel(writer, 'eri', index=False)
writer.close()

