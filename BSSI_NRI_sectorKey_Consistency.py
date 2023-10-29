# проверка соответствия Номера позиции и разгружаемых Sector_key
import os

import numpy as np
import pandas as pd
import xlsxwriter
from pathlib import Path
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

# Директория, где лежат рабочие файлы
os.chdir(r"C:\Users\AVShestakov\HLD_sectorKey")

# Загружаем файлы в переменные `file`
BSSI_data = 'BSSI_tot.xlsx'

df_BSSI = pd.read_excel(BSSI_data, sheet_name='Sheet1')

# удалить пробелы
df_BSSI = df_BSSI.replace(r'\s+', '', regex=True)
df_BSSI = df_BSSI[(df_BSSI['Стандарт'].str.contains('4G')) &
                  (df_BSSI['Worktype'].str.contains('ExtBW') == False) & (
                          df_BSSI['Worktype'].str.contains('оител') == False)
                  & (df_BSSI['Worktype'].str.contains('refarm') == False) & (df_BSSI['Ключ'].notnull())]
# весь df - в строковый тип
df_BSSI['Номер позиции'] = df_BSSI['Номер позиции'].astype(str)
df_BSSI['Ключ'] = df_BSSI['Ключ'].astype(str)

consistency = df_BSSI.apply(lambda x: x['Номер позиции'] in x['Ключ'], axis=1)
df_BSSI.insert(loc=len(df_BSSI.columns), column='Соответствие', value=consistency)

print(df_BSSI.info())


def highlight(df):
    ret = pd.DataFrame(None, index=df.index, columns=df.columns)
    # задаем основные правила закраски в виде словаря, где ключ - наименование столбца, который нужно красить,
    # а значение - булев массив, какие ячейки красить (пока без конкретного цвета)
    masks = {'Номер позиции': (df['Соответствие'] == False)}

    red = df['Соответствие'] == False  # делаем булев массив для красного цвета
    for column, mask in masks.items():
        ret.loc[
            mask & red, column] = 'background-color: red'
    return ret


df_BSSI = df_BSSI.style.apply(highlight, axis=None)

writer_df_BSSI = pd.ExcelWriter('BSSI_NRI_sectorKey_Сonsistency.xlsx', engine='xlsxwriter')
df_BSSI.to_excel(writer_df_BSSI, 'Sheet1', index=False)
writer_df_BSSI.close()

# df_BSSI[['Ключ', 'Ключ1', 'Ключ2']] = df_BSSI['Ключ'].str.split(',', expand=True)

# def highlight(df):
#     ret = pd.DataFrame(None, index=df.index, columns=df.columns)
#     # задаем основные правила закраски в виде словаря, где ключ - наименование столбца, который нужно красить,
#     # а значение - булев массив, какие ячейки красить (пока без конкретного цвета)
#     masks = {'Номер позиции': df['Ключ'].str.contains(df['Номер позиции']), }
#
#     red = df['Ключ'].str.contains(df['Номер позиции'].ne(1))  # делаем булев массив для
#     # красного цвета
#     # перебираем словарь с правилами закраски
#     for column, mask in masks.items():
#         ret.loc[
#             mask & red, column] = 'background-color: red'  # дополняем булев массив (маску) условием закраски в
#         # красный и красим
#     return ret
#
#
# df_BSSI2 = df_BSSI.style.apply(highlight, axis=None)
#
# writer_df_BSSI2 = pd.ExcelWriter('BSSI_NRI_sectorKey_Сonsistency.xlsx', engine='xlsxwriter')
# df_BSSI2.to_excel(writer_df_BSSI2, 'Sheet1', index=False)
# writer_df_BSSI2.close()
