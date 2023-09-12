import os

import numpy as np
import pandas as pd
import xlsxwriter
from pathlib import Path
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

# Директория, где лежат рабочие файлы
os.chdir(r"C:\Users\AVShestakov\HLD_sectorKey\Юг")

# Загружаем файлы в переменные `file`
file = 'BSSI.xlsx'
file2 = 'Task.xlsx'

# Загрузить лист в DataFrame по его имени: df
df_BSSI = pd.read_excel(file, sheet_name='Sheet1').dropna(subset=['Sector_key'])  # BSSI
df_Task = pd.read_excel(file2, sheet_name='Sheet1')  # ТЗ

# BSSI выделяем 4G
df_BSSI_4G = df_BSSI[df_BSSI['Стандарт'].str.contains('4G')]

# Указываем writer библиотеки
writer_df_BSSI_4G = pd.ExcelWriter('BSSI_4G.xlsx', engine='xlsxwriter')

# Записываем DataFrame в файл
df_BSSI_4G.to_excel(writer_df_BSSI_4G, 'Sheet1', index=False)
writer_df_BSSI_4G.close()
# df_Task_4G = df_Task[['Индекс для BSSi', '% Покрытия 4G в полигоне', '% Покрытия 4G в полигоне после']]


# ищем соответствие номеров позиции, все ли совпадают
result_4G = (df_BSSI_4G.merge(df_Task,
                              on='Sector_key',
                              how='outer',
                              suffixes=['', '_new'],
                              indicator=True))

# для проверки соответсвия выделим отдельный df
df_BSSI_check = df_BSSI_4G[
    ['Стандарт', 'Филиал', 'Номер позиции', 'Band', 'LTE BW, MHz', 'Σ eNodeB cummulative', 'Worktype',
     'Sector_key']].copy()


def newsite1800(row):
    if ((row['Band'] == 1800) & (row['LTE BW, MHz'] != '4T4R') & (row['Σ eNodeB cummulative'] > 0) &
            (row['Worktype'].lower() == 'строительство')):
        val = 1
    else:
        val = 0
    return val


def entrance1800(row):
    if ((row['Band'] == 1800) & (row['LTE BW, MHz'] != '4T4R') & (row['Σ eNodeB cummulative'] > 0) &
            (row['Worktype'].lower() == 'модернизация')):
        val = 1
    else:
        val = 0
    return val


def entrance2100(row):
    if ((row['Band'] == 2100) & (row['LTE BW, MHz'] != '4T4R') & (row['Σ eNodeB cummulative'] > 0) &
            (row['Worktype'].lower() == 'модернизация')):
        val = 1
    else:
        val = 0
    return val


def entrance2600(row):
    if ((row['Band'] == 2600) & (row['LTE BW, MHz'] != '4T4R') & (row['Σ eNodeB cummulative'] > 0) &
            (row['Worktype'].lower() == 'модернизация')):
        val = 1
    else:
        val = 0
    return val


def entrance2600TDD(row):
    if ((row['Band'] == '2600 TDD') & (row['LTE BW, MHz'] != '4T4R') & (row['Σ eNodeB cummulative'] > 0) &
            (row['Worktype'].lower() == 'модернизация')):
        val = 1
    else:
        val = 0
    return val


def MIMO(row):
    if ((row['LTE BW, MHz'] == '4T4R') & ((row['Σ eNodeB cummulative'] == 0) | pd.isnull(row['Σ eNodeB cummulative'])) &
            (row['Worktype'].lower() == 'расширение')):
        val = 1
    else:
        val = 0
    return val


def bisector(row):
    if ((row['LTE BW, MHz'] != '4T4R') & (row['Σ eNodeB cummulative'] == 0) & (pd.isnull(row['Σ eNodeB cummulative']))
            & (row['Worktype'].lower() == 'расширение')):
        val = 1
    else:
        val = 0
    return val


df_BSSI_check['1800 NewSite'] = df_BSSI_check.apply(newsite1800, axis=1)
df_BSSI_check['1800 Entrance'] = df_BSSI_check.apply(entrance1800, axis=1)
df_BSSI_check['2100 Entrance'] = df_BSSI_check.apply(entrance2100, axis=1)
df_BSSI_check['2600 Entrance'] = df_BSSI_check.apply(entrance2600, axis=1)
df_BSSI_check['2600TDD Entrance'] = df_BSSI_check.apply(entrance2600TDD, axis=1)
df_BSSI_check['BiSector'] = df_BSSI_check.apply(bisector, axis=1)
df_BSSI_check['MIMO'] = df_BSSI_check.apply(MIMO, axis=1)

# группируем ТР по Sector_key
df_BSSI_check2 = df_BSSI_check[['Sector_key', '1800 NewSite', '1800 Entrance', '2100 Entrance', '2600 Entrance',
                                '2600TDD Entrance', 'BiSector', 'MIMO']]
df_BSSI_check3 = df_BSSI_check2.groupby(['Sector_key'])[
    ['1800 NewSite', '1800 Entrance', '2100 Entrance', '2600 Entrance',
     '2600TDD Entrance', 'BiSector', 'MIMO']].sum().reset_index()

# представляем MIMO и 1800 NewSite для BSSI в виде 1, не суммируем его по sector_key
df_BSSI_check3.loc[df_BSSI_check3['MIMO'] > 1, 'MIMO'] = 1
df_BSSI_check3.loc[df_BSSI_check3['1800 NewSite'] > 1, '1800 NewSite'] = 1
writer_df_BSSI_check3 = pd.ExcelWriter('BSSI_check.xlsx', engine='xlsxwriter')

# Записываем DataFrame в файл
df_BSSI_check3.to_excel(writer_df_BSSI_check3, 'Sheet1', index=False)
writer_df_BSSI_check3.close()

result = (df_BSSI_check3.merge(df_Task,
                               on='Sector_key',
                               how='inner',
                               suffixes=['', '_new'],
                               indicator=True))

# колонки для передвижения в начало
cols_to_move = ['Стандарт', 'Филиал', 'Номер позиции']

# перемещение колонок в начало
result = result[cols_to_move + [x for x in result.columns if x not in cols_to_move]]

result['Sector_key'] = result['Sector_key'].str.slice(0,10)
print(result['Sector_key'])


# подкрашиваем несовпадающие ТР колонок
def highlight(df):
    ret = pd.DataFrame(None, index=df.index, columns=df.columns)
    # задаем основные правила закраски в виде словаря, где ключ - наименование столбца, который нужно красить,
    # а значение - булев массив, какие ячейки красить (пока без конкретного цвета)
    masks = {'newsite': (df['1800 NewSite'] != df['newsite']) & (df['newsite'].notnull()) | ((df['1800 NewSite'] > 0) &
                                                                                             (df['newsite'].isnull())),
             'Add 1800': (df['1800 Entrance'] != df['Add 1800']) & (df['Add 1800'].notnull()) | (
                     (df['1800 Entrance'] > 0) &
                     (df['Add 1800'].isnull())),
             'Add 2100': (df['2100 Entrance'] != df['Add 2100']) & (df['Add 2100'].notnull()) | (
                     (df['2100 Entrance'] > 0) &
                     (df['Add 2100'].isnull())),
             'Add 2600': (df['2600 Entrance'] != df['Add 2600']) & (df['Add 2600'].notnull()) | (
                     (df['2600 Entrance'] > 0) &
                     (df['Add 2600'].isnull())),
             'Add 2600TDD': (df['2600TDD Entrance'] != df['Add 2600TDD']) & (df['Add 2600TDD'].notnull()) | (
                     (df['2600TDD Entrance'] > 0) &
                     (df['Add 2600TDD'].isnull())),
             'Add BiSector': (df['BiSector'] != df['Add BiSector']) & (df['Add BiSector'].notnull()) | (
                     (df['BiSector'] > 0) &
                     (df['Add BiSector'].isnull())),
             'MIMO 4T4R': (df['MIMO'] != df['MIMO 4T4R']) & (df['MIMO 4T4R'].notnull()) | (
                     (df['MIMO'] > 0) &
                     (df['MIMO 4T4R'].isnull())),
             }
    red, yellow = df['Проведена смена ТР'].ne(1), df['Проведена смена ТР'].eq(1)  # делаем булевы массивы для
    # красного и желтого цветов
    # перебираем словарь с правилами закраски
    for column, mask in masks.items():
        ret.loc[
            mask & red, column] = 'background-color: red'  # дополняем булев массив (маску) условием закраски в
        # красный и красим
        ret.loc[
            mask & yellow, column] = 'background-color: yellow'  # дополняем булев массив (маску) условием закраски в
        # желтый и красим

    return ret


result2 = result.style.apply(highlight, axis=None)

writer_result2 = pd.ExcelWriter('consistency2.xlsx', engine='xlsxwriter')

# Записываем DataFrame в файл
result2.to_excel(writer_result2, 'Sheet1', index=False)
writer_result2.close()

# 1800 for new site
df_BSSI_checkNewSite = df_BSSI_check[
    (df_BSSI_check['Band'] == 1800) & (df_BSSI_check['LTE BW, MHz'] != '4T4R')
    & (df_BSSI_check['Σ eNodeB cummulative'] > 0) & (df_BSSI_check['Worktype'].str.contains('оител'))]

# writer_df_BSSI_NewSite = pd.ExcelWriter('result_newsite.xlsx', engine='xlsxwriter')
# # write each DataFrame to a specific sheet
# df_BSSI_NewSite.to_excel(writer_df_BSSI_NewSite, 'NewSite', index=False)
# writer_df_BSSI_NewSite.close()

df_Task_NewSite = df_Task[['Стандарт', 'Филиал', 'Номер позиции', 'newsite', 'Бэнды New Site', 'Sector_key',
                           'Проведена смена ТР']]

df_Task_NewSite = df_Task_NewSite[(df_Task_NewSite['newsite'] > 0) | ((df_Task_NewSite['newsite'].isnull()) &
                                                                      (df_Task_NewSite['Проведена смена ТР'] > 0)) | (
                                          (df_Task_NewSite['newsite'] == 0) &
                                          (df_Task_NewSite['Проведена смена ТР'] > 0))]

result_1800NewSite = (df_BSSI_checkNewSite.merge(df_Task_NewSite,
                                                 on='Sector_key',
                                                 how='outer',
                                                 suffixes=['', '_new'],
                                                 indicator=True))

result_1800NewSite = result_1800NewSite[(result_1800NewSite['Проведена смена ТР'].isnull()) &
                                        ((result_1800NewSite['_merge'] == 'left_only') |
                                         (result_1800NewSite['_merge'] == 'right_only'))]

result_1800NewSite = result_1800NewSite.replace(['left_only', 'right_only'], ['отсутствует в ТЗ', 'отсутствует в BSSI'])

# 1800 entrance
df_BSSI_check1800Entrance = df_BSSI_check[
    (df_BSSI_check['Band'] == 1800) & (df_BSSI_check['LTE BW, MHz'] != '4T4R')
    & (df_BSSI_check['Σ eNodeB cummulative'] > 0) & (df_BSSI_check['Worktype'].str.contains('дернизац'))]

writer_df_BSSI_check1800Entrance = pd.ExcelWriter('BSSI_entrance1800.xlsx', engine='xlsxwriter')
# write each DataFrame to a specific sheet
df_BSSI_check1800Entrance.to_excel(writer_df_BSSI_check1800Entrance, 'entrance', index=False)
writer_df_BSSI_check1800Entrance.close()

df_Task_1800Entrance = df_Task[['Стандарт', 'Филиал', 'Номер позиции', 'Add 1800', 'Sector_key', 'Проведена смена ТР']]

df_Task_1800Entrance = df_Task_1800Entrance[
    (df_Task_1800Entrance['Add 1800'] > 0) | ((df_Task_1800Entrance['Add 1800'].isnull()) &
                                              (df_Task_1800Entrance['Проведена смена ТР'] > 0)) | (
            (df_Task_1800Entrance['Add 1800'] == 0) &
            (df_Task_1800Entrance['Проведена смена ТР'] > 0))]

writer_df_Task_1800Entrance = pd.ExcelWriter('Task_entrance1800.xlsx', engine='xlsxwriter')
# write each DataFrame to a specific sheet
df_Task_1800Entrance.to_excel(writer_df_Task_1800Entrance, 'entrance', index=False)
writer_df_Task_1800Entrance.close()

result_1800Entrance = (df_BSSI_check1800Entrance.merge(df_Task_1800Entrance,
                                                       on='Sector_key',
                                                       how='outer',
                                                       suffixes=['', '_new'],
                                                       indicator=True))

result_1800Entrance = result_1800Entrance[(result_1800Entrance['Проведена смена ТР'].isnull()) &
                                          ((result_1800Entrance['_merge'] == 'left_only') |
                                           (result_1800Entrance['_merge'] == 'right_only'))]

result_1800Entrance = result_1800Entrance.replace(['left_only', 'right_only'],
                                                  ['отсутствует в ТЗ', 'отсутствует в BSSI'])

# 2100 entrance
df_BSSI_2100Entrance = df_BSSI_check[
    (df_BSSI_check['Band'] == 2100) & (df_BSSI_check['LTE BW, MHz'] != '4T4R')
    & (df_BSSI_check['Σ eNodeB cummulative'] > 0) & (df_BSSI_check['Worktype'].str.contains('дернизац'))]

df_Task_2100Entrance = df_Task[['Стандарт', 'Филиал', 'Номер позиции', 'Add 2100', 'Sector_key', 'Проведена смена ТР']]

df_Task_2100Entrance = df_Task_2100Entrance[
    (df_Task_2100Entrance['Add 2100'] > 0) | ((df_Task_2100Entrance['Add 2100'].isnull()) &
                                              (df_Task_2100Entrance['Проведена смена ТР'] > 0)) | (
            (df_Task_2100Entrance['Add 2100'] == 0) &
            (df_Task_2100Entrance['Проведена смена ТР'] > 0))]

result_2100Entrance = (df_BSSI_2100Entrance.merge(df_Task_2100Entrance,
                                                  on='Sector_key',
                                                  how='outer',
                                                  suffixes=['', '_new'],
                                                  indicator=True))

result_2100Entrance = result_2100Entrance[(result_2100Entrance['Проведена смена ТР'].isnull()) &
                                          ((result_2100Entrance['_merge'] == 'left_only') |
                                           (result_2100Entrance['_merge'] == 'right_only'))]

result_2100Entrance = result_2100Entrance.replace(['left_only', 'right_only'],
                                                  ['отсутствует в ТЗ', 'отсутствует в BSSI'])

# 2600 entrance
df_BSSI_2600Entrance = df_BSSI_check[
    (df_BSSI_check['Band'] == 2600) & (df_BSSI_check['LTE BW, MHz'] != '4T4R')
    & (df_BSSI_check['Σ eNodeB cummulative'] > 0) & (df_BSSI_check['Worktype'].str.contains('дернизац'))]

df_Task_2600Entrance = df_Task[['Стандарт', 'Филиал', 'Номер позиции', 'Add 2600', 'Sector_key', 'Проведена смена ТР']]

df_Task_2600Entrance = df_Task_2600Entrance[
    (df_Task_2600Entrance['Add 2600'] > 0) | ((df_Task_2600Entrance['Add 2600'].isnull()) &
                                              (df_Task_2600Entrance['Проведена смена ТР'] > 0)) | (
            (df_Task_2600Entrance['Add 2600'] == 0) &
            (df_Task_2600Entrance['Проведена смена ТР'] > 0))]

result_2600Entrance = (df_BSSI_2600Entrance.merge(df_Task_2600Entrance,
                                                  on='Sector_key',
                                                  how='outer',
                                                  suffixes=['', '_new'],
                                                  indicator=True))

result_2600Entrance = result_2600Entrance[(result_2600Entrance['Проведена смена ТР'].isnull()) &
                                          ((result_2600Entrance['_merge'] == 'left_only') |
                                           (result_2600Entrance['_merge'] == 'right_only'))]

result_2600Entrance = result_2600Entrance.replace(['left_only', 'right_only'],
                                                  ['отсутствует в ТЗ', 'отсутствует в BSSI'])
# 2600TDD entrance
df_BSSI_2600TDD_Entrance = df_BSSI_check[
    (df_BSSI_check['Band'] == '2600 TDD') & (df_BSSI_check['LTE BW, MHz'] != '4T4R')
    & (df_BSSI_check['Σ eNodeB cummulative'] > 0) & (df_BSSI_check['Worktype'].str.contains('дернизац'))]

df_Task_2600TDD_Entrance = df_Task[['Стандарт', 'Филиал', 'Номер позиции', 'Add 2600TDD', 'Sector_key',
                                    'Проведена смена ТР']]

df_Task_2600TDD_Entrance = df_Task_2600TDD_Entrance[
    (df_Task_2600TDD_Entrance['Add 2600TDD'] > 0) | ((df_Task_2600TDD_Entrance['Add 2600TDD'].isnull()) &
                                                     (df_Task_2600TDD_Entrance['Проведена смена ТР'] > 0)) | (
            (df_Task_2600TDD_Entrance['Add 2600TDD'] == 0) &
            (df_Task_2600TDD_Entrance['Проведена смена ТР'] > 0))]

result_2600TDD_Entrance = (df_BSSI_2600TDD_Entrance.merge(df_Task_2600TDD_Entrance,
                                                          on='Sector_key',
                                                          how='outer',
                                                          suffixes=['', '_new'],
                                                          indicator=True))

result_2600TDD_Entrance = result_2600TDD_Entrance[(result_2600TDD_Entrance['Проведена смена ТР'].isnull()) &
                                                  ((result_2600TDD_Entrance['_merge'] == 'left_only') |
                                                   (result_2600TDD_Entrance['_merge'] == 'right_only'))]

result_2600TDD_Entrance = result_2600TDD_Entrance.replace(['left_only', 'right_only'],
                                                          ['отсутствует в ТЗ', 'отсутствует в BSSI'])

# MIMO 4T4R
df_BSSI_MIMO = df_BSSI_check[(df_BSSI_check['LTE BW, MHz'] == '4T4R')
                             & ((df_BSSI_check['Σ eNodeB cummulative'] == 0) |
                                df_BSSI_check['Σ eNodeB cummulative'].isnull()) & (
                                 df_BSSI_check['Worktype'].str.contains('сширен'))]

df_Task_MIMO = df_Task[['Стандарт', 'Филиал', 'Номер позиции', 'MIMO 4T4R', 'Sector_key', 'Проведена смена ТР']]

df_Task_MIMO = df_Task_MIMO[(df_Task_MIMO['MIMO 4T4R'] == 1) | ((df_Task_MIMO['MIMO 4T4R'].isnull()) &
                                                                (df_Task_MIMO[
                                                                     'Проведена смена ТР'] > 0)) | (
                                    (df_Task_MIMO['MIMO 4T4R'] == 0) &
                                    (df_Task_MIMO['Проведена смена ТР'] > 0))]

result_MIMO = (df_BSSI_MIMO.merge(df_Task_MIMO,
                                  on='Sector_key',
                                  how='outer',
                                  suffixes=['', '_new'],
                                  indicator=True))

result_MIMO = result_MIMO[(result_MIMO['Проведена смена ТР'].isnull()) &
                          ((result_MIMO['_merge'] == 'left_only') |
                           (result_MIMO['_merge'] == 'right_only'))]
result_MIMO = result_MIMO.replace(['left_only', 'right_only'],
                                  ['отсутствует в ТЗ', 'отсутствует в BSSI'])
# bisector
df_BSSI_bisector = df_BSSI_check[(df_BSSI_check['LTE BW, MHz'] != '4T4R')
                                 & ((df_BSSI_check['Σ eNodeB cummulative'] == 0) |
                                    df_BSSI_check['Σ eNodeB cummulative'].isnull()) & (
                                     df_BSSI_check['Worktype'].str.contains('сширен'))]

df_Task_bisector = df_Task[['Стандарт', 'Филиал', 'Номер позиции', 'Add BiSector', 'Sector_key', 'Проведена смена ТР']]

df_Task_bisector = df_Task_bisector[(df_Task_bisector['Add BiSector'] == 1) |
                                    ((df_Task_bisector['Add BiSector'].isnull()) &
                                     (df_Task_bisector['Проведена смена ТР'] > 0)) |
                                    ((df_Task_bisector['Add BiSector'] == 0) &
                                     (df_Task_bisector['Проведена смена ТР'] > 0))]

result_bisector = (df_BSSI_bisector.merge(df_Task_bisector,
                                          on='Sector_key',
                                          how='outer',
                                          suffixes=['', '_new'],
                                          indicator=True))

result_bisector = result_bisector[(result_bisector['Проведена смена ТР'].isnull()) &
                                  ((result_bisector['_merge'] == 'left_only') |
                                   (result_bisector['_merge'] == 'right_only'))]
result_bisector = result_bisector.replace(['left_only', 'right_only'],
                                          ['отсутствует в ТЗ', 'отсутствует в BSSI'])

writer = pd.ExcelWriter('result.xlsx', engine='xlsxwriter')
# write each DataFrame to a specific sheet
result_1800NewSite.to_excel(writer, 'NewSite', index=False)
result_1800Entrance.to_excel(writer, '1800Entrance', index=False)
result_2100Entrance.to_excel(writer, '2100Entrance', index=False)
result_2600Entrance.to_excel(writer, '2600Entrance', index=False)
result_2600TDD_Entrance.to_excel(writer, '2600TDD_Entrance', index=False)
result_MIMO.to_excel(writer, 'MIMO', index=False)
result_bisector.to_excel(writer, 'bisector', index=False)

writer.close()
