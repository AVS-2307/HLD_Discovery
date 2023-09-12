import os
import pandas as pd
import xlsxwriter
from pathlib import Path
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

# Директория, где лежат рабочие файлы
os.chdir(r"C:\Users\AVShestakov\Discovery\МОбл")

# Загружаем выгрузку с beeplan
file = 'result_p455_v695_s34_2023_07_06_11_11_49.xlsx'

df_capacity = pd.read_excel(file, sheet_name='steps result')


# выбираем только те позиции, где необходима стройка от 2100 до new_site
def get_steps():
    df_res_capacity_newsite = df_capacity[(df_capacity.enhancerequired8 == 1)]
    df_res_capacity_2600T = df_capacity[(df_capacity.enhancerequired7 == 1) & (df_capacity.ischanged8 == 1)]
    df_res_capacity_2600 = df_capacity[(df_capacity.enhancerequired6 == 1) & (df_capacity.ischanged7 == 1)]
    df_res_capacity_2100 = df_capacity[(df_capacity.enhancerequired5 == 1) & (df_capacity.ischanged6 == 1)]

    df_res_capacity = pd.concat(
        [df_res_capacity_newsite, df_res_capacity_2600T, df_res_capacity_2600, df_res_capacity_2100], axis=0)

    # Указываем writer библиотеки
    # writer_df_res_capacity = pd.ExcelWriter('res_capacity.xlsx', engine='xlsxwriter')

    # Записываем DataFrame в файл
    # df_res_capacity.to_excel(writer_df_res_capacity, 'Sheet 1')
    #
    # writer_df_res_capacity.close()

    return df_res_capacity


def get_ecells():
    ecell_list = 'ECellsList_2023-07-10.xlsx'

    # ecell_list загружаем в dataframe:
    df_ecell_list = pd.read_excel(ecell_list, sheet_name='Лист1')

    return df_ecell_list


def get_cells_weak_coverage_choice():
    data_dir = Path(r"C:\Users\AVShestakov\Discovery\МОбл\Кейс63.Зоны в ГЮ с низким NPS")

    # объединяем все файлы измерений в один
    df_total = (pd.concat([pd.read_excel(f).assign(file_name=f.stem)
                           for f in data_dir.glob("*.xlsx")],
                          ignore_index=True))

    # Находим среднее число измерений
    mean = df_total['MR Count'].mean().round(0)

    # Выбираем только те секторы, где % плохих измерений > 5% и кол-во измерений >= среднему значению * 0,9
    df_choice = df_total.loc[
        ((df_total['MR Count'] >= mean * 0.9) & (df_total['DL Weak Coverage Percentage (%)'] >= 5) & (
                df_total['file_name'] == 'Weak_Coverage_total'))]
    return df_choice


df_choice = get_cells_weak_coverage_choice()
df_ecell_list = get_ecells()


# по eNodeB ID-Cell ID в Weak_Coverage_choice в EcellsList находим sector_key
def get_sector_key(df_choice: df_choice, df_ecell_list: df_ecell_list):
    df_sector_key_weak_coverage = (df_choice.merge(df_ecell_list,
                                                   on='eNodeB ID-Cell ID',
                                                   how='left',
                                                   suffixes=['', '_new'],
                                                   indicator=True))

    df_sector_key_weak_coverage = df_sector_key_weak_coverage.drop(columns=['sector_key', 'sector_key_enh'])

    #
    # writer_sector_key_weak_coverage = pd.ExcelWriter('sector_key_weak_coverage.xlsx', engine='xlsxwriter')
    # df_sector_key_weak_coverage.to_excel(writer_sector_key_weak_coverage, 'Sheet 1')
    #
    # writer_sector_key_weak_coverage.close()
    return df_sector_key_weak_coverage


df_res_capacity = get_steps()
df_sector_key_weak_coverage = get_sector_key(df_choice, df_ecell_list)


# получим секторы, где плохое покрытие и которые требуют расширения
def get_ecells_enhance_required(df_res_capacity: df_res_capacity,
                                df_sector_key_weak_coverage: df_sector_key_weak_coverage):
    df_sector_key_enh_req = (df_res_capacity.merge(df_sector_key_weak_coverage,
                                                   left_on='sector_key',
                                                   right_on='sector_key_new',
                                                   how='left',
                                                   suffixes=['', '_new'],
                                                   indicator='exists'))

    df_sector_key_enh_req = df_sector_key_enh_req.loc[((df_sector_key_enh_req['exists']) == 'both')]

    df_sector_key_enh_req_newSite = df_sector_key_enh_req[(df_sector_key_enh_req.enhancerequired8 == 1)]
    df_sector_key_enh_req_2600T = df_sector_key_enh_req[
        (df_sector_key_enh_req.enhancerequired7 == 1) & (df_sector_key_enh_req.ischanged8 == 1)]
    df_sector_key_enh_req_2600 = df_sector_key_enh_req[
        (df_sector_key_enh_req.enhancerequired6 == 1) & (df_sector_key_enh_req.ischanged7 == 1)]
    df_sector_key_enh_req_2100 = df_sector_key_enh_req[
        (df_sector_key_enh_req.enhancerequired5 == 1) & (df_sector_key_enh_req.ischanged6 == 1)]

    writer_sector_key_enh_req = pd.ExcelWriter('sector_key_enh_req.xlsx', engine='xlsxwriter')

    df_sector_key_enh_req_newSite.to_excel(writer_sector_key_enh_req, 'New Site')
    df_sector_key_enh_req_2600T.to_excel(writer_sector_key_enh_req, '2600T')
    df_sector_key_enh_req_2600.to_excel(writer_sector_key_enh_req, '2600')
    df_sector_key_enh_req_2100.to_excel(writer_sector_key_enh_req, '2100')
    # df_sector_key_enh_req.to_excel(writer_sector_key_enh_req, 'Sheet 1')

    writer_sector_key_enh_req.close()


if __name__ == "__main__":
    get_steps()
    get_ecells()
    get_cells_weak_coverage_choice()
    get_sector_key(df_choice, df_ecell_list)
    get_ecells_enhance_required(df_res_capacity, df_sector_key_weak_coverage)
