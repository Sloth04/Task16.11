import numpy as np
import pandas as pd
from glob import glob
import os
import xlsxwriter

pd.options.mode.chained_assignment = None  # default='warn'


owners = {'ПрАТ «Укргідроенерго» 62X5590123794668': ['DNIP1HPP', 'SEREDHPP', 'KANIVHPP', 'KAKHHPP',
                                                         'DNIP2HPP', 'KREMHPP', 'KYIVHPP', 'DNISTHPP'],
              'ПрАТ «Харківська ТЕЦ-5» 56XQ00005D36E00G': ['KHAR5CHPP'],
              'ТОВ «ДТЕК СХІДЕНЕРГО» 56X000000001590J': ['KURTPP', 'LUHTPP'],
              'АТ «ДТЕК ДНІПРОЕНЕРГО» 21X000000001335A': ['ZAPTPP', 'PRYDNTPP', 'KRYVTPP'],
              'АТ «ДТЕК ЗАХІДЕНЕРГО» 23X-UKR-ZAKHID-4': ['BURSHTPP-BEI', 'LADTPP', 'DOBTPP']}


def mod(key, df, dct_mask):
    writer = pd.ExcelWriter(f'./output/{key}.xlsx', engine='xlsxwriter')
    for item in dct_mask[key]:
        split_list = item.split()
        df_temp = df.loc[(df['Power plant'] == split_list[0]) & (df['Product type'] == split_list[1])]
        df_temp.loc[df_temp.index.min(), 'Total'] = df_temp['Deniushka'].sum()
        df_temp.to_excel(writer, sheet_name=item, index=False)
    writer.save()


def main():
    dict_mask = {}
    list_query = []
    cwd = os.path.dirname(os.path.abspath(__file__))
    target = os.path.join(cwd, "input", '*.xlsx')
    df = pd.read_excel(glob(target)[0], 0)
    df.loc[~df['Monitoring result'], 'Deniushka'] = 0
    for owner in owners.keys():
        list_name = owners.get(owner)
        for item in range(len(list_name)):
            query = ([f'{owners[owner][item]} {x}' for x in set(df[df['Power plant'] == owners[owner][item]]['Product type'])])
            list_query += query
        dict_mask[owner] = list_query
        list_query = []
        mod(owner, df, dict_mask)


if __name__ == '__main__':
    main()
