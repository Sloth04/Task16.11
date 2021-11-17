import numpy as np
import pandas as pd
from glob import glob
import os
import xlsxwriter

owners = {'ПрАТ «Укргідроенерго» 62X5590123794668': ['DNIP1HPP аРВЧ_з', 'DNIP1HPP аРВЧ_с',
                                                     'SEREDHPP аРВЧ_з', 'SEREDHPP аРВЧ_с',
                                                     'KANIVHPP аРВЧ_з', 'KANIVHPP аРВЧ_с',
                                                     'KAKHHPP аРВЧ_з',  'KAKHHPP аРВЧ_с',
                                                     'DNIP2HPP аРВЧ_з', 'DNIP2HPP аРВЧ_с',
                                                     'KREMHPP аРВЧ_з',  'KREMHPP аРВЧ_с',
                                                     'KYIVHPP аРВЧ_з',  'KYIVHPP аРВЧ_с',
                                                     'DNISTHPP рРВЧ_з'],
          'ПрАТ «Харківська ТЕЦ-5» 56XQ00005D36E00G': ['KHAR5CHPP аРВЧ_р', 'KHAR5CHPP аРВЧ_с',
                                                       'KHAR5CHPP РПЧ_с','KHAR5CHPP рРВЧ_з',
                                                       'KHAR5CHPP рРВЧ_р'],
          'ТОВ «ДТЕК СХІДЕНЕРГО» 56X000000001590J': ['KURTPP аРВЧ_з', 'KURTPP аРВЧ_c',
                                                     'KURTPP аРВЧ_p', 'KURTPP РПЧ_с',
                                                     'LUHTPP РПЧ_с'],
          'АТ «ДТЕК ДНІПРОЕНЕРГО» 21X000000001335A': ['ZAPTPP РПЧ_с', 'ZAPTPP аРВЧ_р',
                                                      'ZAPTPP аРВЧ_с', 'ZAPTPP аРВЧ_з',
                                                      'PRYDNTPP рРВЧ_з', 'KRYVTPP рРВЧ_з'],
          'АТ «ДТЕК ЗАХІДЕНЕРГО» 23X-UKR-ZAKHID-4': ['BURSHTPP-BEI аРВЧ_з', 'BURSHTPP-BEI аРВЧ_с',
                                                     'BURSHTPP-BEI аРВЧ_р', 'BURSHTPP-BEI РПЧ_с',
                                                     'LADTPP рРВЧ_з', 'DOBTPP рРВЧ_з']}


def mod(key, df):
    dict_df = {}
    for item in owners[key]:
        split_list = item.split()
        df_temp = df.loc[(df['Power plant'] == split_list[0]) & (df['Product type'] == split_list[1])]
        dict_df[item] = df_temp
    writer = pd.ExcelWriter(f'./output/{key}.xlsx', engine='xlsxwriter')
    for item in dict_df.keys():
        dict_df[item].to_excel(writer, sheet_name=item, index=False)
    writer.save()
    # print(dict_df)


def main():
    cwd = os.path.dirname(os.path.abspath(__file__))
    target = os.path.join(cwd, "input", '*.xlsx')
    df = pd.read_excel(glob(target)[0], 0)
    df.loc[~df['Monitoring result'], 'Deniushka'] = 0
    for owner in owners.keys():
        mod(owner, df)


if __name__ == '__main__':
    main()
