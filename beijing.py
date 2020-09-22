

def con_mysql(vin_list):
    import pymysql
    con = pymysql.connect(host='10.129.41.89', port=4226, user='rvmdb', passwd='bZQ8EU*C', db='rvm')
    # con = MySQLdb.connect(host='10.130.142.115', port=3306, user='root', passwd='saic@2019', db='rvm')
    c = con.cursor()
    sql = """
            SELECT
                vin,
                clxh,
                package_no,
                package_type,
                iccid,
                sim_card
            FROM
                t_vehicle
            WHERE
                vin IN %s
            """
    c.execute(sql, (vin_list,))
    info_list = c.fetchall()
    c.close()
    return info_list


def check_and_fill(input_file):
    import pandas as pd
    writer = pd.ExcelWriter('BeiJing' + '.xlsx')
    data_frame = pd.read_excel(input_file, index_col=None, skiprows=2)
    vin_list = tuple(data_frame['VIN码'])
    info_list = pd.DataFrame(con_mysql(vin_list),
                             columns=['vin', 'clxh', 'package_no', 'package_type', 'iccid', 'sim_card'])
    for vin in vin_list:
        excel_no = data_frame.loc[data_frame['VIN码'] == vin, '动力蓄电池编码'].values
        mysql_no = info_list.loc[info_list['vin'] == vin, 'package_no'].values
        if excel_no.isnan():
            if mysql_no.isnan():
                print('{}\t为空\t在表格和数据库中均为空'.format(vin))
            else:
                data_frame.loc[data_frame['VIN码'] == vin, '动力蓄电池编码'] = mysql_no
                print('{}\t填充为\t{}'.format(vin, mysql_no))
        else:
            if mysql_no.isnan():
                print('{}\t为空\t在数据库中为空'.format(vin))
            else:
                if excel_no.strip() == mysql_no.strip():
                    pass
                else:
                    print('{}\t不一致\t'.format(vin))


def check_battery(input_file):
    """
    :param input_file:输入的文件
    :return:
    根据电池编码解析电池生产日期
    """
    import pandas as pd
    writer = pd.ExcelWriter('Output1_BeiJing_BatteryDateChecked' + '.xlsx')
    data_frame = pd.read_excel(input_file, index_col=None, skiprows=1)
    year_24based = {'1': '2011',
                    '2': '2012',
                    '3': '2013',
                    '4': '2014',
                    '5': '2015',
                    '6': '2016',
                    '7': '2017',
                    '8': '2018',
                    '9': '2019',
                    'A': '2020',
                    'B': '2021',
                    'C': '2022',
                    'D': '2023',
                    'E': '2024',
                    'F': '2025',
                    'G': '2026',
                    'H': '2027',
                    'J': '2028',
                    'K': '2029',
                    'L': '2030',
                    'M': '2031',
                    'N': '2032',
                    'P': '2033',
                    'R': '2034',
                    'S': '2035',
                    'T': '2036',
                    'V': '2037',
                    'W': '2038',
                    'X': '2039',
                    'Y': '2040'}
    year_15based = {'1': '2001',
                    '2': '2002',
                    '3': '2003',
                    '4': '2004',
                    '5': '2005',
                    '6': '2006',
                    '7': '2007',
                    '8': '2008',
                    '9': '2009',
                    'A': '2010',
                    'B': '2011',
                    'C': '2012',
                    'D': '2013',
                    'E': '2014',
                    'F': '2015',
                    'G': '2016',
                    'H': '2017',
                    'J': '2018',
                    'K': '2019',
                    'L': '2020',
                    'M': '2021',
                    'N': '2022',
                    'P': '2023',
                    'R': '2024',
                    'S': '2025',
                    'T': '2026',
                    'V': '2027',
                    'W': '2028',
                    'X': '2029',
                    'Y': '2030'
                    }
    month_and_date = {"0": "31",
                      "1": "1",
                      "2": "2",
                      "3": "3",
                      "4": "4",
                      "5": "5",
                      "6": "6",
                      "7": "7",
                      "8": "8",
                      "9": "9",
                      "A": "10",
                      "B": "11",
                      "C": "12",
                      "D": "13",
                      "E": "14",
                      "F": "15",
                      "G": "16",
                      "H": "17",
                      "J": "18",
                      "K": "19",
                      "L": "20",
                      "M": "21",
                      "N": "22",
                      "P": "23",
                      "R": "24",
                      "S": "25",
                      "T": "26",
                      "V": "27",
                      "W": "28",
                      "X": "29",
                      "Y": "30"
                      }
    vin_list = data_frame['VIN码']
    for row in range(0, len(vin_list)):
        if len(data_frame.loc[row, '动力蓄电池编码']) == 24:
            data_frame.loc[row, '动力蓄电池生产日期'] = ('{}/{}/{}'.format(year_24based[data_frame.loc[row, '动力蓄电池编码'][14]],
                                                                  month_and_date[data_frame.loc[row, '动力蓄电池编码'][15]],
                                                                  month_and_date[data_frame.loc[row, '动力蓄电池编码'][16]], ))
        elif len(data_frame.loc[row, '动力蓄电池编码']) == 15:
            data_frame.loc[row, '动力蓄电池生产日期'] = ('{}/{}/{}'.format(year_15based[data_frame.loc[row, '动力蓄电池编码'][7]],
                                                                  month_and_date[data_frame.loc[row, '动力蓄电池编码'][8]],
                                                                  data_frame.loc[row, '动力蓄电池编码'][9:11], ))
        else:
            print(data_frame.loc[row, 'VIN码'], '\t', '未匹配电池编码')

    data_frame.to_excel(writer, encoding="utf-8", index=None)
    writer.save()
    writer.close()


def model_split():
    """
    根据车型信息拆分表格便于上传
    """
    import pandas as pd
    input_file = str('Output1_BeiJing_BatteryDateChecked' + '.xlsx')
    data_frame = pd.read_excel(input_file, index_col=None, skiprows=0)


    vehicle_types = data_frame['备注'].unique().tolist()
    count = 0
    for vehicle_type in vehicle_types:
        wanted_vehicles = data_frame[data_frame['备注'].isin([vehicle_type])]
        print('{0:<20}{1:<10}'.format(vehicle_type, len(wanted_vehicles)))
        count += len(wanted_vehicles)
        vehicle_excel_name = '安全监控平台车辆信息批量导入模板_北京顺义_' + str(vehicle_type) + '(' + str(
            len(wanted_vehicles)) + '-X)' + '.xlsx'
        writer = pd.ExcelWriter(vehicle_excel_name)
        wanted_vehicles.to_excel(writer, sheet_name='Sheet1', encoding="utf-8", index=None, startrow=1)
        worksheet = writer.sheets['Sheet1']
        worksheet.write_string(0, 0, "注：限制导入最多500条！ vin码、动力蓄电池编码、动力蓄电池生产日期为必填项；vin码、车牌号不允许重复；销售日期为时间格式（2016/5/30"
                                     "）；动力蓄电池编码与动力蓄电池生产日期可为多个，最多为4个(多个时用“，”分开),"
                                     "2个是成对出现，2个的个数与所选车型动力蓄电池包个数相等；所属子车企/4S店：填写子车企/4S店全称，如匹配不上，则分配至主车企中("
                                     "子车企/4S店操作，该字段无效)；")
        writer.save()
        writer.close()
    print('总车辆数：{0}'.format(count))


if __name__ == "__main__":
    input_file = r'BeiJing.xlsx'
    check_battery(input_file)
    model_split()
