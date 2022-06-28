import pandas as pd
import tkinter.messagebox as MsgBox

from datetime import datetime


class ExcelData:

    __rawData = pd.DataFrame()
    __AIssue = pd.DataFrame()

    __onHand = dict()
    __onHand_AIssue = dict()

    __needToOrder = dict()

    def __init__(self):
        pass

    def readExcel(self, filedir):
        try:
            df = pd.read_excel(filedir, sheet_name='MO_Allocate')
            new_header = df.iloc[0]  # grab the first row for the header
            df = df[1:]  # take the data less the header row
            df.columns = new_header  # set the header row as the df header

            df.fillna(method='ffill', inplace=True)

            df = df.drop(len(df))
            df = df.drop(columns=["Draw no."])
            df = df.drop(columns=df.columns[9:14])
            df = df.drop(columns=df.columns[13:])
            df.fillna('-', inplace=True)

            for row_index, row in df.iterrows():
                if ('12' in str(row['Item No.'])[:2] or
                    '26' in str(row['Model'])[:2] or
                    '50' in str(row['Item No.'])[:2] or
                    '51' in str(row['Item No.'])[:2] or
                    'HB' in str(row['M/O No.']) or
                    'HM' in str(row['M/O No.']) or
                    'C' in str(row['M/O No.']) or
                    'C' in str(row['Model'])[0] or
                    'D' in str(row['M/O No.']) or
                        'P' in str(row['Item No.'])[0]):
                    df = df.drop(row_index)

            for key, value in df['ISSUE_DATE'].iteritems():
                if type(value) == datetime:
                    y = value.year
                    d = "0" + \
                        str(value.month) if value.month < 10 else str(value.month)
                    m = "0" + \
                        str(value.day) if value.day < 10 else str(value.day)
                    # x = d +'/'+ m +'/'+ str(y)
                    x = str(y) + '-' + m + '-' + d
                    df.loc[key, 'ISSUE_DATE'] = x
                elif type(value) == str:
                    df.loc[key, 'ISSUE_DATE'] = datetime.strptime(
                        value, "%d/%m/%Y")
                else:
                    df.loc[key, 'ISSUE_DATE'] = value.to_pydatetime()

            df['ISSUE_DATE'] = pd.to_datetime(
                df['ISSUE_DATE'], format="%Y-%m-%d", errors='ignore')
            df = df.astype({'M/O Qty.': 'float', 'ALC Qty.': 'float',
                           "Mat't_Onhand": 'float', 'Item No.': 'str'})
            df = df.sort_values(by=['ISSUE_DATE', 'M/O No.', 'M/O Qty.'])
            self.__rawData = df
        except Exception as e:
            MsgBox.showwarning(f'Sheet (MO_Allocate) Not Found', e)

    def readExcelStock(self, filedir):
        try:
            df = pd.read_excel(filedir, sheet_name='Recieve')
            new_header = df.iloc[0]  # grab the first row for the header
            df = df[1:]  # take the data less the header row
            df.columns = new_header  # set the header row as the df header

            result = self.__onHand

            for row_index, row in df.iterrows():
                # print(df.iloc[row_index-1][0])
                x = df.iloc[row_index-1][0]
                y = df.iloc[row_index-1][1]
                try:
                    result[x] = result.get(x) + int(y)
                except Exception as e:
                    print(e)
        except Exception as e:
            MsgBox.showwarning(f'Sheet (Recieve) Not Found', e)

    def createRawDataHeader(self):
        return tuple(self.__rawData.columns)

    def createRawData(self):
        return self.__rawData.values.tolist()

    def createOnHandData(self):
        result = {}
        df = self.__rawData
        for row_index, row in df.iterrows():
            result[str(df.loc[row_index, 'Item No.'])] = float(
                df.loc[row_index, "Mat't_Onhand"])

        self.__onHand = result

    def createOnHandData_type(self, typeofOnHand):
        result = self.__onHand
        sap_other = []
        shaft = []
        rotor = []
        magnet = []
        spacer = []
        stator = []
        flange = []

        for k in result:
            if '10' in str(k)[:2]:
                sap_other.append((k, result.get(k)))
            elif '53' in str(k)[:2]:
                shaft.append(((k, result.get(k))))
            elif '26' in str(k)[:2]:
                rotor.append(((k, result.get(k))))
            elif '58' in str(k)[:2]:
                magnet.append(((k, result.get(k))))
            elif '56' in str(k)[:2]:
                spacer.append(((k, result.get(k))))
            elif '19' in str(k)[:2]:
                stator.append(((k, result.get(k))))
            elif '51' in str(k)[:2]:
                flange.append(((k, result.get(k))))
            else:
                sap_other.append((k, result.get(k)))

        if typeofOnHand == 'shaft':
            return shaft
        elif typeofOnHand == 'rotor':
            return rotor
        elif typeofOnHand == 'magnet':
            return magnet
        elif typeofOnHand == 'spacer':
            return spacer
        elif typeofOnHand == 'stator':
            return stator
        elif typeofOnHand == 'flange':
            return flange
        else:
            return sap_other

    def createRequestPartData(self, typeofOnHand):
        result = self.__onHand_AIssue
        sap_other = []
        shaft = []
        rotor = []
        magnet = []
        spacer = []
        stator = []
        flange = []

        for k in result:
            if '10' in str(k)[:2]:
                sap_other.append((k, result.get(k)))
            elif '53' in str(k)[:2]:
                shaft.append(((k, result.get(k))))
            elif '26' in str(k)[:2]:
                rotor.append(((k, result.get(k))))
            elif '58' in str(k)[:2]:
                magnet.append(((k, result.get(k))))
            elif '56' in str(k)[:2]:
                spacer.append(((k, result.get(k))))
            elif '19' in str(k)[:2]:
                stator.append(((k, result.get(k))))
            elif '51' in str(k)[:2]:
                flange.append(((k, result.get(k))))
            else:
                sap_other.append((k, result.get(k)))

        if typeofOnHand == 'shaft':
            return shaft
        elif typeofOnHand == 'rotor':
            return rotor
        elif typeofOnHand == 'magnet':
            return magnet
        elif typeofOnHand == 'spacer':
            return spacer
        elif typeofOnHand == 'stator':
            return stator
        elif typeofOnHand == 'flange':
            return flange
        else:
            return sap_other

    def createNeedToOrder(self, typeofOnHand):
        result = self.__needToOrder
        sap_other = []
        shaft = []
        rotor = []
        magnet = []
        spacer = []
        stator = []
        flange = []

        for k in result:
            if '10' in str(k)[:2]:
                sap_other.append((k, result.get(k)))
            elif '53' in str(k)[:2]:
                shaft.append(((k, result.get(k))))
            elif '26' in str(k)[:2]:
                rotor.append(((k, result.get(k))))
            elif '58' in str(k)[:2]:
                magnet.append(((k, result.get(k))))
            elif '56' in str(k)[:2]:
                spacer.append(((k, result.get(k))))
            elif '19' in str(k)[:2]:
                stator.append(((k, result.get(k))))
            elif '51' in str(k)[:2]:
                flange.append(((k, result.get(k))))
            else:
                sap_other.append((k, result.get(k)))

        if typeofOnHand == 'shaft':
            return shaft
        elif typeofOnHand == 'rotor':
            return rotor
        elif typeofOnHand == 'magnet':
            return magnet
        elif typeofOnHand == 'spacer':
            return spacer
        elif typeofOnHand == 'stator':
            return stator
        elif typeofOnHand == 'flange':
            return flange
        else:
            return sap_other

    def createDailyHeader(self):
        return tuple(self.__AIssue.columns)

    # https://www.geeksforgeeks.org/python-boolean-list-and-and-or-operations/
    def create_After_Issue(self):
        df = self.__rawData.copy()
        res = self.__onHand.copy()

        df = df[["ISSUE_DATE", "M/O No.", "Model",
                 "Item No.", "M/O Qty.", "ALC Qty."]]

        for row_index, row in df.iterrows():
            check = []
            z = row['M/O No.']
            a = df.index[df['M/O No.'] == z].to_list()

            for i in a:

                stock = float(res.get(str(df.loc[i, 'Item No.'])))
                issue = float(df.loc[i, 'ALC Qty.'])

                if stock - issue >= 0:
                    check.append(True)
                else:
                    check.append(False)

                if len(check) == len(a) and all(check):
                    for j in a:

                        if str(df.loc[j, 'Item No.']) == '10000062686':
                            print(str(df.loc[j, 'Item No.']), float(
                                res.get(str(df.loc[j, 'Item No.']))), float(df.loc[j, 'ALC Qty.']))

                        _stock = float(res.get(str(df.loc[j, 'Item No.'])))
                        _issue = float(df.loc[j, 'ALC Qty.'])

                        df.loc[j, 'B/Issue'] = _stock
                        df.loc[j, 'A/Issue'] = _stock - _issue
                        res[str(df.loc[j, 'Item No.'])] = float(
                            df.loc[j, 'A/Issue'])

                else:
                    pass

        print(res.get('10000062686'))
        df.fillna('-', inplace=True)
        self.__onHand_AIssue = res
        self.__AIssue = df

    def createDailyIssue(self, p_type) -> pd.DataFrame:
        df = self.__AIssue.copy()

        x = None
        y = None

        if p_type == 'rotor':
            x = '19'
            y = 'B'
        elif p_type == 'stator':
            x = '14'
            y = 'A'

        for row_index, row in df.iterrows():
            if x in str(row['Model'])[:2] or x in str(row['Item No.'])[:2] or y in str(row['M/O No.']):
                df = df.drop(row_index)
            elif row['A/Issue'] == '-':
                z = row['M/O No.']
                a = df.index[df['M/O No.'] == z].to_list()
                df = df.drop(a)
        return df

    def createShortage(self, p_type):
        res = self.__onHand_AIssue.copy()
        df0 = self.__AIssue.copy()
        df1 = self.createDailyIssue('rotor')
        df2 = self.createDailyIssue('stator')

        combine = list(set(df1.index.to_list() + df2.index.to_list()))
        df0.drop(combine, inplace=True)

        x = None
        y = None

        if p_type == 'rotor':
            x = '19'
            y = 'B'
        elif p_type == 'stator':
            x = '14'
            y = 'A'

        for row_index, row in df0.iterrows():
            if x in str(row['Model'])[:2] or x in str(row['Item No.'])[:2] or y in str(row['M/O No.']):
                df0 = df0.drop(row_index)
            else:
                stock = float(res.get(str(df0.loc[row_index, 'Item No.'])))
                issue = float(df0.loc[row_index, 'ALC Qty.'])

                df0.loc[row_index, 'B/Issue'] = stock
                df0.loc[row_index, 'A/Issue'] = stock - issue
                res[str(df0.loc[row_index, 'Item No.'])] -= issue

        self.__needToOrder = res

        return df0
