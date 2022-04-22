from functools import reduce
import operator
from unittest import result
import pandas as pd

from datetime import datetime

class ExcelData:

    __rawData = pd.DataFrame()
    __before_after = pd.DataFrame()
    __onHand = dict()
    __onHand_AIssue = dict()

    def __init__(self):
        pass

    def readExcel(self,filedir):
        df = pd.read_excel(filedir, sheet_name='MO_Allocate')
        new_header = df.iloc[0] #grab the first row for the header
        df = df[1:] #take the data less the header row
        df.columns = new_header #set the header row as the df header
        
        df = df.drop(len(df))
        df = df.drop(columns=df.columns[9:14])
        df = df.drop(columns=df.columns[13:])
        df.fillna('-', inplace = True)
        for key, value in df['M/O No.'].iteritems():
            if 'HB' in str(value) or 'C' in str(value) or 'D' in str(value):
                df = df.drop(key)

        for key, value in df['ISSUE_DATE'].iteritems():
            if type(value) == datetime :
                y = value.year
                d = "0"+ str(value.month) if value.month < 10 else str(value.month)
                m = "0"+ str(value.day) if value.day < 10 else str(value.day)
                # x = d +'/'+ m +'/'+ str(y)
                x = str(y) +'-'+ m +'-'+ d
                df.loc[key,'ISSUE_DATE'] = x
            elif type(value) == str :
                df.loc[key,'ISSUE_DATE'] = datetime.strptime(value, "%d/%m/%Y")
            else:
                print(type(value))
                df.loc[key,'ISSUE_DATE'] = datetime.strptime(value, "%d/%m/%Y")

        df['ISSUE_DATE'] = pd.to_datetime(df['ISSUE_DATE'], format = "%Y-%m-%d", errors='ignore')
        df = df.sort_values(by=['ISSUE_DATE','M/O No.'])
        self.__rawData = df


    def readExcelStock(self,filedir):
        df = pd.read_excel(filedir, sheet_name='Recieve')
        new_header = df.iloc[0] #grab the first row for the header
        df = df[1:] #take the data less the header row
        df.columns = new_header #set the header row as the df header
        
        result = self.__onHand

        # for part in df.iloc[:,0]:
        #     print(part)

        for row_index,row in df.iterrows():
            # print(df.iloc[row_index-1][0])
            x = df.iloc[row_index-1][0]
            y = df.iloc[row_index-1][1]
            try:
                # print(x)
                # print(result.get(x))
                result[x] = result.get(x) + int(y)
            except Exception as e:
                # result[x] = int(y)
                print(e)

    def createRawDataHeader(self):
        return tuple(self.__rawData.columns)

    def createRawData(self):
        return self.__rawData.values.tolist()

    def createOnHandData(self):
        result = {}
        df = self.__rawData
        for row_index,row in df.iterrows():
            result[df.loc[row_index,'Item No.']] = df.loc[row_index,"Mat't_Onhand"]

        self.__onHand = result

    def createOnHandData_type(self,typeofOnHand):
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
                sap_other.append((k,result.get(k)))
            elif '53' in str(k)[:2]:
                shaft.append(((k,result.get(k))))
            elif '26' in str(k)[:2]:
                rotor.append(((k,result.get(k))))
            elif '58' in str(k)[:2]:
                magnet.append(((k,result.get(k))))
            elif '56' in str(k)[:2]:
                spacer.append(((k,result.get(k))))
            elif '19' in str(k)[:2]:
                stator.append(((k,result.get(k))))
            elif '51' in str(k)[:2]:
                flange.append(((k,result.get(k))))
            else:
                sap_other.append((k,result.get(k)))

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

    def createRequestPartData(self,typeofOnHand):
        sap_other = []
        shaft = []
        rotor = []
        magnet = []
        spacer = []
        stator = []
        flange = []

        result = self.__onHand_AIssue

        for k in result:
            if '10' in str(k)[:2]:
                sap_other.append((k,result.get(k)))
            elif '53' in str(k)[:2]:
                shaft.append(((k,result.get(k))))
            elif '26' in str(k)[:2]:
                rotor.append(((k,result.get(k))))
            elif '58' in str(k)[:2]:
                magnet.append(((k,result.get(k))))
            elif '56' in str(k)[:2]:
                spacer.append(((k,result.get(k))))
            elif '19' in str(k)[:2]:
                stator.append(((k,result.get(k))))
            elif '51' in str(k)[:2]:
                flange.append(((k,result.get(k))))
            else:
                sap_other.append((k,result.get(k)))

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

    def create_Before_After(self):
        df  = self.__rawData.copy()
        res = self.__onHand.copy()

        df = df[["ISSUE_DATE", "M/O No.", "Model", "Item No.", "M/O Qty.", "ALC Qty."]]
        
        for row_index,row in df.iterrows():
            
            x = float(res.get(df.loc[row_index,'Item No.']))
            y = float(df.loc[row_index,'ALC Qty.'])

            df.loc[row_index,'B/Issue'] = x
            df.loc[row_index,'A/Issue'] = x - y
            res[df.loc[row_index,'Item No.']] = df.loc[row_index,'A/Issue']
        
        self.__onHand_AIssue = res
        self.__before_after = df

    def createDailyHeader(self):
        return tuple(self.__before_after.columns)

    # https://www.adamsmith.haus/python/answers/how-to-reorder-columns-in-a-pandas-dataframe-in-python
    def createDailyIssue(self, p_type) -> pd.DataFrame:
        df = self.__before_after.copy()

        x = None
        y = None

        if p_type == 'rotor':
            x = '19'
            y = 'B'
        elif p_type == 'stator':
            x = '14'
            y = 'A'
            
        for row_index, row in df.iterrows():
            if x in str(row['Model'])[:2] or x in str(row['Item No.'])[:2] or '51' in str(row['Item No.'])[:2] or y in str(row['M/O No.']):
                df = df.drop(row_index)
            elif row['A/Issue'] < 0:
                z = row['M/O No.']
                a = df.index[df['M/O No.'] == z].to_list()
                df = df.drop(a)

        return df

    def createShortage(self,p_type):
        df0 = self.__before_after.copy()
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
            if x in str(row['Model'])[:2] or x in str(row['Item No.'])[:2] or '51' in str(row['Item No.'])[:2] or y in str(row['M/O No.']):
                df0 = df0.drop(row_index)
        
        return df0