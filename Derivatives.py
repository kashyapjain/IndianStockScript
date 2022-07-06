import logging
from os import name
from dateutil.relativedelta import relativedelta
import pandas
from pandas.tseries.offsets import Day
from nsepy import*
from datetime import *
from datetime import timedelta
from matplotlib import *
import numpy
import openpyxl


#----- Global Variable -----

StockName = ""


#----- Getting Futures Data -----

def FutData(ContractDate,FC_StartDate,FC_EndDate):
    global StockName

    if(FC_EndDate>date.today()):
        FC_EndDate = date.today() - relativedelta(days=1)
        
    FuturesHistData = get_history(symbol=StockName, futures=True ,expiry_date=ContractDate, start=FC_StartDate, end=FC_EndDate)

    return FuturesHistData


#----- Getting Futures Data of FirstMonth -----

def GetFuturesData_FM(CurrentDate,PYPM_ContractExpiryDate):
    global StockName

    df = pandas.DataFrame()

    PM_ContractDate = PYPM_ContractExpiryDate

    while(PM_ContractDate<=CurrentDate):            
        thisMonthDate = PM_ContractDate + relativedelta(months=1)

        TM_ContractDate = max(get_expiry_date(thisMonthDate.year,thisMonthDate.month,stock=True,index=False))

        PM_ContractDate = PM_ContractDate + relativedelta(days=1)

        FutureContractData = FutData(TM_ContractDate,PM_ContractDate,TM_ContractDate)      
        FutureContractData =  AddDatetoDF(FutureContractData)

        df = pandas.concat([df, FutureContractData], ignore_index=True)

        PM_ContractDate = TM_ContractDate

    return df


#----- Getting Futures Data of FirstMonth -----

def GetFuturesData_SM(CurrentDate,PYPM_ContractExpiryDate):
    global StockName

    df = pandas.DataFrame()

    PM_ContractDate = PYPM_ContractExpiryDate

    while(PM_ContractDate<=CurrentDate):            
        thisMonthDate = PM_ContractDate + relativedelta(months=2)

        TM_ContractDate = max(get_expiry_date(thisMonthDate.year,thisMonthDate.month,stock=True,index=False))

        oneMinusMonthDate = TM_ContractDate - relativedelta(months=1)
        oneMinusTM_ContractDate = max(get_expiry_date(oneMinusMonthDate.year,oneMinusMonthDate.month,stock=True,index=False))

        PM_ContractDate = PM_ContractDate + relativedelta(days=1)

        FutureContractData = FutData(TM_ContractDate,PM_ContractDate,oneMinusTM_ContractDate)      
        FutureContractData =  AddDatetoDF(FutureContractData)

        df = pandas.concat([df, FutureContractData], ignore_index=True)

        oneMinusMonthDate = TM_ContractDate - relativedelta(months=1)
        oneMinusTM_ContractDate = max(get_expiry_date(oneMinusMonthDate.year,oneMinusMonthDate.month,stock=True,index=False))

        PM_ContractDate = oneMinusTM_ContractDate

    return df


#----- Adding Date to DataFrame -----

def AddDatetoDF(df):
    DateList=[]

    for i in df.index:
        DateList.append(i)

    df['Date']=DateList

    return df

#----- Operations -----

def Operations(EquityHistData,df1,df2):
    AnalysisDF = pandas.DataFrame()

    AnalysisDF.insert(0,"Date",EquityHistData['Date'])
    
    ExpiryDateList = DataToList(df1['Expiry'])
    AnalysisDF['Expiry'] = ExpiryDateList

    AnalysisDF.insert(2,"Close(Price)",EquityHistData['Close'])

    _Del = (EquityHistData['Deliverable Volume'] * EquityHistData['VWAP']) / 10000000

    AnalysisDF.insert(3,"Del",_Del)


    _DAD5 = AvgDel_5(DataToList(AnalysisDF['Del']))
    AnalysisDF.insert(4,"5DAD",_DAD5)

    df1_COI = DataToList(df1['Open Interest'])
    _COI = AddLists(df1_COI,df2['Open Interest'])         
    AnalysisDF['COI'] = _COI

    ACinCOI = AbsolouteChange(_COI)
    AnalysisDF['ACinD Share'] = ACinCOI

    AnalysisDF['Blank'] = " "
    
    AnalysisDF['Date2'] = AnalysisDF['Date']

    ADF_Price = DataToList(AnalysisDF['Close(Price)'])
    PCD_ADF = PerChgInData(ADF_Price)
    AnalysisDF['PC Price'] = PCD_ADF

    
    List_DEL = DataToList(_Del)
    PC_Del = LDL(List_DEL,_DAD5)
    AnalysisDF['PC Del'] = PC_Del


    PC_OI = PerChgInData(_COI)
    AnalysisDF['PC COI'] = PC_OI

    AnalysisDF['Blank2'] = " "

    AnalysisDF['Long'] = " "

    AnalysisDF['Short'] = " "
    
    InitLS(AnalysisDF)

    AnalysisDF['Blank3'] = " "

    AnalysisDF['VWAP'] = EquityHistData['VWAP']

    AnalysisDF['O'] = EquityHistData['Open']

    AnalysisDF['H'] = EquityHistData['High']

    AnalysisDF['L'] = EquityHistData['Low']

    AnalysisDF['C'] = EquityHistData['Close']

    AnalysisDF = GetEquityData_52W_HL(AnalysisDF)

    AnalysisDF['52WHigh'] = pcHigh(AnalysisDF['C'],AnalysisDF['52WH'])

    AnalysisDF['52WLow'] = pcLow(AnalysisDF['C'],AnalysisDF['52WL'])

    return AnalysisDF


#----- Percentage Change in 52WH -----

def pcHigh(Close,H):
    i = 0
    pcH = []

    while(i < len(H)):
        Val = (H[i] - Close[i]) / H[i]

        Val = round(Val * 100 , 2)

        pcH.append(Val)

        i = i + 1

    return pcH


#----- Percentage Change in 52WL -----

def pcLow(Close,L):
    i = 0
    pcL = []

    while(i < len(L)):
        Val = (Close[i] - L[i]) / L[i]

        Val = round(Val * 100 , 2)

        pcL.append(Val)

        i = i + 1

    return pcL


#----- Initialising Long Short -----

def InitLS(df):
    i = 1
    df['Long'][0] = 0
    df['Short'][0] = 0
    
    while(i<len(df['ACinD Share'])):
        if(df.iloc[i]['PC Price'] > 0 and df.iloc[i]['PC COI'] > 0):
            df['Long'][i] = df['ACinD Share'][i]

        elif(df.iloc[i]['PC Price'] > 0 and df.iloc[i]['PC COI'] < 0):
            df['Short'][i] = df['ACinD Share'][i]

        elif(df.iloc[i]['PC Price'] < 0 and df.iloc[i]['PC COI'] < 0):
            df['Long'][i] = df['ACinD Share'][i]

        elif(df.iloc[i]['PC Price'] < 0 and df.iloc[i]['PC COI'] > 0):
            df['Short'][i] = df['ACinD Share'][i]
            
        i = i + 1




#----- Formatting DataFrame -----

def FormatDF(df):
    SD_Price = numpy.std(df['PC Price'])
    SD_Del = sum(df['PC Del'][4:]) / len(df['PC Del'][4:])
    SD_COI = numpy.std(df['PC COI'])
    
    df['PC Price'][0] = SD_Price
    df['PC Del'][0] = SD_Del
    df['PC COI'][0] = SD_COI
    

    FormatedDF = df.style\
        .apply(lambda x: ['background-color: %s' % 'grey' if ExpiryChange(x,df) else 'background-colour: %s' % 'white' for x in df['Date']], axis=0)\
        .applymap(lambda x: 'color: %s' % 'blue;font-weight: bold' if MonthChange(x,df) else 'background-colour: %s' % 'white', subset = ['COI'])\
        .applymap(lambda x: 'color: %s' % 'green; font-weight: bold' if x > SD_Price else 'background-colour: %s' % 'white', subset = ['PC Price'])\
        .applymap(lambda x: 'background-color: %s' % 'yellow' if x > SD_Del else 'background-colour: %s' % 'white', subset = ['PC Del'])\
        .applymap(lambda x: 'color: %s' % 'green; font-weight: bold' if x > SD_COI else 'background-colour: %s' % 'white', subset = ['PC COI'])\
        .applymap(lambda x: 'color: %s' % 'red;font-weight: bold' if x < -SD_Price else 'background-colour: %s' % 'white', subset = ['PC Price'])\
        .applymap(lambda x: 'color: %s' % 'red;font-weight: bold' if x < -SD_COI else 'background-colour: %s' % 'white', subset = ['PC COI'])\
       
    return FormatedDF


#----- Change of Expiry -----

def ExpiryChange(x,df):
    MatchedDF = df.loc[df['Date'] == x]
    if(MatchedDF.iloc[0]['Expiry'] == MatchedDF.iloc[0]['Date']):
        return True
    else:
        return False
    
    

#----- Check for Month change -----

def MonthChange(x,df):
    MatchedDF = df.loc[df['COI'] == x]

    if(df.iloc[0]['Date'] == MatchedDF.iloc[0]['Date']):
        return True

    i = 1

    while(i<len(df['Date'])):
        if(df.iloc[i]['Date'] == MatchedDF.iloc[0]['Date']):
            if(df.iloc[i]['Date'].month > df.iloc[i-1]['Date'].month):
                return True

        i = i + 1


#----- Absouloute Change in Data -----

def AbsolouteChange(DataList):
    i = 1
    AC_DL = [0]

    while(i<len(DataList)):
        AC_DL.append(DataList[i] - DataList[i-1])

        i = i + 1

    return AC_DL


#----- List Devided by Percentage Change List -----

def LDL(df1,df2):
    i = 0
    _ldl =[]

    while(i<len(df1)):
        _ldl.append(round((df1[i] / df2[i])*100,2))

        i= i+ 1

    return _ldl


#----- 5 Day Average of df -----

def  AvgDel_5(df):
    i = 4
    _AvgDel = [1,1,1,1,1]
    
    while(i<len(df)-1):
        j = i

        sum = 0

        while(j>=i-4):
            sum = sum + df[j]

            j=j-1
        
        avg = sum / 5
        _AvgDel.append(avg)

        i=i+1

    return _AvgDel


#----- Percentage Change in Data -----

def PerChgInData(df):
    PCD = [0,]
    j=1

    while(j<len(df)):
        Change = round(((df[j]-df[j-1]) / df[j-1])*100,2)

        PCD.append(Change)
        
        j=j+1
    
    return PCD


#----- Adding Lists -----

def AddLists(df1,df2):
    _AddList = []
    j=0

    for k in df2:
         _AddList.append(df1[j] + k)

         j=j+1

    return _AddList


#----- Converting Data into List -----

def DataToList(df):
    df2 = []

    for i in df:
         df2.append(i)

    return df2


#----- Get 52 Week High Low -----

def Get52W_HL(P2y_To_P1y_EquityData):
    i = 0

    H_52W = P2y_To_P1y_EquityData['High'][i]
    L_52W = P2y_To_P1y_EquityData['Low'][i]

    while(i < len(P2y_To_P1y_EquityData['Close'])):
        if(H_52W < P2y_To_P1y_EquityData['High'][i]):
            H_52W = P2y_To_P1y_EquityData['High'][i]            
        
        if(L_52W > P2y_To_P1y_EquityData['Low'][i]):
            L_52W = P2y_To_P1y_EquityData['Low'][i]
        
        i = i + 1

    return {"52WH":H_52W,"52WL":L_52W}


#----- Add 52 Week High Low -----

def Add_52W_HL(EquityData,HL_52W):
    i = 0

    H_52W = HL_52W['52WH']
    L_52W = HL_52W['52WL']

    H_List = []
    L_List = []

    while(i < len(EquityData['C'])):
        if(EquityData['H'][i] > H_52W):
            H_52W = EquityData['H'][i]
            H_List.append(EquityData['H'][i])     
        
        else:
            H_List.append(H_52W)

        if(EquityData['L'][i] < L_52W):
            L_52W = EquityData['L'][i]
            L_List.append(EquityData['L'][i])
            
        else:
            L_List.append(L_52W)

        i = i + 1

    EquityData['52WH'] = H_List
    EquityData['52WL'] = L_List

    return EquityData


#----- Get  Equity Data -----

def GetEquityData_52W_HL(df):
    StartDate = date.today() - relativedelta(years=2)
    EndDate = date.today() - relativedelta(years=1)
    
    P2y_To_P1y_EquityData = get_history(symbol=StockName,start=StartDate,end=EndDate)

    HL_52W = Get52W_HL(P2y_To_P1y_EquityData)

    EquityData_52WHL = Add_52W_HL(df,HL_52W)

    return EquityData_52WHL


#----- Main -----

def main():
    global StockName

    logging.basicConfig(level=logging.DEBUG)

    StockName = input("\nThe Data will be of Past 1 Year\nEnter the symbol :")

    CurrentDate = date.today()  
    PY_CurrentDate = CurrentDate - relativedelta(years=1)

    P_ExpiryMonthDate = PY_CurrentDate - relativedelta(months=1)

    PYPM_ContractExpiryDate= max(get_expiry_date(PY_CurrentDate.year,P_ExpiryMonthDate.month,stock=True,index=False))

    #----- Equity -----

    StartForEquity = PYPM_ContractExpiryDate + relativedelta(days=1)
    
    CurrentDate = CurrentDate - relativedelta(days=1)

    EquityHistData = get_history(symbol=StockName,start=StartForEquity, end=CurrentDate)
    
    EquityHistData = AddDatetoDF(EquityHistData)

    #------------------

    df1 = GetFuturesData_FM(CurrentDate,PYPM_ContractExpiryDate)

    df2 = GetFuturesData_SM(CurrentDate,PYPM_ContractExpiryDate)

    i = 0

    while(i<len(EquityHistData['Date'])):
        if(EquityHistData['Date'][i] != df1['Date'][i] and EquityHistData['Date'][i] != df2['Date'][i]):
            if(df1['Date'][i]>EquityHistData['Date'][i]):
                EquityHistData.drop([EquityHistData['Date'][i]] , inplace = True)

                print(f"\n\n{EquityHistData['Date'][i]} Equity has been deleted\n\n")
            else:
                df1.drop([df1['Date'][i]] , inplace = True)

                print(f"\n\n{df1['Date'][i]} derivatives has been deleted\n\n")

        i =i +1

    print(EquityHistData)
    print(df1)
    print(df2)

    AnalysisDF = Operations(EquityHistData,df1,df2)  
    
    # FilePath = f'E:\kashy_mxe0hzp\zerodha\StocksExcel\\{StockName}{date.today()}.xlsx'
    # AnalysisDF.to_excel(FilePath, engine='openpyxl', index=False)

#----------------------------------------------------------------------------------------------------

if __name__=="__main__":
    main()
    
# FilePath = f'E:\kashy_mxe0hzp\zerodha\StocksExcel\\{StockName}{date.today()}.xlsx'
# df = pandas. read_excel (FilePath, sheet_name='Sheet1' , engine='openpyxl')

# FormatedDF = FormatDF(df)

# FormatedDF
# FormatedDF.to_excel(FilePath, engine='openpyxl', index=False)
