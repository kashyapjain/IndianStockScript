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

#----- Formating DataFrame -----

def FormatDF(df):
    SD_Price = numpy.std(df['PC Price'])
    SD_Del = numpy.std(df['PC Del'])
    SD_COI = numpy.std(df['PC COI'])
    

    FormatedDF = df.style\
        .applymap(lambda x: 'background-color: %s' % 'red' if MonthChange(x,df) else 'background-colour: %s' % 'white', subset = ['COI'])\
        .applymap(lambda x: 'background-color: %s' % 'red' if x < SD_Price else 'background-colour: %s' % 'white', subset = ['PC Price'])\
        .applymap(lambda x: 'background-color: %s' % 'red' if x < SD_Del else 'background-colour: %s' % 'white', subset = ['PC Del'])\
        .applymap(lambda x: 'background-color: %s' % 'red' if x < SD_COI else 'background-colour: %s' % 'white', subset = ['PC COI'])\
        .format({'PC Price': "{:.2%}"})\
        .format({'PC Del': "{:.2%}"})\
        .format({'PC COI': "{:.2%}"})

    return FormatedDF

#----- Check for Month change -----

def MonthChange(x,df):
    MatchedDF = df.loc[df['COI'] == x]

    if(df.iloc[0]['Date'] == MatchedDF.iloc[0]['Date']):
        return True

    i = 1

    while(i<len(df['Date'])):
        if(df.iloc[i]['Date'] == MatchedDF.iloc[0]['Date']):
            if(df.iloc[i]['Date'] > df.iloc[i-1]['Date']):
                return True

        i = i + 1