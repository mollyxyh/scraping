import pandas as pd
from bs4 import BeautifulSoup
import numpy as np
import urllib
import urllib.request
from datetime import datetime, timedelta
import time
from urllib.request import FancyURLopener
import pickle
import xlrd
import xlwt
import ssl
import os

# using when certification verified failed
# import wmi
# import httplib
# httplib.HTTPConnection._http_vsn = 10
# httplib.HTTPConnection._http_vsn_str = 'HTTP/1.0'
# nic_configs = wmi.WMI().Win32_NetworkAdapterConfiguration(IPEnabled=True)
# nic = nic_configs[0]
# nic.EnableDHCP()


def readInput(path, sheetname, userdefined):

    df = pd.read_excel(path, sheetname)
    filepath = df.at[0, userdefined]
    sheetname = df.at[1, userdefined]
    output = df.at[2, userdefined]
    focus = []
    ptr = 3
    while df.at[ptr, userdefined] != 'END':
        if df.at[ptr, userdefined] == 'Y':
            focus.append(df.at[ptr, 'AREA'])
        ptr += 1

    res = {'filepath': filepath, 'sheetname': sheetname, 'output': output, 'focus': focus}
    return res


def getData(path, sheetname, title):

    xl = pd.ExcelFile(path)
    df = xl.parse(sheetname)
    df = df[title]

    return df


def writeExcel(df, path, sheetname):

    df.to_excel(path, sheetname)


def getSector(ticker):

    website_before_ticker = 'https://researchtools.fidelity.com/ftgw/mloptions/goto/optionChain?symbol='
    website_after_ticker = '&Search=Search'
    fidelity_website = website_before_ticker + ticker + website_after_ticker
    try:
        context = ssl._create_unverified_context()
        sock = urllib.request.urlopen(fidelity_website, context=context)
    except:
        try:
            time.sleep(5)
            context = ssl._create_unverified_context()
            sock = urllib.request.urlopen(fidelity_website,context=context)
        except:
            sector = 'N/A'
            return sector
    try:
        htmlSource = sock.read().decode('utf-8')
    except:
        htmlSource = sock.read().decode('GBK')
    sock.close()
    soup = BeautifulSoup(htmlSource)
    selector = soup.findAll('td')
    try:
        selector_s = selector[8]
    except:
        sector = 'N/A'
        return sector

    print(ticker)
    if selector_s is None:
        ticker_try = ticker[:(len(ticker)-1)] + '/' + ticker[len(ticker)-1]
        fidelity_website_try = website_before_ticker+ticker_try+website_after_ticker
        try:
            sock_try = urllib.request.urlopen(fidelity_website_try)
        except:
            try:
                time.sleep(5)
                sock_try = urllib.request.urlopen(fidelity_website_try)
            except:
                sector = 'N/A'
                return sector
        try:
            htmlSource_try = sock_try.read().decode('utf-8')
        except:
            htmlSource_try = sock_try.read().decode('GBK')
        sock_try.close()
        soup_try = BeautifulSoup(htmlSource_try)
        selector_try = soup_try.findAll('td')
        try:
            selector_try_s = selector_try[8]
        except:
            sector = 'N/A'
            return sector
        if selector_try_s is None:
            sector = 'N/A'
        else:
            sector = selector_try_s.text.strip()
    else:
        sector = selector_s.text.strip()
    print(sector)
    return sector


def getVol3m_MktCap_Beta(ticker):
    website_before_ticker = 'http://finance.yahoo.com/q?s='
    yahoo_website = website_before_ticker+ticker
    try:
        sock = urllib.request.urlopen(yahoo_website)
    except:
        time.sleep(5)
        sock = urllib.request.urlopen(yahoo_website)
    try:
        htmlSource = sock.read().decode('utf-8')
    except:
        htmlSource = sock.read().decode('GBK')
    sock.close()
    soup = BeautifulSoup(htmlSource)
    table_2 = soup.findAll('table')
    try:
        table2 = table_2[0]
        if table2 is None:
            vol = 'N/A'
            cap = 'N/A'
            headings2 = []
        else:
            headings2 = [th.get_text() for th in table2.find_all('th')]
            numbers2 = [td.get_text() for td in table2.find_all('td')]
        if 'Avg Vol (3m)' in numbers2:
            vol_index = numbers2.index('Avg Vol (3m)')
            vol = numbers2[vol_index+1]
            vol = vol.replace('"', '').replace(',', '')
            print(vol)
        else:
            table2 = table_2[1]
            if table2 is None:
                vol = 'N/A'
                cap = 'N/A'
                headings2 = []
            else:
                headings2 = [th.get_text() for th in table2.find_all('th')]
                numbers2 = [td.get_text() for td in table2.find_all('td')]
            if 'Avg Vol (3m)' in numbers2:
                # vol_index = headings2.index('Avg Vol (3m):')
                vol_index = numbers2.index('Avg Vol (3m)')
                vol = numbers2[vol_index+1]
                vol = vol.replace('"', '').replace(',', '')
                print (vol)
            else:
                table2 = table_2[2]
                if table2 is None:
                    vol = 'N/A'
                    cap = 'N/A'
                    headings2 = []
                else:
                    headings2 = [th.get_text() for th in table2.find_all('th')]
                    numbers2 = [td.get_text() for td in table2.find_all('td')]
                if 'Avg Vol (3m)' in numbers2:
                    # vol_index = headings2.index('Avg Vol (3m):')
                    vol_index = numbers2.index('Avg Vol (3m)')
                    vol = numbers2[vol_index+1]
                    vol = vol.replace('"', '').replace(',', '')
                    print (vol)
                else:
                    vol = 'N/A'
    except:
        vol = 'N/A'
    table_1 = soup.findAll('table')
    try:
        table1 = table_1[0]
        if table1 is None:
            beta = 'N/A'
            headings1 = []
        else:
            headings1 = [th.get_text() for th in table1.find_all('th')]
            numbers1 = [td.get_text() for td in table1.find_all('td')]
        if 'Beta' in numbers1:
            beta_index = numbers1.index('Beta')
            beta = numbers1[beta_index+1]
            print(beta)
        else:
            table1 = table_1[1]
            if table1 is None:
                beta = 'N/A'
                headings1 = []
            else:
                headings1 = [th.get_text() for th in table1.find_all('th')]
                numbers1 = [td.get_text() for td in table1.find_all('td')]
            if 'Beta' in numbers1:
                beta_index = numbers1.index('Beta')
                beta = numbers1[beta_index+1]
                print (beta)
            else:
                table1 = table_1[2]
                if table1 is None:
                    beta = 'N/A'
                    headings1 = []
                else:
                    headings1 = [th.get_text() for th in table1.find_all('th')]
                    numbers1 = [td.get_text() for td in table1.find_all('td')]
                if 'Beta' in numbers1:
                    beta_index = numbers1.index('Beta')
                    beta = numbers1[beta_index+1]
                    print(beta)
                else:
                    table1 = table_1[3]
                    if table1 is None:
                        beta = 'N/A'
                        headings1 = []
                    else:
                        headings1 = [th.get_text() for th in table1.find_all('th')]
                        numbers1 = [td.get_text() for td in table1.find_all('td')]
                    if 'Beta' in numbers1:
                        beta_index = numbers1.index('Beta')
                        beta = numbers1[beta_index+1]
                        print(beta)
                    else:
                        beta = 'N/A'
    except:
        beta = 'N/A'
    res = {'Vol3m': vol, 'Beta': beta}
    return res


def getVol10d(ticker, to_date, from_date):
    from_date_url = '&a=%s&b=%s&c=%s' % (from_date.month-1, from_date.day, from_date.year)
    to_date_url = '&d=%s&e=%s&f=%s' % (to_date.month-1, to_date.day, to_date.year)
    query_url = 'http://real-chart.finance.yahoo.com/table.csv?s=' + ticker + from_date_url + to_date_url + 'g=d&ignore=.csv'
    try:
        vol_10d_df = pd.read_csv(query_url)
        vol_10d = vol_10d_df['Volume']
        avg_vol_10d_temp = np.mean(vol_10d)
        avg_vol_10d = int(round(avg_vol_10d_temp))
    except:
        try:
            ticker_try = ticker[:(len(ticker)-1)] + '-' + ticker[len(ticker)-1]
            query_url2 = 'http://real-chart.finance.yahoo.com/table.csv?s=' + ticker_try + from_date_url + to_date_url + 'g=d&ignore=.csv'
            vol_10d = pd.read_csv(query_url2)['Volume']
            avg_vol_10d_temp = np.mean(vol_10d)
            avg_vol_10d = int(round(avg_vol_10d_temp))
        except:
            avg_vol_10d = 'N/A'
    return avg_vol_10d


def getM10d(ticker, to_date, from_date):
    from_date_url = '&a=%s&b=%s&c=%s' % (from_date.month-1, from_date.day, from_date.year)
    to_date_url = '&d=%s&e=%s&f=%s' % (to_date.month-1, to_date.day, to_date.year)
    query_url = 'http://real-chart.finance.yahoo.com/table.csv?s=' + ticker + from_date_url + to_date_url + 'g=d&ignore=.csv'
    try:
        m_10d_df = pd.read_csv(query_url)
        m_10d = m_10d_df['Volume']
        median_vol_10d = np.median(m_10d)
    except:
        try:
            ticker_try = ticker[:(len(ticker)-1)] + '-' + ticker[len(ticker)-1]
            query_url2 = 'http://real-chart.finance.yahoo.com/table.csv?s=' + ticker_try + from_date_url + to_date_url + 'g=d&ignore=.csv'
            m_10d = pd.read_csv(query_url2)['Volume']
            median_vol_10d = np.median(m_10d)
        except:
            median_vol_10d = 'N/A'
    return median_vol_10d


def getM60d(ticker, to_date, add60_date):
    add60_date_url = '&a=%s&b=%s&c=%s' % (add60_date.month-1, add60_date.day, add60_date.year)
    to_date_url = '&d=%s&e=%s&f=%s' % (to_date.month-1, to_date.day, to_date.year)
    query_url = 'http://real-chart.finance.yahoo.com/table.csv?s=' + ticker + add60_date_url + to_date_url + 'g=d&ignore=.csv'
    try:
        m_60d_df = pd.read_csv(query_url)
        m_60d = m_60d_df['Volume']
        median_vol_60d = np.median(m_60d)
    except:
        try:
            ticker_try = ticker[:(len(ticker)-1)] + '-' + ticker[len(ticker)-1]
            query_url2 = 'http://real-chart.finance.yahoo.com/table.csv?s=' + ticker_try + add60_date_url + to_date_url + 'g=d&ignore=.csv'
            m_60d = pd.read_csv(query_url2)['Volume']
            median_vol_60d = np.median(m_60d)
        except:
            median_vol_60d = 'N/A'
    return median_vol_60d


def updateData(data, focus, to_date, from_date, add60_date):
    print(data)
    inputlst = data['Ticker']
    lstlen = len(inputlst)
    check_sector = []
    check_vol10d = []
    check_vol3m = []
    check_m60d = []
    check_m10d = []
    check_beta = []
    check_mktcap = []
    sector_lst = ['Communications', 'Consumer Discretionary', 'Consumer Staples', 'Energy','Financials', 'Health Care',
                  'Industrials', 'Materials', 'Technology', 'Utilities']
    try:
        f = open(path_pickle+'Universe_trade.pickle', 'rb')
        start_num = pickle.load(f)
    except IOError:
        print("The process is starting from the first ticker. No need to continue from checkpoint")
        start_num = 0
    for i in range(start_num, lstlen):
        print(i)
        _ticker = str(inputlst[i])
        if 'sector' in focus:
            _sector = getSector(_ticker)
            if _sector in sector_lst:
                data.loc[i, 'SectorCode'] = _sector
            else:
                check_sector.append(inputlst[i])
                data.loc[i, 'SectorCode'] = _sector
                if _sector == 'N/A':
                    print('CHECK', inputlst[i], ', CANNOT GET SECTOR FROM FIDELITY!')
                else:
                    print('CHECK', inputlst[i], ',SECTOR NOT IN THE SECTOR LIST!')
        if 'Avg 10day' in focus:
            _vol10d = getVol10d(_ticker, to_date, from_date)
            data.ix[i, 'AvgVol10d'] = _vol10d
            if _vol10d == 'N/A':
                check_vol10d.append(inputlst[i])
                print('CHECK', inputlst[i], ', CANNOT GET AVGVOL10d')
        if 'Avg 3m vol' in focus or 'beta' in focus:
            _Vol3m_MktCap_Beta = getVol3m_MktCap_Beta(_ticker)
            if _Vol3m_MktCap_Beta['Vol3m'] == 'N/A' and _Vol3m_MktCap_Beta['Beta'] == 'N/A':
                _ticker_try = _ticker[:-1]+'/'+_ticker[-1:]
                _Vol3m_MktCap_Beta = getVol3m_MktCap_Beta(_ticker_try)
        if 'Avg 3m vol' in focus:
            _vol3m = _Vol3m_MktCap_Beta['Vol3m']
            data.ix[i, 'AvgVol3m'] = _vol3m
            if _vol3m == 'N/A':
                check_vol3m.append(inputlst[i])
                print('CHECK', inputlst[i], ', CANNOT GET AVGVOL3m')
        if 'Median 10day' in _focus:
            _m10d = getM10d(_ticker, to_date, from_date)
            data.ix[i, '10d_Median_vol'] = _m10d
            if _m10d == 'N/A':
                check_m10d.append(inputlst[i])
                print ('CHECK', inputlst[i], ', CANNOT GET MEDIAN10d')
        if 'Median 60day' in _focus:
            _m60d = getM60d(_ticker, to_date, add60_date)
            data.ix[i, '60d_Median_vol'] = _m60d
            if _m60d == 'N/A':
                check_m60d.append(inputlst[i])
                print('CHECK', inputlst[i], ', CANNOT GET MEDIAN60d')
        if 'beta' in focus:
            _beta = _Vol3m_MktCap_Beta['Beta']
            data.ix[i, 'Beta3Y'] = _beta
            if _beta == 'N/A':
                check_beta.append(inputlst[i])
                print ('CHECK', inputlst[i], ', CANNOT GET BETA')
        pickle_counter = i
        '''add new'''
        f = open(path_pickle + 'Universe_trade.pickle', 'wb')
        pickle.dump(pickle_counter, f)
        f.close()

        writeExcel(data, _outpath, _sheetname)
        i += 1

    res = {'data': data, 'check_sector': check_sector, 'check_vol10d': check_vol10d, 'check_vol3m': check_vol3m,
           'check_m10d': check_m10d, 'check_m60d': check_m60d, 'check_beta': check_beta, 'check_mktcap': check_mktcap}
    return res


_inputPath = '/Users/pengjiawei/Desktop/m_global/config/Price & Volume config_NEW.xlsx'
path_pickle = '/Users/pengjiawei/Desktop/m_global/project4/pickle files/'
_inputSheetname = 'Tradable Universe Update'
_userdefined = 'USERDEFINED'
_to_date = datetime.now().date()
_from_date = _to_date + (timedelta(days=-14))
_60_date = _to_date + (timedelta(days=-84))
_result = readInput(_inputPath, _inputSheetname, _userdefined)
_filepath = _result['filepath']
_sheetname = _result['sheetname']
_focus = _result['focus']
_outpath = _result['output']
_title = ['Ticker', 'Name', 'LastSale', 'MarketCap', 'ADR', 'IPOyear', 'Exchange']
if 'Avg 10day' in _focus:
    _title.append('AvgVol10d')
if 'Avg 3m vol' in _focus:
    _title.append('AvgVol3m')
if 'Median 10day' in _focus:
        _title.append('10d_Median_vol')
if 'Median 60day' in _focus:
        _title.append('60d_Median_vol')
if 'beta' in _focus:
    _title.append('Beta3Y')
if 'sector' in _focus:
    _title.append('SectorCode')
# create new excel file or read an existed file
print(_outpath)
try:
        wb = xlrd.open_workbook(_outpath)
        output_exist = 1
except IOError:
        wb = xlwt.Workbook(style_compression=2)
        output_exist = 0
if output_exist == 0:        
        _data = getData(_filepath, _sheetname, _title)
else:
        _data = getData(_outpath, _sheetname, _title)
_newdata_result = updateData(_data, _focus, _to_date, _from_date, _60_date)
_newdata = _newdata_result['data']
#print(_newdata)

_check_sector = pd.DataFrame(np.array(_newdata_result['check_sector']).T, columns=['CHECK SECTOR'])
_check_sector.to_csv('/Users/pengjiawei/Desktop/m_global/project4/Universe_output/check_sector.txt')
_check_vol10d = pd.DataFrame(np.array(_newdata_result['check_vol10d']).T, columns=['CHECK VOL10d'])
_check_vol10d.to_csv('/Users/pengjiawei/Desktop/m_global/project4/Universe_output/check_vol10d.txt')
_check_vol3m = pd.DataFrame(np.array(_newdata_result['check_vol3m']).T, columns=['CHECK VOL3M'])
_check_vol3m.to_csv('/Users/pengjiawei/Desktop/m_global/project4/Universe_output/check_vol3m.txt')
_check_m10d = pd.DataFrame(np.array(_newdata_result['check_m10d']).T, columns=['CHECK Median10d'])
_check_m10d.to_csv('/Users/pengjiawei/Desktop/m_global/project4/Universe_output/check_m10d.txt')
_check_m60d = pd.DataFrame(np.array(_newdata_result['check_m60d']).T, columns=['CHECK Median60d'])
_check_m60d.to_csv('/Users/pengjiawei/Desktop/m_global/project4/Universe_output/check_m60d.txt')
_check_beta = pd.DataFrame(np.array(_newdata_result['check_beta']).T, columns=['CHECK BETA'])
_check_beta.to_csv('/Users/pengjiawei/Desktop/m_global/project4/Universe_output/check_beta.txt')
_check_mktcap = pd.DataFrame(np.array(_newdata_result['check_mktcap']).T, columns=['CHECK MKTCAP'])
_check_sector.to_csv('/Users/pengjiawei/Desktop/m_global/project4/Universe_output/check_mktcap.txt')


wb_origin = xlrd.open_workbook(_outpath)
sheet_origin = wb_origin.sheet_by_index(0)
row_origin = sheet_origin.nrows
print("Generating the formatting outputfile")
new = xlwt.Workbook()
new_sheet = new.add_sheet('sheet')
for i in range(0,row_origin):
    temp_1 = sheet_origin.cell(i, 1).value
    new_sheet.write(i, 0, temp_1)
    temp_2 = sheet_origin.cell(i, 2).value
    new_sheet.write(i, 1, temp_2)
    temp_3 = sheet_origin.cell(i, 8).value
    new_sheet.write(i, 2, temp_3)
    temp_4 = sheet_origin.cell(i, 9).value
    new_sheet.write(i, 3, temp_4)
    temp_5 = sheet_origin.cell(i, 12).value
    new_sheet.write(i, 4, temp_5)
    temp_6 = sheet_origin.cell(i, 4).value
    new_sheet.write(i, 5, temp_6)
    temp_7 = sheet_origin.cell(i, 13).value
    new_sheet.write(i, 6, temp_7)
new.save('/Users/pengjiawei/Desktop/m_global/project4/Universe_output/equities_tradeable.csv')
xls = pd.ExcelFile('/Users/pengjiawei/Desktop/m_global/project4/Universe_output/equities_tradeable.csv')
df = xls.parse('sheet', index_col=None, na_values=['NA'])
for i in range(len(df)):
    if ',' in df['Name'][i]:
        df.loc[i, 'Name'] = ''.join(df['Name'][i].split(','))
df.to_csv('/Users/pengjiawei/Desktop/m_global/project4/Universe_output/equities_tradeable_new.csv',  sep='|', index=False)
filepath = '/Users/pengjiawei/Desktop/m_global/project4/Universe_output/equities_tradeable.csv'
if os.path.exists(filepath):
        os.remove(filepath)
print("Script finished")





