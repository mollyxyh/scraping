import pandas as pd
import urllib.request as urllib2
import datetime as dt
import time


def google_intraday(ticker, period, window):

    print('Current ticker: ' + ticker)

    url_root = 'http://www.google.com/finance/getprices?i='
    url_root += str(period) + '&p=' + str(window)
    url_root += 'd&f=d,o,h,l,c,v&df=cpct&q=' + ticker  # + '&x=TYO'
    print(url_root)

    response = urllib2.urlopen(url_root)
    data = response.read()
    encoding = response.info().get_content_charset('utf-8')
    data = data.decode(encoding).split('\n')

    # actual data starts at index = 7
    # first line contains full timestamp, every other line is offset of period from timestamp
    parsed_data = []
    anchor_stamp = ''
    end = len(data)
    # check whether data is available from GOOGLE FINANCE
    if end == 7:
        parsed_data.append((None, None, None, None, None, None, None))
    else:
        for j in range(7, end-1):
            cdata = data[j].split(',')
            if 'TIMEZONE_OFFSET' in cdata[0]:
                continue
            elif 'a' in cdata[0]:
                # first one record anchor timestamp
                anchor_stamp = cdata[0].replace('a', '')
                cts = int(anchor_stamp)
                parsed_data.append((str(dt.datetime.fromtimestamp(float(cts)))[0:10],
                                    str(dt.datetime.fromtimestamp(float(cts)))[11:16], float(cdata[1]),
                                    float(cdata[2]), float(cdata[3]), float(cdata[4]), float(cdata[5])))
            else:
                try:
                    coffset = int(cdata[0])
                    cts = int(anchor_stamp) + (coffset * period)
                    parsed_data.append((str(dt.datetime.fromtimestamp(float(cts)))[0:10],
                                        str(dt.datetime.fromtimestamp(float(cts)))[11:16], float(cdata[1]),
                                        float(cdata[2]), float(cdata[3]), float(cdata[4]), float(cdata[5])))
                except Exception as e:
                    print(e,)  # for time zone offsets thrown into data)
                    print(' ' + str(j) + ' ' + str(end))
    df = pd.DataFrame(data=parsed_data)
    df.columns = ['date', 'time', 'close', 'high', 'low', 'open', 'volume']
    df.index = df.date
    del df['date']

    return df


def yahoo_intraday(ticker, window):

    print('Current ticker: ' + ticker)

    url_root = 'http://chartapi.finance.yahoo.com/instrument/1.0/'
    url_root += ticker + '/chartdata;type=quote;range=' + str(window) + 'd/csv'
    print(url_root)

    response = urllib2.urlopen(url_root)
    data = response.read()
    encoding = response.info().get_content_charset('utf-8')
    data = data.decode(encoding).split('\n')

    # actual data starts at index = 17
    # first line contains full timestamp, every other line is offset of period from timestamp
    parsed_data = []
    end = len(data)
    # check whether data is available from GOOGLE FINANCE
    if end == 4:
        parsed_data.append((None, None, None, None, None, None, None))
    else:
        if window == 1:
            line = 0
        else:
            line = window
        for j in range(17 + line, end-1):
            cdata = data[j].split(',')
            try:
                stamp = int(cdata[0])
                parsed_data.append((str(dt.datetime.fromtimestamp(float(stamp)))[0:10],
                                    str(dt.datetime.fromtimestamp(float(stamp)))[11:16], float(cdata[1]),
                                    float(cdata[2]), float(cdata[3]), float(cdata[4]), float(cdata[5])))
            except Exception as e:
                print(e,)  # for time zone offsets thrown into data)
                print(' ' + str(j) + ' ' + str(end))
    df = pd.DataFrame(data=parsed_data)
    df.columns = ['date', 'time', 'close', 'high', 'low', 'open', 'volume']
    df.index = df.date
    del df['date']

    return df


def timestr2timestamp(timestr):
    timestr = dt.datetime.strptime(timestr, '%Y-%m-%d %H:%M:%S').timetuple()
    stamp = int(time.mktime(timestr))

    return stamp


def timestamp2timestr(stamp):
    return str(dt.datetime.fromtimestamp(int(stamp)))


#print(yahoo_intraday('A', 1))
a = google_intraday('AAPL', 60, 10).loc['2016-10-27']
print(a)
#print(timestr2timestamp('2016-10-25 09:30:00'))
#print(timestamp2timestr(1477402200))
