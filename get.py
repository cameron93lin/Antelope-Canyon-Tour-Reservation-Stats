import requests
import re
import datetime
from dateutil.relativedelta import relativedelta
import json
import xlrd
import xlwt
from lxml import etree

type = {'[Adventurous Antelope] Prime-Time Tour 10': 'asp',
        '[Navajo Tours] Guided Sightseer\'s Tour': 'php',
        '[Antelope Canyon] Sightseer\'s Tour': 'fareharbor',
        '[Antelope Slot Canyon] Antelope Slot Canyon Scenic Tours': 'avtrax',
        '[Ken\'s Tours] General Tour': 'avtrax',
        '[Dixie Ellis] Sightseeing Tours': 'fareharbor2'}
startDate = datetime.datetime.strptime('2018-06-16', '%Y-%m-%d')
endDate = datetime.datetime.strptime('2018-06-22', '%Y-%m-%d')
result = {}
line = 0
Upper_Canyon_urls = {}
Lower_Canyon_urls = {}
siteList = ['[Adventurous Antelope] Prime-Time Tour 10']
tourTime = {'[Adventurous Antelope] Prime-Time Tour 10': ['10:15 AM', '12:30 PM']}
selectTime = tourTime
currentDate = startDate
Upper_Canyon_urls['[Adventurous Antelope] Prime-Time Tour 10'] = 'https://www.navajoantelopecanyon.com/Availability4.asp?intTourNum=10&strDate=$Date$&strTourTime=$Time$'
Upper_Canyon_urls['[Navajo Tours] Guided Sightseer\'s Tour'] = 'https://app.acuityscheduling.com/schedule.php?action=availableTimes&showSelect=0&fulldate=1&owner=11266394'
Upper_Canyon_urls['[Antelope Canyon] Sightseer\'s Tour'] = 'https://fareharbor.com/api/v1/companies/antelopecanyon/items/49363/calendar/$Month$/?allow_grouped=yes'
Upper_Canyon_urls['[Antelope Slot Canyon] Antelope Slot Canyon Scenic Tours'] = 'https://asct.avtrax.com/cgi-bin/oecgi3.exe/AvTrax?m=tour_entry&tour=SCENIC'
Lower_Canyon_urls['[Ken\'s Tours] General Tour'] = 'https://kens.avtrax.com/cgi-bin/oecgi3.exe/AvTrax?m=tour_entry&template=0'
Lower_Canyon_urls['[Dixie Ellis] Sightseeing Tours'] = 'https://fareharbor.com/api/v1/companies/antelopelowercanyon/items/25367/availabilities/date/$Date$/'

def get_select():
    print('Current support site and tour:')
    for idx, val in enumerate(siteList):
        print(str(idx+1)+':', val)
    select_sites = input('Please select site and tour (type -1 for all site search):')
    if select_sites == '-1':
        print(select_sites)
    else:
        print(int(select_sites)-1)


def type_asp(comName, url, time, date, xlstable):
    global line
    r = requests.get(url)
    pattern = re.compile(r'(\d+) Adults or Children')
    if pattern.findall(r.text)[0] != '0':
        print(date.strftime("%Y-%m-%d"), time, ' | Available: '+pattern.findall(r.text)[0])
        style = xlwt.XFStyle()
        style.num_format_str = '[$-en-US]h:mm AM/PM;@'
        xlstable.write(line, 0, comName)
        xlstable.write(line, 1, date.strftime("%Y/%m/%d"))
        xlstable.write(line, 2, datetime.datetime.strptime(time, "%I:%M %p"), style)
        xlstable.write(line, 3, pattern.findall(r.text)[0])
        line = line+1


def type_php(comName, url, date, xlstable):
    global line
    data = {'date': date.strftime("%Y-%m-%d"), 'type': '184691', 'calendar': '89085'}
    r = requests.post(url, data=data)
    pattern = re.compile(r'<label .*>(.*)</label>')
    for available_time in pattern.findall(r.text):
        print(date.strftime("%Y-%m-%d"), available_time)
        style = xlwt.XFStyle()
        style.num_format_str = '[$-en-US]h:mm AM/PM;@'
        xlstable.write(line, 0, comName)
        xlstable.write(line, 1, date.strftime("%Y/%m/%d"))
        xlstable.write(line, 2, datetime.datetime.strptime(available_time, "%I:%M%p"), style)
        line = line+1


def type_fare(comName, url, xlstable):
    global line
    r = requests.get(url)
    json_data = json.loads(r.text)
    for week in json_data['calendar']['weeks']:
        for day in week['days']:
            if day['availabilities']:
                for available_time in day['availabilities']:
                    if available_time['is_bookable']:
                        available_datetime = datetime.datetime.strptime(available_time['start_at'], '%Y-%m-%dT%H:%M:%S')
                        if available_datetime > startDate:
                            if available_datetime < endDate+datetime.timedelta(days=1):
                                print(available_datetime.strftime("%Y-%m-%d"), available_datetime.strftime("%I:%M %p"),
                                      ' | Available: ' + str(available_time['bookable_capacity']))
                                xlstable.write(line, 0, comName)
                                style = xlwt.XFStyle()
                                style.num_format_str = '[$-en-US]h:mm AM/PM;@'
                                xlstable.write(line, 1, available_datetime.strftime("%Y/%m/%d"))
                                xlstable.write(line, 2, datetime.datetime.strptime(available_datetime.strftime("%I:%M %p"), "%I:%M %p"), style)
                                xlstable.write(line, 3, available_time['bookable_capacity'])
                                line = line+1


def type_fare2(comName, url, xlstable):
    global line
    r = requests.get(url)
    json_data = json.loads(r.text)
    for available_time in json_data['availabilities']:
        if available_time['is_bookable']:
            available_datetime = datetime.datetime.strptime(available_time['start_at'], '%Y-%m-%dT%H:%M:%S')
            if available_datetime > startDate:
                if available_datetime < endDate+datetime.timedelta(days=1):
                    print(available_datetime.strftime("%Y-%m-%d"), available_datetime.strftime("%I:%M %p"),
                          ' | Available: ' + str(available_time['bookable_capacity']))
                    style = xlwt.XFStyle()
                    style.num_format_str = '[$-en-US]h:mm AM/PM;@'
                    xlstable.write(line, 0, comName)
                    xlstable.write(line, 1, available_datetime.strftime("%Y/%m/%d"))
                    xlstable.write(line, 2, datetime.datetime.strptime(available_datetime.strftime("%I:%M %p"), "%I:%M %p"), style)
                    xlstable.write(line, 3, available_time['bookable_capacity'])
                    line = line+1


def type_avtrax(comName, url, date, tour_type, xlstable):
    global line
    r = requests.get(url)
    pattern = re.compile(r'<input type="hidden" name="instance" value="(.*)" />')
    instance = pattern.findall(r.text)[0]
    data={
        # "time_out_utc": "8199",
        "m": "tour_details",
        # "max_date": "20181231",
        "tour": tour_type,
        "dep_cal": date.strftime("%d %b %Y"),
        "adults": "",
        "adults_big": "2000",
        "kids": "0",
        "instance": instance,
    }
    url = url.split('/AvTrax')[0] + '/AvTrax'
    r = requests.post(url, data=data)
    html = etree.HTML(r.text)
    for n in html.xpath("//table[@id='FlightSelect']/tr/td[1]"):
        # print(etree.tostring(n).decode("utf-8"))
        if re.search(r'Sold Out', etree.tostring(n).decode("utf-8")):
            pass
        elif re.search(r'<input type="radio"', etree.tostring(n).decode("utf-8")):
            print(date.strftime("%d %b %Y"), ' '.join(n.xpath("../td[2]/text()")[0].split()), ' | Available: '+'70')
            style = xlwt.XFStyle()
            style.num_format_str = '[$-en-US]h:mm AM/PM;@'
            xlstable.write(line, 0, comName)
            xlstable.write(line, 1, date.strftime("%d %b %Y"))
            xlstable.write(line, 2, datetime.datetime.strptime(' '.join(n.xpath("../td[2]/text()")[0].split()), "%I:%M%p"), style)
            xlstable.write(line, 3, 70)
        else:
            pattern = re.compile(r'<strong>Only (.*) available</strong>')
            print(date.strftime("%d %b %Y"), ' '.join(n.xpath("../td[2]/text()")[0].split()),
                  ' | Available: ' + pattern.findall(etree.tostring(n).decode("utf-8"))[0])
            style = xlwt.XFStyle()
            style.num_format_str = '[$-en-US]h:mm AM/PM;@'
            xlstable.write(line, 0, comName)
            xlstable.write(line, 1, date.strftime("%Y/%m/%d"))
            xlstable.write(line, 2, datetime.datetime.strptime(' '.join(n.xpath("../td[2]/text()")[0].split()), "%I:%M%p"), style)
            xlstable.write(line, 3, pattern.findall(etree.tostring(n).decode("utf-8"))[0])
            line = line+1


if __name__ == "__main__":
    # get_select()
    file = xlwt.Workbook()
    table = file.add_sheet('Upper Antelope Canyon')
    table.write(0, 0, 'Company')
    table.write(0, 1, 'Date')
    table.write(0, 2, 'Time')
    table.write(0, 3, 'Available')
    print('Upper Antelope Canyon:')
    print('**********************')
    print('')
    line = line+1
    for urlName in Upper_Canyon_urls:
        currentDate = startDate
        print(urlName)
        if type[urlName] == 'asp':
            while currentDate <= endDate:
                for time in selectTime[urlName]:
                    thisUrl = Upper_Canyon_urls[urlName].replace('$Date$', currentDate.strftime("%Y-%m-%d"))
                    thisUrl = thisUrl.replace('$Time$', time)
                    type_asp(urlName, thisUrl, time, currentDate, table)
                currentDate = currentDate + datetime.timedelta(days=1)
            print('')
        elif type[urlName] == 'php':
            while currentDate <= endDate:
                type_php(urlName, Upper_Canyon_urls[urlName], currentDate, table)
                currentDate = currentDate + datetime.timedelta(days=1)
            print('')
        elif type[urlName] == 'fareharbor':
            date = startDate
            delta = relativedelta(months=1)
            month_list = []
            while 1:
                month_list.append(date.strftime("%Y/%m"))
                date = date + delta
                if date > endDate:
                    if date.month > endDate.month:
                        break
            for month in month_list:
                thisUrl = Upper_Canyon_urls[urlName].replace('$Month$', month)
                type_fare(urlName, thisUrl, table)
        else:
            while currentDate <= endDate:
                type_avtrax(urlName, Upper_Canyon_urls[urlName], currentDate, 'SCENIC', table)
                currentDate = currentDate + datetime.timedelta(days=1)
        print('-------------------')
    table = file.add_sheet('Lower Antelope Canyon')
    table.write(0, 0, 'Company')
    table.write(0, 1, 'Date')
    table.write(0, 2, 'Time')
    table.write(0, 3, 'Available')
    print('**********************')
    print('')
    print('')
    print('Lower Antelope Canyon:')
    print('**********************')
    print('')
    line = 0
    line = line+1
    for urlName in Lower_Canyon_urls:
        currentDate = startDate
        print(urlName)
        if type[urlName] == 'avtrax':
            while currentDate <= endDate:
                type_avtrax(urlName, Lower_Canyon_urls[urlName], currentDate, 'GENERAL', table)
                currentDate = currentDate + datetime.timedelta(days=1)
        elif type[urlName] == 'fareharbor2':
            while currentDate <= endDate:
                thisUrl = Lower_Canyon_urls[urlName].replace('$Date$', currentDate.strftime("%Y-%m-%d"))
                type_fare2(urlName, thisUrl, table)
                currentDate = currentDate + datetime.timedelta(days=1)
        print('-------------------')
    file.save('result.xls')
