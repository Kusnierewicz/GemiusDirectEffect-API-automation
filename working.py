#!C:/Python34/python.exe -u

from urllib.request import urlopen
import csv
import time
import os
from xml.dom import minidom
import shutil, errno
from xlsxwriter.workbook import Workbook

def getApiKey(login, password):
    xmldoc = minidom.parse(urlopen("http://gdeapi.gemius.com/OpenSession.php?ignoreEmptyParams=Y&login=%s&passwd=%s" % (login, password)))
    api_key = xmldoc.getElementsByTagName("sessionID")[0].firstChild.nodeValue
    return api_key

def closeSession(key):
        close_session = "http://gdeapi.gemius.com/CloseSession.php?ignoreEmptyParams=Y&sessionID=%s" % (key)
        finalData = urlopen(close_session)
        xmldoc = minidom.parse(finalData)
        status = xmldoc.getElementsByTagName("status")[0].firstChild.nodeValue
        print ("Session has been successfuly closed.")
        return 

def searchCampaign(sessionID, matchingField, status, sortOrder):
    if status == "all":
        status = ''
    basic_www = "http://gdeapi.gemius.com/"
    command = "SearchCampaign.php?"
    core_attributes = "ignoreEmptyParams=Y&sessionID=%s" % (sessionID)
    www_end = "&matchingField=%s&status=%s&sortOrder=%s" % (matchingField, status, sortOrder)
    final_www = "%s%s%s%s" % (basic_www, command, core_attributes, www_end)
    return final_www

def searchPlaces(sessionID, campaignID, sortOrder):
    basic_www = "http://gdeapi.gemius.com/"
    command = "GetPlacementsList.php?"
    core_attributes = "ignoreEmptyParams=Y&sessionID=%s&" % (sessionID)
    www_end = "campaignID=%s&sortField=%s" % (campaignID, sortOrder)
    final_www = "%s%s%s%s" % (basic_www, command, core_attributes, www_end)
    return final_www

def getCreative(sessionID, campaignID, placementID, pattern, matchingField):
    basic_www = "http://gdeapi.gemius.com/"
    command = "SearchCreative.php?"
    core_attributes = "ignoreEmptyParams=Y&sessionID=%s&campaignID=%s&placementIDs=%s" % (sessionID, campaignID, placementID)
    www_end = "&pattern=%s&matchingField=%s" % (pattern, matchingField)
    final_www = "%s%s%s%s" % (basic_www, command, core_attributes, www_end)
    return final_www

def returnCreativeName(NodeList):
    creatives = NodeList.getElementsByTagName("creative")
    for c in creatives:
        return c.getElementsByTagName("matchingField")[0].firstChild.nodeValue

def getStats(sessionID, campaignID, dimensionID, indicatorID, placementID, timeDivision, lowerTimeUnit, upperTimeUnit):
    basic_www = "http://gdeapi.gemius.com/"
    command = "GetBasicStats.php?"
    core_attributes = "ignoreEmptyParams=Y&sessionID=%s&dimensionIDs=%s&indicatorIDs=%s&campaignIDs=%s&placementIDs=%s" % (sessionID, dimensionID, indicatorID, campaignID, placementID)
    www_end = "&timeDivision=%s&lowerTimeUnit=%s&upperTimeUnit=%s" % (timeDivision, lowerTimeUnit, upperTimeUnit)
    final_www = "%s%s%s%s" % (basic_www, command, core_attributes, www_end)
    return final_www

def printRow(file_name, col):
    with open(file_name, 'rt') as f:
         reader = csv.reader(f)
         for column in reader:
             print (column[col-1])

def actWithRow(file_name, col, action):
    with open(file_name, 'rt') as f:
         reader = csv.reader(f)
         for column in reader:
             return column[col-1]

def findAndCopy(csv, root, indicator, nextSign):
    value = None
    try:
        statisticsNumber = root.getElementsByTagName("statisticsNumber")[0].firstChild.nodeValue
        if int(statisticsNumber) > 1:
            value = 0
            for i in range(int(statisticsNumber)):
                value = value + int(root.getElementsByTagName(indicator)[i].firstChild.nodeValue)
        elif int(statisticsNumber) <= 1:
            value = root.getElementsByTagName(indicator)[0].firstChild.nodeValue
            
    except Exception:
        try:
            value = root.getElementsByTagName(indicator)[0].firstChild.nodeValue
        except Exception:
            value = "no data avalible"

    csv.write(str(value)) 
    print (indicator + ": " + str(value))  
    csv.write(nextSign)
    return

def findAndReturn(root, indicator):
    return root.getElementsByTagName(indicator)[0].firstChild.nodeValue

def createCampaignsList(api_key, matching_field, report_status, sorting, directory, campaign_list):
    xmldoc = minidom.parse(urlopen(searchCampaign(api_key, matching_field, report_status, sorting)))
    campaigns = xmldoc.getElementsByTagName("campaign")
    csv_file = open(campaign_list, 'w')
    for campaign in campaigns:
        findAndCopy(csv_file, campaign, "campaignID", ',')        
        findAndCopy(csv_file, campaign, "matchingField", '\n')
    csv_file.close()
    return

def calc_time(Time):
    text = "%s minutes" % (Time / 60.0)
    return text

def period_testing(one, two):

    test1 = (len(one) == 14)
    test2 = (len(two) == 14)
    if test1 and test2:
        try:
            test3 = one.isdigit()
        except Exception:
            test3 = False
        try:
            test4 = two.isdigit()
        except Exception:
            test4 = False
        if test3 and test4:
            a = int(one)
            b = int(two)
            diff = b - a
            return diff > 0            

def get_dates(original_period):
    period = original_period
    input_correct = False
    while not input_correct:
        if not period:
            start_date = ''
            end_date = ''
            division = "General"
            input_correct = True
        else:
            try:
                start_date = period.split(',')[0]
            except Exception:
                start_date = None
            try:
                end_date = period.split(',')[1]
            except Exception:
                end_date = None
                
            if start_date and end_date:
                if period_testing(start_date, end_date):
                    division = "Month"
                    input_correct = True
                else:
                    print("dates you have provided are wrong")
                    period = input(
                    """
                    Please provide correct custom dates or leave blank.
                    """)
            else:
                print("Did you use the coma to seperate start and end dates?")
                period = input(
                    """
                    Please provide correct custom dates or leave blank.
                    """)
    
    return division, start_date, end_date 
    

#---------------------------------------
quit = 'n'
while not quit == 'y':    
    StartTime = time.clock()
    #rootDir = os.path.dirname(os.path.realpath(__file__))
    api_login = input("Input you Gemius API login:")
    api_pass = input("Input you Gemius API password:")
    print ("Starting program at %s." % (time.ctime()))

    work_status = input(
"""
What would you like to do:
    a) download Gemius reports directly
    b) select specific campaigns
    c) download selected Gemius report from file "Selected_campaigns.csv"
    d) quit
""")

    period = input(
"""
If you would you like to set a custom time period please do it like this:

    since YYYYMMDDHHMMSS to YYYYMMDDHHMMSS.
    
Example of correct input:

    20150812000000,20150913000000

Which means 2015-08-12 to 2015-09-13. If you wouldn't like to set any custom preiod, please leave it blank and hit enter.
""")

    if work_status == 'd':
        quit = 'y'
        print ('See ya next time!')
    else:
        time_division, start_date, end_date = get_dates(period)
        print ("start date = " + start_date+ "\n", "end date = " + end_date + "\n", "time div = " + time_division+ "\n")

        time_stamp = time.strftime("%d%m%H%M%S")
    
        if work_status == "c":
            print ("Campaigns from the file 'Selected_campaigns.csv' will be retrived.")
            campaign_list = "..\\Selected_campaigns.csv"
            report_status = "Preselected"
            directory = "..\\%s_campaigns_%s\\" % (report_status, time_stamp)

        else:
            report_status = input("Specify status of campaigns to be collected (all, finished, current): ")
            directory = "..\\%s_campaigns_%s\\" % (report_status, time_stamp)
            campaign_list = "%s%s_campaigns.csv" % (directory, report_status)

        if not os.path.exists(directory):
            os.makedirs(directory)

        print ("Retriving %s campaigns from %s account" %(report_status, api_login))
        api_key = getApiKey(api_login, api_pass)

        if not work_status == 'c':
            createCampaignsList(api_key, "name", report_status, "asc", directory, campaign_list)

        if work_status == "b":
            print_dir = '\\' + directory
            print ("Information retrived correctly. The Campaign list is saved in %s directory." % (print_dir))
            file_list_status = input("Enter 'y' when you will save a final campaign list.")
            if not file_list_status == "y":
                closeSession(api_key)

        camp_ids = []
        camp_names = []

        with open(campaign_list, 'rt') as f:
             reader = csv.reader(f)
             for column in reader:
                 camp_ids.append(column[0])
                 camp_names.append(column[1])

        num = 0.0
        for camp in camp_ids:
            xmldoc = minidom.parse(urlopen(searchPlaces(api_key, camp, "name")))
            num = num + len(xmldoc.getElementsByTagName("placement"))

        i = 0
        j = 0.0
        print ("Processing...")
        for camp in camp_ids:
            
            file_name = '%s%s%s.csv' % (directory, camp_names[i], camp_ids[i])
            csv_file2 = open(file_name, 'w')
            xmldoc = minidom.parse(urlopen(searchPlaces(api_key, camp, "name")))
            placements = xmldoc.getElementsByTagName("placement")
            fresh_placements = []

            for placement in placements:
                if findAndReturn(placement, "isGdePlus",) == "Y":
                    fresh_placements.append(placement)

            for placement in fresh_placements:
                placeID = findAndReturn(placement, "placementID",)
                #---------------
                creative = None
                c = 0
                while creative == None:
                    creative = returnCreativeName(minidom.parse(urlopen(getCreative(api_key, camp, placeID, c, "name"))))
                    c += 1
                #---------------

                impressions = minidom.parse(urlopen(getStats(api_key, camp, 20, 4, placeID, time_division, start_date, end_date)))
                clicks = minidom.parse(urlopen(getStats(api_key, camp, 20, 2, placeID, time_division, start_date, end_date)))
                post_action = minidom.parse(urlopen(getStats(api_key, camp, 20, 22, placeID, time_division, start_date, end_date)))
                post_view = minidom.parse(urlopen(getStats(api_key, camp, 20, 23, placeID, time_division, start_date, end_date)))

                #---------------
                csv_file2.write(camp_names[i])

                csv_file2.write(",")
                #---------------
                #findAndCopy(csv_file2, placement, "placementID", ',')
                
                findAndCopy(csv_file2, placement, "name", ',')
                #---------------
                csv_file2.write(str(creative))        
                csv_file2.write(",")
                #---------------
                findAndCopy(csv_file2, impressions, "impressions", ',')
                findAndCopy(csv_file2, clicks, "clicks", ',')
                findAndCopy(csv_file2, post_action, "postClickActions", ',')
                findAndCopy(csv_file2, post_view, "postViewActions", '\n')
                j = j + 1
                if j / num >= 0.1:
                    j = 0.0
            csv_file2.close()
            i = i + 1

        print ("Operation completed.")
        #print_dir = '\\' + directory
        print ("Information retrived correctly. Files are saved in directory: %s" % (directory))
        closeSession(api_key)

        merged_file_name = "..//merged_" + report_status + "_campaigns_" + time_stamp + ".csv"
        excel_name = "..//merged_" + report_status + "_campaigns_" + time_stamp
        print ("Creating file: " + merged_file_name)
        fout=open(merged_file_name,"a")
          
        for fn in os.listdir(directory):
            if not fn == report_status + "_campaigns.csv":
                #print (fn)
                f = open(directory + "//" + fn)
                for line in f:
                    fout.write(line)
                f.close() 
        fout.close()
        
        workbook = Workbook(excel_name + '.xlsx')
        worksheet = workbook.add_worksheet()
        with open(merged_file_name, 'rt') as f:
            reader = csv.reader(f)
            for r, row in enumerate(reader):
                for c, col in enumerate(row):
                    worksheet.write(r, c, col)
        workbook.close()

        print ("Computation took %s" % (calc_time(time.clock() - StartTime)))
        if input('Quit?? (y/n)') == 'y':
            quit = 'y'
