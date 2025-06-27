'''
--------------------------------------------------------------------------------
| No |  Date     | Version |      remarks
--------------------------------------------------------------------------------
| 1  | 01 DEC 20 | V1.0    |  Initial version
--------------------------------------------------------------------------------
| 2  | 11 DEC 20 | V1.1    |  Restructure the algorithm to improve runtime
--------------------------------------------------------------------------------
| 3  | 26 Mar 21 | V1.2    |  Add Report module
--------------------------------------------------------------------------------
| 4  | 29 Mar 21 | V1.3    |  Fix issue on get first online date and time
--------------------------------------------------------------------------------
| 5  | 06 Apr 21 | V1.4    |  Fix issue on pre-status bugs
--------------------------------------------------------------------------------
| 6  | 06 Nov 21 | V1.5    |  Fix issue on export2ls bugs
--------------------------------------------------------------------------------
| 7  | 08 Nov 21 | V1.6    |  Fix issue on export2ls bugs
--------------------------------------------------------------------------------
| 8  | 12 Jun 23 | V1.7    |  Remove Substation Type Filter ( include port to master for DCU )
--------------------------------------------------------------------------------
| 9  | 23 Jun 23 | V1.8    |  Offline criterial become !=2
--------------------------------------------------------------------------------
| 10 | 19 Jun 24 | V1.9    |  add filer based on region
--------------------------------------------------------------------------------
| 11 | 20 Jun 24 | V1.10   |  clear screen before print
--------------------------------------------------------------------------------
| 12 | 26 May 25 | V1.11   |  Improve DB search time
--------------------------------------------------------------------------------
| 13 | 27 Jun 25 | V1.12   |  Fix accurarcy
--------------------------------------------------------------------------------

'''

import sys
import mysql.connector
from mysql.connector import Error

# import cx_Oracle
import xlsxwriter
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *


class MainDialog(QMainWindow):

    def __init__(self, parent=None):
        super(MainDialog, self).__init__(parent)

        version = "Version 1.12 - MYSQL (27_JUN_2025)"
        print("version = " + version)

        ## Global Variable ##
        self.IP = "192.168.201.151"

        self.DB_CONN = object
        self.DB_CURSOR = object
        self.SUB_LIST = []
        self.DICTIONARY = {}
        self.ORCT = {}
        self.AVERAGE_AVAILABILITY = 0

        self.START_DATE = ""
        self.STOP_DATE = ""

        ### IP ##
        self.hostIpLabel = QLabel("HOST IP:", self)
        self.hostIpLabel.setGeometry(QRect(20, 20, 90, 30))

        self.hostIpEdit = QLineEdit(self.IP, self)
        self.hostIpEdit.setGeometry(QRect(120, 20, 180, 30))

        ## START ##
        self.startLabel = QLabel("START", self)
        self.startLabel.setGeometry(QRect(20, 60, 90, 30))

        self.startDateTime = QDateTimeEdit(self)
        self.startDateTime.setGeometry(QRect(120, 60, 180, 30))
        self.startDateTime.setDateTime(QDateTime(QDate(2023, 6, 22), QTime(00, 00, 00, 000)))

        ## STOP ##
        self.stopLabel = QLabel("STOP", self)
        self.stopLabel.setGeometry(QRect(20, 100, 90, 30))

        self.stopDateTime = QDateTimeEdit(self)
        self.stopDateTime.setGeometry(QRect(120, 100, 180, 30))
        self.stopDateTime.setDateTime(QDateTime(QDate.currentDate(), QTime(23, 59, 00, 000)))

        ## REGION ##
        self.RegionLabel = QLabel("By REGION", self)
        self.RegionLabel.setGeometry(QRect(440, 100, 90, 30))

        self.RegionComboBox = QComboBox(self)
        self.RegionComboBox.addItems( [ "ALL" , "WR-KCH [0 - 3999]" ,"CR-SIBU [4000 - 7999]" , "BR-BINTULU [8000 - 11999]" , "NR-MIRI [12000 - 15999]"])
        self.RegionComboBox.setGeometry(QRect(540, 100, 180, 30))



        ## RUN ##
        self.runButton = QPushButton("RUN", self)
        self.runButton.setGeometry(QRect(320, 20, 90, 50))

        ## EXPORT ##
        self.exportButton = QPushButton("EXPORT", self)
        self.exportButton.setGeometry(QRect(320, 80, 90, 50))

        ## AVERAGE LCD ##
        self.lcdNumber = QLCDNumber(self)
        self.lcdNumber.setGeometry(QRect(1090, 20, 180, 110))

        ## TABLE ##
        self.tableWidget = QTableWidget(self)
        self.tableWidget.setGeometry(QRect(20, 150, 1250, 700))
        self.headers = ["Subs", "Recorded_Online_Time", "Selected_Start_Time", "Selected_Stop_Time", "Category", "DC",
                        "Total_Offline(s)", "Total_Online(s)", "Availability"]
        self.tableWidget.clear()
        self.tableWidget.setSortingEnabled(False)
        self.tableWidget.setRowCount(0)
        self.tableWidget.setColumnCount(len(self.headers))

        self.tableWidget.setHorizontalHeaderLabels(self.headers)
        self.tableWidget.horizontalHeader().setStyleSheet("color:darkblue")
        self.tableWidget.setAutoFillBackground(True)
        self.tableWidget.resizeColumnsToContents()

        ## SIGNAL & SLOT ##
        self.runButton.clicked.connect(self.run)
        self.exportButton.clicked.connect(self.export2xlsx)

        self.setWindowTitle("KCH DSCADA DF8000 CA [" + version + "]")
        self.resize(1300, 900)

    ###################
    ## connect to DB ##
    ###################
    def connect_to_db(self):

        print("IP:" + str(self.hostIpEdit.text()))
        # dsn_tns = cx_Oracle.makedsn(str(self.hostIpEdit.text()), '1521', service_name='orcl')
        # dsn_tns = cx_Oracle.makedsn(str(self.hostIpEdit.text()), '3306', service_name='localhost')

        try:
            # conn = cx_Oracle.connect(user=r'xopens', password='ytdf000', dsn=dsn_tns)
            # conn = cx_Oracle.connect(user=r'root', password='', dsn=dsn_tns)
            connection = mysql.connector.connect(host=str(self.hostIpEdit.text()),
                                                 database='xopens',
                                                 user='xopens',
                                                 password='ytdf000'
                                                 )
            print("DEBUG connection = ", connection)
            if connection.is_connected():
                db_Info = connection.get_server_info()
                print("Connected to MySQL Server version ", db_Info)
                cursor = connection.cursor()
                cursor.execute("select database();")
                record = cursor.fetchone()
                print("You're connected to database: ", record)
                return connection, cursor

        except Error as e:
            print("Error while connecting to MySQL", e)

        # finally:
        #    if connection.is_connected():
        #        return connection, cursor

    ###################
    ## get sub list  ##
    ###################
    def get_sub_list_from_db(self, DB_cursor):
        ret_sub_list = []
        r_region = self.RegionComboBox.currentIndex()

        if r_region == 0:
            print("ALL")
            sql_command = "select * from CHNL_PARAM_TAB where chnl_use_flag = 1"
        if r_region == 1:
            print("WR-KCH [0 - 3999]")
            sql_command = "select * from CHNL_PARAM_TAB where chnl_use_flag = 1 and chnl_no >= 0 and chnl_no <= 3999"
        if r_region == 2:
            print("CR-SIBU [4000 - 7999]")
            sql_command = "select * from CHNL_PARAM_TAB where chnl_use_flag = 1 and chnl_no >= 4000 and chnl_no <= 7999"
        if r_region == 3:
            print("BR-BINTULU [8000 - 11999]")
            sql_command = "select * from CHNL_PARAM_TAB where chnl_use_flag = 1 and chnl_no >= 8000 and chnl_no <= 11999"
        if r_region == 4:
            print("NR-MIRI [12000 - 15999]")
            sql_command = "select * from CHNL_PARAM_TAB where chnl_use_flag = 1 and chnl_no >= 12000 and chnl_no <= 15999"

        print("DEBUG: r_region =" + str(r_region))

        DB_cursor.execute(sql_command)
        res = DB_cursor.fetchall()


        for s in res:
            sub_NAME = s[0]
            sub_DESCRIPTION = s[1]
            sub_TYPE = s[3]


            sub_NAME = s[0].strip().upper()
            if sub_NAME != "TEST1":
                ret_sub_list.append([sub_NAME, s[1].strip()])

        return ret_sub_list

    def get_all_event_from_db(self, DB_cursor, SUB_LIST, user_input_start_date, user_input_start_time,
                              user_input_stop_date, user_input_stop_time):
        try:
            from collections import defaultdict

            sub_codes = [s[0].strip() for s in SUB_LIST if s[0].strip().upper() != "TEST1"]
            if not sub_codes:
                print("No valid substations found.")
                return {}

            start_full = int(user_input_start_date + user_input_start_time)
            stop_full = int(user_input_stop_date + user_input_stop_time)

            in_clause = ','.join(['%s'] * len(sub_codes))
            sql = (
                "SELECT EVENT_OBJ_NAME0, YYMMDD, HHMMSSMS, CUR_STATUS, Event_Type "
                "FROM His_Event_Tab "
                "WHERE Event_Type = 1004 AND Event_Cat_No = 1 "
                f"AND EVENT_OBJ_NAME0 IN ({in_clause}) "
                "AND CONCAT(YYMMDD, LPAD(HHMMSSMS, 9, '0')) BETWEEN %s AND %s "
                "ORDER BY EVENT_OBJ_NAME0, YYMMDD, HHMMSSMS"
            )

            params = sub_codes + [start_full, stop_full]
            print(f"DEBUG SQL:\n{sql}")
            print(f"DEBUG PARAMS:\n{params}")

            DB_cursor.execute(sql, params)
            all_events = DB_cursor.fetchall()

            events_by_sub = defaultdict(list)
            for row in all_events:
                if row[0] is None or row[1] is None or row[2] is None:
                    continue
                sub = row[0].strip()
                events_by_sub[sub].append((row[1], row[2], row[3]))

            for s in SUB_LIST:
                sub_code = s[0].strip()
                des = s[1].strip()

                if sub_code.upper() == "TEST1":
                    continue

                self.DICTIONARY[sub_code] = {
                    "Subs": [f"{sub_code}-{des}"],
                    "Recorded_Online_Time": "NA",
                    "Selected_Start_Time": "NA",
                    "Selected_Stop_Time": "NA",
                    "Category": "NA",
                    "DC": "NA",
                    "Total_Offline(s)": "NA",
                    "Total_Online(s)": "NA",
                    "Availability": "NA",
                    "pre_status": "NA",
                    "event_list": [],
                }

                sub_events = events_by_sub.get(sub_code, [])
                if not sub_events:
                    continue

                online_events = [e for e in sub_events if e[2] == 2]
                if online_events:
                    earliest_online = sorted(online_events, key=lambda x: (int(x[0]), int(x[1])))[0]
                    self.DICTIONARY[sub_code]["Recorded_Online_Time"] = [earliest_online[0], earliest_online[1]]
                else:
                    continue

                first_online_date = int(self.DICTIONARY[sub_code]["Recorded_Online_Time"][0])
                first_online_time = int(self.DICTIONARY[sub_code]["Recorded_Online_Time"][1])
                start_date = int(user_input_start_date)
                start_time = int(user_input_start_time)
                stop_date = int(user_input_stop_date)
                stop_time = int(user_input_stop_time)

                start_full_int = int(user_input_start_date + user_input_start_time)
                stop_full_int = int(user_input_stop_date + user_input_stop_time)

                if (first_online_date < start_date or
                        (first_online_date == start_date and first_online_time <= start_time)):

                    self.DICTIONARY[sub_code]["Category"] = "CASE_1"
                    self.DICTIONARY[sub_code]["Selected_Start_Time"] = [user_input_start_date, user_input_start_time]
                    self.DICTIONARY[sub_code]["Selected_Stop_Time"] = [user_input_stop_date, user_input_stop_time]

                    prior_events = [e for e in sub_events if int(str(e[0]) + str(e[1]).zfill(9)) <= start_full_int]
                    if prior_events:
                        latest_event = sorted(prior_events, key=lambda x: (int(x[0]), int(x[1])), reverse=True)[0]
                        self.DICTIONARY[sub_code]["pre_status"] = str(latest_event[2])
                    else:
                        continue

                elif (first_online_date > stop_date or
                      (first_online_date == stop_date and first_online_time >= stop_time)):
                    self.DICTIONARY[sub_code]["Category"] = "CASE_3"
                    continue

                else:
                    self.DICTIONARY[sub_code]["Category"] = "CASE_2"
                    self.DICTIONARY[sub_code]["Selected_Start_Time"] = [
                        self.DICTIONARY[sub_code]["Recorded_Online_Time"][0],
                        self.DICTIONARY[sub_code]["Recorded_Online_Time"][1]]
                    self.DICTIONARY[sub_code]["Selected_Stop_Time"] = [user_input_stop_date, user_input_stop_time]
                    self.DICTIONARY[sub_code]["pre_status"] = "2"

                # Filter event list inside time window
                if self.DICTIONARY[sub_code]["Category"] in ("CASE_1", "CASE_2"):
                    self.DICTIONARY[sub_code]["event_list"] = [
                        e for e in sub_events
                        if start_full_int <= int(str(e[0]) + str(e[1]).zfill(9)) <= stop_full_int
                    ]

            return self.DICTIONARY

        except Error as e:
            print("Error while fetching event data", e)
            return {}

    #################################
    ## get_oracle_date_time_format ##
    #################################
    def get_oracle_date_time_format(self, q, type):

        q_str = q.toString(Qt.ISODate)
        # 2020-11-28T00:00:00
        # 2020-11-28T23:59:59
        year = q_str.split("T")[0].split("-")[0]
        month = q_str.split("T")[0].split("-")[1]
        day = q_str.split("T")[0].split("-")[2]
        hour = q_str.split("T")[1].split(":")[0]
        minute = q_str.split("T")[1].split(":")[1]

        second = "00"
        msecond = "000"

        ret_date = year + month + day
        ret_time = hour + minute + second + msecond

        if type == "date":
            return ret_date
        if type == "time":
            return ret_time

    ############################################
    ## get_total_sec_from_orcl_time_range_int ##
    ############################################
    def get_total_sec_from_orcl_time_range_int(self, start_str_list, stop_str_list):

        totalSec = 0

        start_date_str = str(start_str_list[0])
        start_time_str = str(start_str_list[1])
        if len(start_time_str) < 9:
            start_time_str = "0" * (9 - len(start_time_str)) + start_time_str

        stop_date_str = str(stop_str_list[0])
        stop_time_str = str(stop_str_list[1])
        if len(stop_time_str) < 9:
            stop_time_str = "0" * (9 - len(stop_time_str)) + stop_time_str

        # print("DEBUG 0: " + str(start_str_list) + "-" + str(stop_str_list) )
        f_year = int(start_date_str[:4])
        f_month = int(start_date_str[4:6])
        f_day = int(start_date_str[6:8])
        f_hour = int(start_time_str[:2])
        f_min = int(start_time_str[2:4])
        f_sec = int(start_time_str[4:6])
        f_msec = int(start_time_str[6:9])
        fromQtime = QDateTime(QDate(f_year, f_month, f_day), QTime(f_hour, f_min, f_sec, f_msec))

        t_year = int(stop_date_str[:4])
        t_month = int(stop_date_str[4:6])
        t_day = int(stop_date_str[6:8])
        t_hour = int(stop_time_str[:2])
        t_min = int(stop_time_str[2:4])
        t_sec = int(stop_time_str[4:6])
        t_msec = int(stop_time_str[6:9])
        toQtime = QDateTime(QDate(t_year, t_month, t_day), QTime(t_hour, t_min, t_sec, t_msec))

        totalSec = fromQtime.secsTo(toQtime)
        # print("DEBUG 1: " + str(start_str_list) + "-" + str(stop_str_list) + "==>" + str(totalSec))
        return totalSec

    def getOfflineInterval(self, off_time, on_time):

        retval = 0

        off_msec = off_time[-3:]
        if off_msec == "":
            off_msec = 0
        else:
            off_msec = int(off_msec) / 1000

        off_sec = off_time[-5:-3]
        if off_sec == "":
            off_sec = 0
        else:
            off_sec = int(off_sec)

        off_min = off_time[-7:-5]
        if off_min == "":
            off_min = 0
        else:
            off_min = int(off_min) * 60

        off_hour = off_time[-9:-7]
        if off_hour == "":
            off_hour = 0
        else:
            off_hour = int(off_hour) * 3600

        on_msec = on_time[-3:]
        if on_msec == "":
            on_msec = 0
        else:
            on_msec = int(on_msec) / 1000

        on_sec = on_time[-5:-3]
        if on_sec == "":
            on_sec = 0
        else:
            on_sec = int(on_sec)

        on_min = on_time[-7:-5]
        if on_min == "":
            on_min = 0
        else:
            on_min = int(on_min) * 60

        on_hour = on_time[-9:-7]
        if on_hour == "":
            on_hour = 0
        else:
            on_hour = int(on_hour) * 3600

        retval = (on_hour + on_min + on_sec + on_msec) - (off_msec + off_sec + off_min + off_hour)

        return (retval)

    ###########
    ### RUN ###
    ###########
    def run(self):

        self.DICTIONARY = {}
        ### 1. connect DB ###
        self.DB_CONN, self.DB_CURSOR = self.connect_to_db()
        print("\nstage 1 [connect DB] -> ok ")

        ### 2. get substation list ###
        # ==> [ [code , des] . [code , des] , [code , des]  ]
        print("time = " + str(QDateTime.currentDateTime()))
        if not self.DB_CONN == False or not self.DB_CURSOR == False:
            self.SUB_LIST = self.get_sub_list_from_db(self.DB_CURSOR)
        print("stage 2 [get substation list]-> ok (" + str(len(self.SUB_LIST)) + ")")

        ### 3.0. get all_event list ###
        #
        print("time = " + str(QDateTime.currentDateTime()))
        if not self.DB_CONN == False or not self.DB_CURSOR == False:
            user_input_start_date = self.get_oracle_date_time_format(self.startDateTime.dateTime(), "date")
            user_input_start_time = self.get_oracle_date_time_format(self.startDateTime.dateTime(), "time")
            user_input_stop_date = self.get_oracle_date_time_format(self.stopDateTime.dateTime(), "date")
            user_input_stop_time = self.get_oracle_date_time_format(self.stopDateTime.dateTime(), "time")

            self.ORCT = self.get_all_event_from_db(self.DB_CURSOR, self.SUB_LIST, user_input_start_date,
                                                   user_input_start_time, user_input_stop_date, user_input_stop_time)

        print("stage 3.0 [get_all_event_from_db]-> ok ")

        ### 4. Analysis ####
        ### 4. Analysis ####
        print("time = " + str(QDateTime.currentDateTime()))
        self.AVERAGE_AVAILABILITY = 0.0

        for sub_list in self.SUB_LIST:
            sub_code = sub_list[0].strip()

            # Safety checks
            if sub_code not in self.DICTIONARY:
                print(f"Skipping {sub_code}: not in DICTIONARY")
                continue

            sub_data = self.DICTIONARY[sub_code]

            if (sub_data["Recorded_Online_Time"] == "NA" or
                    sub_data["Category"] == "CASE_3" or
                    sub_data["pre_status"] == "NA" or
                    not isinstance(sub_data["event_list"], list)):
                print(f"Skipping {sub_code}: invalid data")
                continue

            start_flag = sub_data["pre_status"]
            window_event_list = sub_data["event_list"]

            if len(window_event_list) == 0:
                if start_flag == "2":
                    sub_data["DC"] = 0
                    sub_data["Total_Offline(s)"] = 0
                    sub_data["Total_Online(s)"] = self.get_total_sec_from_orcl_time_range_int(
                        sub_data["Selected_Start_Time"], sub_data["Selected_Stop_Time"])
                    sub_data["Availability"] = 100.0

                elif start_flag in ("3", "6"):
                    sub_data["DC"] = 1
                    sub_data["Total_Offline(s)"] = self.get_total_sec_from_orcl_time_range_int(
                        sub_data["Selected_Start_Time"], sub_data["Selected_Stop_Time"])
                    sub_data["Total_Online(s)"] = 0
                    sub_data["Availability"] = 0.0

            else:
                cur_flag_int = int(start_flag)
                cur_date_str = str(sub_data["Selected_Start_Time"][0])
                cur_time_str = str(sub_data["Selected_Start_Time"][1])
                dc_count = 0
                offline_sec = 0

                for event in window_event_list:
                    try:
                        next_flag_int = int(event[2])
                        next_date_str = str(event[0])
                        next_time_str = str(event[1])

                        # offline → online transition (cur_flag 3/6 → 2)
                        if cur_flag_int - next_flag_int == 1 or cur_flag_int - next_flag_int == 4:
                            dc_count += 1
                            offline_sec += self.get_total_sec_from_orcl_time_range_int(
                                [cur_date_str, cur_time_str], [next_date_str, next_time_str])

                        cur_flag_int = next_flag_int
                        cur_date_str = next_date_str
                        cur_time_str = next_time_str
                    except Exception as e:
                        print(f"Error parsing event for {sub_code}: {e}")
                        continue

                # Last state is offline → add remaining time
                if cur_flag_int in (3, 6):
                    next_date_str = str(sub_data["Selected_Stop_Time"][0])
                    next_time_str = str(sub_data["Selected_Stop_Time"][1])
                    dc_count += 1
                    offline_sec += self.get_total_sec_from_orcl_time_range_int(
                        [cur_date_str, cur_time_str], [next_date_str, next_time_str])

                sub_data["DC"] = dc_count
                sub_data["Total_Offline(s)"] = offline_sec

                total_sec = self.get_total_sec_from_orcl_time_range_int(
                    sub_data["Selected_Start_Time"], sub_data["Selected_Stop_Time"])
                online_sec = total_sec - offline_sec
                sub_data["Total_Online(s)"] = online_sec

                if total_sec > 0:
                    availability = (online_sec / total_sec) * 100
                    sub_data["Availability"] = "{:.2f}".format(availability)

        # Average availability
        if not self.DB_CONN == False or not self.DB_CURSOR == False:
            count = 0
            avg = 0.0
            for sub_code in self.DICTIONARY:
                if self.DICTIONARY[sub_code]["Availability"] != "NA":
                    count += 1
                    avg += float(self.DICTIONARY[sub_code]["Availability"])
            if count > 0:
                self.AVERAGE_AVAILABILITY = avg / count

        print("stage 4 [Analysis] -> ok")

        ## 5 Polulate Table ##
        print("time = " + str(QDateTime.currentDateTime()))
        ## RESET TABLE ##
        self.tableWidget.clear()
        self.tableWidget.setSortingEnabled(False)
        self.tableWidget.setRowCount(0)
        self.tableWidget.setColumnCount(len(self.headers))
        self.tableWidget.setHorizontalHeaderLabels(self.headers)
        self.tableWidget.horizontalHeader().setStyleSheet("color:darkblue")
        self.tableWidget.setAutoFillBackground(True)
        self.tableWidget.resizeColumnsToContents()

        self.tableWidget.setRowCount(len(self.DICTIONARY))
        r = 0
        for sub_code in self.DICTIONARY:
            c = 0
            for i in self.DICTIONARY[sub_code]:

                if (i == "Recorded_Online_Time" or i == "Selected_Start_Time" or i == "Selected_Stop_Time") and str(
                        self.DICTIONARY[sub_code][i]) != "NA":

                    i_date = str(self.DICTIONARY[sub_code][i][0])
                    i_time = str(self.DICTIONARY[sub_code][i][1])
                    i_date_format = i_date[:4] + "-" + i_date[4:6] + "-" + i_date[6:8]
                    if len(i_time) == 8:
                        i_time_format = i_time[:1] + ":" + i_time[1:3] + ":" + i_time[3:5] + "." + i_time[5:8]
                    else:
                        i_time_format = i_time[:2] + ":" + i_time[2:4] + ":" + i_time[4:6] + "." + i_time[6:9]
                    item = QTableWidgetItem(str([i_date_format, i_time_format]))
                    self.tableWidget.setItem(r, c, item)

                elif i == "Availability" and str(self.DICTIONARY[sub_code][i]) != "NA":
                    item = QTableWidgetItem(str(self.DICTIONARY[sub_code][i]))
                    if float(self.DICTIONARY[sub_code][i]) >= 99.7:
                        item.setBackground(QColor("lightgreen"))
                    if float(self.DICTIONARY[sub_code][i]) < 99.7:
                        item.setBackground(QColor("magenta"))

                    self.tableWidget.setItem(r, c, item)


                else:

                    item = QTableWidgetItem(str(self.DICTIONARY[sub_code][i]))
                    # print("r=" + str(r) + " c=" + str(c) + " item =" + str(item))
                    self.tableWidget.setItem(r, c, item)
                c = c + 1
            r = r + 1

        ## Display Average ##

        self.lcdNumber.display(self.AVERAGE_AVAILABILITY)
        if self.AVERAGE_AVAILABILITY >= 99.7:
            self.lcdNumber.setStyleSheet("QLCDNumber { background-color: lightgreen }")
        if self.AVERAGE_AVAILABILITY < 99.7:
            self.lcdNumber.setStyleSheet("QLCDNumber { background-color: magenta }")

        self.tableWidget.resizeColumnsToContents()
        print("stage 5 [Populate Table] -> ok")

        ## Export to excel File ##
        # print("stage 6 [Export] -> In Progress!!")

        if not self.DB_CONN == False or not self.DB_CURSOR == False:
            self.DB_CURSOR.close()
            self.DB_CONN.close()

        QMessageBox.information(self, "Info", "Program End")
        print("Program End !")
        print("time = " + str(QDateTime.currentDateTime()))

    def export2xlsx(self):
        try:
            filename = QFileDialog.getSaveFileName(self, 'Save Report', '', "Excel Files (*.xlsx)")[0]
            if not filename:
                return

            if not filename.endswith(".xlsx"):
                filename += ".xlsx"

            workbook = xlsxwriter.Workbook(filename, {'in_memory': True})
            summary_ws = workbook.add_worksheet("Summary")

            # Define mapping between headers and actual dictionary keys
            header_key_map = {
                "Subs": "Subs",
                "Recorded_Online_Time": "Recorded_Online_Time",
                "Selected_Start_Time": "Selected_Start_Time",
                "Selected_Stop_Time": "Selected_Stop_Time",
                "Category": "Category",
                "DC": "DC",
                "Total_Offline(s)": "Total_Offline(s)",
                "Total_Online(s)": "Total_Online(s)",
                "Availability": "Availability"
            }

            # Write headers
            for col, header in enumerate(self.headers):
                summary_ws.write(0, col, header)

            # Write data
            for row_idx, sub_code in enumerate(self.DICTIONARY, start=1):
                data = self.DICTIONARY[sub_code]
                for col_idx, header in enumerate(self.headers):
                    dict_key = header_key_map.get(header)
                    value = data.get(dict_key, "NA") if dict_key else "NA"

                    if isinstance(value, list):
                        try:
                            # If datetime format: [yyyymmdd, hhmmssms]
                            if len(value) == 2 and str(value[0]).isdigit() and str(value[1]).isdigit():
                                date_str = str(value[0])
                                time_str = str(value[1]).zfill(9)
                                formatted = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]} {time_str[:2]}:{time_str[2:4]}:{time_str[4:6]}.{time_str[6:]}"
                                summary_ws.write(row_idx, col_idx, formatted)
                            else:
                                # If list like ['M001-DESCRIPTION']
                                summary_ws.write(row_idx, col_idx, ", ".join(str(v) for v in value))
                        except Exception:
                            summary_ws.write(row_idx, col_idx, str(value))
                    else:
                        summary_ws.write(row_idx, col_idx, str(value))

            # Add event list worksheets
            for sub_code in self.DICTIONARY:
                event_list = self.DICTIONARY[sub_code].get("event_list", "NA")
                if event_list == "NA" or not isinstance(event_list, list) or len(event_list) == 0:
                    continue

                safe_sheet_name = sub_code[:31].replace(" ", "_")  # Excel sheet name limit
                try:
                    ws = workbook.add_worksheet(safe_sheet_name)
                except:
                    print(f"Warning: Failed to create worksheet for {sub_code}")
                    continue

                ws.write_row(0, 0, ["No", "Date", "Time", "Status"])
                for i, e in enumerate(event_list, start=1):
                    try:
                        yyyymmdd = str(e[0])
                        hhmmssms = str(e[1]).zfill(9)
                        cur_status = int(e[2])

                        date_fmt = f"{yyyymmdd[:4]}-{yyyymmdd[4:6]}-{yyyymmdd[6:]}"
                        time_fmt = f"{hhmmssms[:2]}:{hhmmssms[2:4]}:{hhmmssms[4:6]}.{hhmmssms[6:]}"
                        status = "online" if cur_status == 2 else "offline"

                        ws.write(i, 0, i)
                        ws.write(i, 1, date_fmt)
                        ws.write(i, 2, time_fmt)
                        ws.write(i, 3, status)
                    except Exception as ex:
                        print(f"Error writing event row {i} for {sub_code}: {ex}")

            workbook.close()
            QMessageBox.information(self, "Info", "Export to Excel completed!")

        except xlsxwriter.exceptions.FileCreateError as e:
            QMessageBox.warning(self, "Export Error", f"File could not be saved:\n{e}")
        except Exception as ex:
            QMessageBox.critical(self, "Unexpected Error", f"An error occurred:\n{ex}")
            import traceback
            traceback.print_exc()


app = QApplication(sys.argv)
dialog = MainDialog()
dialog.show()
app.exec_()