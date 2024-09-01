import serial
import time
from datetime import datetime
import subprocess
import sqlite3
from sqlite3 import Error
from win32wifi import Win32Wifi as ww
import re

db_file = "c:\\users\\cbben\\onedrive\\documents\\github\\pythonGPSCom\\wifidata.sqlite"
comport = "COM4"

def create_connection(db_file):
    conn = None
    try:
        conn = sqlite3.connect(db_file)
    except Error as e:
        print(e)
    return conn

def get_gpspos():
    while (1):
        # Wait until there is data waiting in the serial buffer
        if (ser.in_waiting > 0):

            # GPS Operations - Read data out of the buffer until a carraige return / new line is found
            serialString = ser.readline()

            # GPS Operations - Print the contents of the serial data
            if re.match("^\$GPRMC", serialString.decode('Ascii')):
                gpsdataln = serialString.decode('Ascii')
                msgid, utctime, status, lat, ns_ind, long, ew_ind, speed, course, date, magvar, mode, chksum = gpsdataln.split(
                    ",")
                latdd = int(float(lat) / 100)
                latmm = float(lat) - latdd * 100
                latdec = latdd + latmm / 60
                if (ns_ind == "S"):
                    latdec = latdec * -1
                longdd = int(float(long) / 100)
                longmm = float(long) - longdd * 100
                longdec = longdd + longmm / 60
                if (ew_ind == "W"):
                    longdec = longdec * -1
                timehhmmss, fracsec = utctime.split(".")
                timehhmmss = ':'.join(re.findall('..', timehhmmss))
                date = '-'.join(re.findall('..', date))
                datetimeraw = date + "T" + str(timehhmmss) + "." + str(fracsec)
                datetimeobj = datetime.strptime(datetimeraw, '%d-%m-%yT%H:%M:%S.%f')
                return datetimeobj,lat,ns_ind, long, ew_ind, latdec, longdec
                break

def wifiscan():
    interfaces = ww.getWirelessInterfaces()
    for val in interfaces:
        #print (val.guid)
        handle = ww.WlanOpenHandle()
        scan = ww.WlanScan(handle,val.guid)
        ww.WlanCloseHandle(handle)

def getnetworks():
    ##print(str(datetimeobj),lat+ns_ind,long+ew_ind,latdec,longdec)
    winwifilst = subprocess.check_output(["netsh", "wlan", "show", "networks", "mode=bssid"])
    winwifilst = winwifilst.decode("ascii")  # needed in python 3
    winwifilst = winwifilst.replace("\r", "")
    winwifilstitem = winwifilst.split("\n")
    #print (winwifilst)
    for line in winwifilstitem:
                if (re.match("^SSID",line)):
                    SSIDName= line.split(" : ")
                    SSIDName = str(SSIDName[1])
                elif(re.match("\s+Network type.+:",line)):
                    NetType = line.split(" : ")
                    NetType = str(NetType[1])
                elif(re.match("\s+Authentication.+:",line)):
                    Authtype = line.split(" : ")
                    AuthType = str(Authtype[1])
                elif(re.match("\s+Encryption.+:",line)):
                    Enctype = line.split(" : ")
                    EncType = str(Enctype[1])
                elif(re.match("\s+BSSID.+:",line)):
                    BSMac = line.split(" : ")
                    BSNum = BSMac[0].strip()
                    BSNum = BSNum.split(" ")
                    BSNum = BSNum[1]
                    BSMac = str(BSMac[1])
                elif(re.match("\s+Signal.+:",line)):
                    Signal = line.split(" : ")
                    Signal = str(Signal[1])
                elif(re.match("\s+Radio\s+type.+:",line)):
                    RadioType = line.split(" : ")
                    RadioType = str(RadioType[1])
                elif(re.match("\s+Band.+:",line)):
                    RadioBand = line.split(" : ")
                    RadioBand = str(RadioBand[1])
                elif(re.match("\s+Channel\s+:\s+\d+",line)):
                    Channel = line.split(" : ")
                    Channel = str(Channel[1])
                elif(re.match("\s+Basic\s+rates.*:\s+",line)):
                    BRates= line.split(" : ")
                    BRates = str(BRates[1])
                elif(re.match("\s+Other\s+rates.*:\s+",line)):
                    ORates= line.split(" : ")
                    ORates = str(ORates[1])
                    newrecord = (str(datetimeobj), lat + ns_ind, long + ew_ind, latdec, longdec, SSIDName, NetType, AuthType, EncType, BSNum, BSMac, Signal, RadioType,RadioBand,Channel, BRates, ORates)
                    sql = ''' INSERT INTO WifiData(datetime,lat1,lon1,lat2,lon2,ssidname,nettype,authtype,enctype,basenum,basemac,signalstrength,radiotype,radioband,channel,brateslst,orateslst)
                              VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) '''
                    cur = conn.cursor()
                    cur.execute(sql, newrecord)
                    conn.commit()

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    ser = serial.Serial('COM4', 9600, timeout=0,parity=serial.PARITY_EVEN, rtscts=1)
    conn = create_connection(db_file)
    locserial = 0

    while(1):
        #grab current gps location and UTC timestamp
        datetimeobj,lat,ns_ind, long, ew_ind, latdec, longdec = get_gpspos()
        locserialnew = latdec+longdec
        if (locserialnew != locserial):
            #force the device to rescan and repopular nearby network data
            wifiscan()
            #getcurrent networks using netsh command
            getnetworks()
            #set locserial to revent writing to the db when you aren't moving (duplicate survey)
            locserial = locserialnew
            #take a moment before scanning again
            time.sleep(30)

#GPS Message Formats https://www.redhat.com/architect/architects-guide-gps-and-gps-data-formats#:~:text=Bonus%20Section%3A%20The%20details%20of%20the%20GPS%20sentence,format%20similar%20to%20GPVTG.%20...%205%20GPGSA.%20