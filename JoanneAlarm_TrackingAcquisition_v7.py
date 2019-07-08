import sys
sys.path.append("C:\\opencv\\build\\python\\2.7")
import copy

import numpy as np
import cv2
import math
import time
from time import strftime
import os
from numpy import array
import xlwt
import xlrd
import datetime
import serial

def nothing(x):
    pass

connected = False
ser = serial.Serial("COM5", 9600, writeTimeout = 0)
#ser = serial.Serial("COM16", 9600, writeTimeout = 0)
while connected == False:
    print 'Initialising Arduino...'
    ser_in = ser.read()
    connected = True

RGB = 0 # 1 stands for color video and 0 stands for B/W video

##Experiment parameters
lastknownframerate= 30 #Choose from the following 10, 20, 30
total_experiment_time = 20 # in seconds
stimulus_on1 = 300 # Valve1 ON In seconds
stimulus_off1 = 303 # Valve1 OFF in seconds

stimulus_on2 = 310 # Valve2 ON In seconds
stimulus_off2 = 313 # Valve2 OFF in seconds

FrameWidth = 1680  # 1680 for Basler camera and 1280 for Logitech camera
FrameHeight = 480   # 480q
TankWidth = 200 # in mm
SearchSize= 80 # ROI size to track fish (30 for Basler camera and 20 for Logitech)
## Tank1 ======================================
## Details for saving file
# Write as valveA_valveB if different
FishName1 = 'Fish25'
StimulusType1 = 'B2_0.5'
FishType1 = 'TLEK_0818'
FishSex1 = 'Male'
SS_made_date1 = '18122018'
#=============================================

## Tank2 ======================================
## Details for saving file
# Write as valveA_valveB if different
FishName2 = 'Fish26'
StimulusType2 = 'B2_0.5'
FishType2 = 'TLEK_0818'
FishSex2 = 'Female'
SS_made_date2 = '18122018'
#=============================================



total= lastknownframerate*total_experiment_time;  # Total number of frames. The frame rate right now is 9.91 fps


stamp = np.zeros((100,200,3))

GroupName = raw_input("Enter Group Name: ")
cv2.namedWindow('frame')


# Capture image wirh specified size and frame rate
cap = cv2.VideoCapture(0)

# For the fourcc to work, add C:\opencv\build\x86\vc10\bin to System PATH
cap.set(6 ,cv2.cv.CV_FOURCC('M', 'J', 'P', 'G') );
fourcc = cv2.cv.CV_FOURCC('M','J','P','G')

cap.set(3, FrameWidth) # change Width
cap.set(4, FrameHeight) # change Height

# Setup SimpleBlobDetector parameters
params = cv2.SimpleBlobDetector_Params()
params.minDistBetweenBlobs = 10; #100

params.minArea = 20;  # The minimal pixels that show a difference  #30
params.maxArea = 100;
if RGB == 1:
    params.minThreshold = 15;  # The threshold of the pixel color difference
else:
    params.minThreshold = 50;  #  40 for Basler
params.filterByCircularity = False;
params.filterByConvexity = False;
params.filterByInertia = False;
params.filterByColor = False;
params.filterByArea = True;

detector = cv2.SimpleBlobDetector(params)
# ============================================
# Blob detection
detector = cv2.SimpleBlobDetector(params)

ii = 0
i = 0
tt4=''
tt6=list(xrange(total+500))
Font = cv2.FONT_HERSHEY_SIMPLEX
dummy= 0
preset=0

data1=[];
data1Stamp=[];
oldL1=0
small1=0
tt81=''
tt811=0
tt811a=0
tt811b=0
tt811c=0

data2=[];
data2Stamp=[];
temp1Y1=0
temp1Y2=0
temp1X1=0
temp1X2=0
temp2Y1=0
temp2Y2=0
temp2X1=0
temp2X2=0

oldL2=0
oldX2=0
oldY2=0
small2=0
tt82=''
tt812=0
tt812a=0
tt812b=0
tt812c=0

save1=0
setupbox = 0
ValveFlag = 0
keypoints = []
StimulusOnMinute1=0
StimulusOnSecond1=0
StimulusOffMinute1=0
StimulusOffSecond1=0

StimulusOnMinute2=0
StimulusOnSecond2=0
StimulusOffMinute2=0
StimulusOffSecond2=0

setframerate = 1/float(lastknownframerate)

tt81 = "C:\\Users\\Administrator\\Desktop\\AlarmVideo\\" + GroupName
tt82 = "C:\\Users\\Administrator\\Desktop\\AlarmVideo\\" + GroupName
tt83 = "C:\\Users\\Administrator\\Desktop\\AlarmVideo\\" + GroupName
tt84 = "C:\\Users\\Administrator\\Desktop\\AlarmVideo\\" + GroupName



#  Below is the mouse evemt
boxes=[]
drawing = False
mode = 0
fourcc = cv2.cv.CV_FOURCC('M','J','P','G')


def set_up_ROIs(event,x,y,flags,param):
    if len(boxes) <= 3:
        global boxnumber, drawing, img, gray2, arr, arr1, arr2, arr3, dummy, mode, tank1a, tank2a, tank3a, tank4a, StimulusOnMinute, StimulusOffMinute, StimulusOnSecond, StimulusOffSecond, OldY1, OldX1
        boxnumner = 0
        if event == cv2.EVENT_LBUTTONDOWN:
            drawing = True
            sbox = [x, y]
            boxes.append(sbox)
        elif event == cv2.EVENT_MOUSEMOVE:
            if drawing == True:
                if len(boxes) == 1:
                    arr = array(gray5)
                    cv2.rectangle(arr,(boxes[-1][0],boxes[-1][-1]),(x, y),(0,0,255),1)
                    cv2.imshow('frame',arr)
                    stamp = np.zeros((100,200,3))
                elif len(boxes) == 3:
                    arr1 = array(gray5)
                    cv2.rectangle(arr1,(boxes[-1][0],boxes[-1][-1]),(x, y),(0,0,255),1)
                    cv2.imshow('frame',arr1)
        elif event == cv2.EVENT_LBUTTONUP:
            drawing = not drawing
            ebox = [x, y]
            boxes.append(ebox)
            cv2.rectangle(gray5,(boxes[-2][-2],boxes[-2][-1]),(boxes[-1][-2],boxes[-1][-1]), (0,0,255),1)
            cv2.imshow('frame',gray5)
            if len(boxes) == 2:
                if not os.path.exists(tt81):
                    os.makedirs(tt81)
            elif len(boxes) == 4:
                if not os.path.exists(tt82):
                    os.makedirs(tt82)
            if len(boxes) == 4:
                mode = 1
                dummy = 1
                i=0

# ====================================================
oldX1=0
oldY1=0
kernel = np.ones((5, 5), np.uint8)
fgbg = cv2.BackgroundSubtractorMOG2()

ret, gray4 = cap.read()
gray4 = np.float32(gray4)

while(i<=(total-1+preset)):
    ret, frame = cap.read()
    
    gray2 = frame [:]
    gray6 = copy.deepcopy(frame)
    #gray6 = frame [:]

    if dummy == 1:
        if setupbox == 0:
            if total_experiment_time < 3600:
                StimulusOnMinute1 = int(strftime("%M",)) + (stimulus_on1 / 60)
                StimulusOnSecond1 = int(strftime("%S",)) + (stimulus_on1 % 60)
                if StimulusOnSecond1 >=60:
                    StimulusOnMinute1 = StimulusOnMinute1 + 1
                    StimulusOnSecond1 = StimulusOnSecond1 % 60
                if StimulusOnMinute1 >=60:
                    StimulusOnMinute1 = StimulusOnMinute1 % 60
                StimulusOffMinute1 = int(strftime("%M",)) + (stimulus_off1 / 60)
                StimulusOffSecond1 = int(strftime("%S",)) + (stimulus_off1 % 60)
                if StimulusOffSecond1 >=60:
                    StimulusOffMinute1 = StimulusOffMinute1 +1
                    StimulusOffSecond1 = StimulusOffSecond1 % 60
                if StimulusOffMinute1 >=60:
                    StimulusOffMinute1 = StimulusOffMinute1 % 60

                StimulusOnMinute2 = int(strftime("%M",)) + (stimulus_on2 / 60)
                StimulusOnSecond2 = int(strftime("%S",)) + (stimulus_on2 % 60)
                if StimulusOnSecond2 >=60:
                    StimulusOnMinute2 = StimulusOnMinute2 + 1
                    StimulusOnSecond2 = StimulusOnSecond2 % 60
                if StimulusOnMinute2 >=60:
                    StimulusOnMinute2 = StimulusOnMinute2 % 60
                StimulusOffMinute2 = int(strftime("%M",)) + (stimulus_off2 / 60)
                StimulusOffSecond2 = int(strftime("%S",)) + (stimulus_off2 % 60)
                if StimulusOffSecond2 >=60:
                    StimulusOffMinute2 = StimulusOffMinute2 +1
                    StimulusOffSecond2 = StimulusOffSecond2 % 60
                if StimulusOffMinute2 >=60:
                    StimulusOffMinute2 = StimulusOffMinute2 % 60
                starttime = datetime.datetime.now()
                frametime = starttime
            setupbox = 1
            if RGB == 1:
                #Below is for RGB
                Tank1= cv2.VideoWriter(tt81 + '\\Tank1.avi',fourcc, 20.0, (boxes[1][0]-boxes[0][0],boxes[1][1]-boxes[0][1]))
                Tank2= cv2.VideoWriter(tt82 + '\\Tank2.avi',fourcc, 16.0, (boxes[3][0]-boxes[2][0],boxes[3][1]-boxes[2][1]))
                Tank3= cv2.VideoWriter(tt83 + '\\Tank3.avi',fourcc, 16.0, (boxes[5][0]-boxes[4][0],boxes[5][1]-boxes[4][1]))
                Tank4= cv2.VideoWriter(tt84 + '\\Tank4.avi',fourcc, 16.0, (boxes[7][0]-boxes[6][0],boxes[7][1]-boxes[6][1]))
            else:
                #Below is for B/W
                Tank1= cv2.VideoWriter(tt81 + '\\Tank1.avi',fourcc, lastknownframerate, (boxes[1][0]-boxes[0][0],boxes[1][1]-boxes[0][1]),0)
                Tank2= cv2.VideoWriter(tt82 + '\\Tank2.avi',fourcc, lastknownframerate, (boxes[3][0]-boxes[2][0],boxes[3][1]-boxes[2][1]),0)
                Tank1_Raw= cv2.VideoWriter(tt83 + '\\Tank1_Raw.avi',fourcc, lastknownframerate, (boxes[1][0]-boxes[0][0]+50,boxes[1][1]-boxes[0][1]+70),0)
                Tank2_Raw= cv2.VideoWriter(tt84 + '\\Tank2_Raw.avi',fourcc, lastknownframerate, (boxes[3][0]-boxes[2][0]+50,boxes[3][1]-boxes[2][1]+70),0)

    if mode == 1:
        CurrentTime = strftime("%H",) + strftime("%M",) + strftime("%S",)
        if ValveFlag == 0:
            if int(strftime("%M",)) == int(StimulusOnMinute1):
                if int(strftime("%S",)) == int(StimulusOnSecond1):
                    print "LED ON : First Valve open"
                    ser.write('F')
                    ValveFlag = 1
        elif ValveFlag == 1:
            if int(strftime("%M",)) == int(StimulusOffMinute1):
                if int(strftime("%S",)) == int(StimulusOffSecond1):
                    print "LED OFF : First Valve closed"
                    ser.write('L')
                    ValveFlag = 2
        elif ValveFlag == 2:
            if int(strftime("%M",)) == int(StimulusOnMinute2):
                if int(strftime("%S",)) == int(StimulusOnSecond2):
                    print "LED ON : Second Valve open"
                    ser.write('S')
                    ValveFlag = 3
        elif ValveFlag == 3:
            if int(strftime("%M",)) == int(StimulusOffMinute2):
                if int(strftime("%S",)) == int(StimulusOffSecond2):
                    print "LED OFF : Second Valve closed"
                    ser.write('L')
                    ValveFlag = 4

        tt = time.time()
        tt1 = str(tt).find('.',1,len(str(tt)))
        tt2 = str(tt)[tt1-6:tt1]
        tt5 = str(tt)[tt1+1:len(str(tt))]
        tt4 = tt2 + tt5.ljust(2,'0')

        cv2.rectangle(gray2,(boxes[0][0],boxes[0][1]),(boxes[1][0],boxes[1][1]), (0,0,255),1)
        cv2.rectangle(gray2,(boxes[0][0]-50,boxes[0][1]-50),(boxes[1][0],boxes[1][1]+20), (0,255,0),1)
        cv2.rectangle(gray2,(boxes[2][0],boxes[2][1]),(boxes[3][0],boxes[3][1]), (0,0,255),1)
        cv2.rectangle(gray2,(boxes[2][0],boxes[2][1]-50),(boxes[3][0]+50,boxes[3][1]+20), (0,255,0),1)

        # Below is for Tank 1 tracking =======================================================
        img1 = cv2.subtract(gray2[boxes[0][1]:boxes[1][1],boxes[0][0]:boxes[1][0]],gray5[boxes[0][1]:boxes[1][1],boxes[0][0]:boxes[1][0]]) # global view
        img1 = cv2.cvtColor(img1,cv2.COLOR_BGR2GRAY)
        img1a = cv2.GaussianBlur(img1,(25,25),0)
        img1a = cv2.dilate(img1a,kernel,iterations=1)
        ret, img1b = cv2.threshold(img1a,65,255,0)  # Change the threshold
        edged = cv2.Canny(img1b,30,200)
        (cnts, _) = cv2.findContours(edged.copy(), cv2.RETR_TREE, cv2.CHAIN_APPROX_NONE)
        cnts = sorted(cnts, key=cv2.contourArea, reverse = True)[:10]
        if len(cnts) == 0:
            x = 0
            y = 0
            w = 0
            h = 0
            data1.append([0, 0, 0, 0, 0, 0, ValveFlag])
        else:
            [x, y, w, h] = cv2.boundingRect(cnts[0])
            data1.append([x+w/2+boxes[0][0], y+h/2+boxes[0][1], x, y, w, h, ValveFlag])
        cv2.rectangle(gray2,(x-2+boxes[0][0],y-2++boxes[0][1]),(x+w+2+boxes[0][0],y+h+2++boxes[0][1]),(0,255,0),2)
        cv2.circle(gray2, (x+w/2+boxes[0][0], y+h/2+boxes[0][1]), 10, (0,0,255),-1)

        # Below is for Tank 2 tracking =======================================================
        img2 = cv2.subtract(gray2[boxes[2][1]:boxes[3][1],boxes[2][0]:boxes[3][0]],gray5[boxes[2][1]:boxes[3][1],boxes[2][0]:boxes[3][0]]) # global view
        img2 = cv2.cvtColor(img2,cv2.COLOR_BGR2GRAY)
        img2a = cv2.GaussianBlur(img2,(25,25),0)
        img2a = cv2.dilate(img2a,kernel,iterations=1)
        ret, img2b = cv2.threshold(img2a,30,255,0)
        edged = cv2.Canny(img2b,30,200)
        (cnts, _) = cv2.findContours(edged.copy(), cv2.RETR_TREE, cv2.CHAIN_APPROX_NONE)
        cnts = sorted(cnts, key=cv2.contourArea, reverse = True)[:10]
        if len(cnts) == 0:
            x = 0
            y = 0
            w = 0
            h = 0
            data2.append([0, 0, 0, 0, 0, 0, ValveFlag])
        else:
            [x, y, w, h] = cv2.boundingRect(cnts[0])
            data2.append([x+w/2+boxes[2][0], y+h/2+boxes[2][1], x, y, w, h, ValveFlag])
        cv2.rectangle(gray2,(x-2+boxes[2][0],y-2++boxes[2][1]),(x+w+2+boxes[2][0],y+h+2++boxes[2][1]),(0,255,0),2)
        cv2.circle(gray2, (x+w/2+boxes[2][0], y+h/2+boxes[2][1]), 10, (0,0,255),-1)

        cv2.imshow('frame',gray2)
        if i%lastknownframerate*1 == 1:
            oldL1=0
            oldL2=0
        gray2 = cv2.cvtColor(gray2, cv2.COLOR_BGR2GRAY)
        gray6 = cv2.cvtColor(gray6, cv2.COLOR_BGR2GRAY)
        tt6[i]=int(tt4)

        tank1 = gray2[boxes[0][1]:boxes[1][1],boxes[0][0]:boxes[1][0]]
        tank1_raw = gray6[boxes[0][1]-50:boxes[1][1]+20,boxes[0][0]-50:boxes[1][0]]
        Tank1.write(tank1)
        Tank1_Raw.write(tank1_raw)


        data1Stamp.append([i-preset+1,tt6[i]])


        tank2 = gray2[boxes[2][1]:boxes[3][1],boxes[2][0]:boxes[3][0]]
        tank2_raw = gray6[boxes[2][1]-50:boxes[3][1]+20,boxes[2][0]:boxes[3][0]+50]
        Tank2.write(tank2)
        Tank2_Raw.write(tank2_raw)
        data2Stamp.append([i-preset+1,tt6[i]])

        frametime = datetime.datetime.now()
        elapsed_time = 0

        while (elapsed_time < (setframerate-0.032)):  # This is to control the frame rate 9.5 means 95 ms and this will give you about 10.45 fps(1000/95)
            tt = time.time()
            tt = time.time()
            tt1 = str(tt).find('.',1,len(str(tt)))
            tt2 = str(tt)[tt1-6:tt1]
            tt5 = str(tt)[tt1+1:len(str(tt))]
            tt4 = tt2 + tt5.ljust(2,'0')
            elapsed_time = (datetime.datetime.now() - frametime).total_seconds()
#            print (1/lastknownframerate)
    else:
        preset = preset + 1
        if preset < 3:
            cv2.imshow('frame',gray2)
        else:
            cv2.accumulateWeighted(gray2,gray4,0.1)
            gray5 = cv2.convertScaleAbs(gray4)
            cv2.imshow('frame',gray5)
    i=i+1
    if i%(lastknownframerate*60) == 0:
        print datetime.datetime.now() - starttime
    cv2.setMouseCallback('frame',set_up_ROIs)

    if cv2.waitKey(dummy) & 0xFF == ord('q'):
        save1=1
        break
     

tt7=float((int(tt6[total-1])-int(tt6[preset])))
tt7=total/(tt7/100)
print str(tt7) + ' frame per second'

# When everything done, release the capture
cap.release()

CurrentTime = strftime("%H",) + strftime("%M",) + strftime("%S",)
CurrentDate = strftime("%y",) + strftime("%m",) + strftime("%d",)


if save1==0:
    Tank1.release()
    Tank2.release()
    Tank1_Raw.release()
    Tank2_Raw.release()

    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Tank1")
    sheet2 = book.add_sheet("Tank2")

    rows = len(data1)
    sheet1.write(0, 0, 'Exp Time')
    sheet1.write(0, 1, time.ctime())
    sheet1.write(1, 0, 'Tank #')
    sheet1.write(1, 1, '1')
    sheet1.write(2, 0, 'Fish Group')
    sheet1.write(3, 0, 'Fish ID')
    sheet1.write(4, 0, 'Fish DOB')
    sheet1.write(5, 0, 'Tank-P1')
    sheet1.write(5, 2, boxes[0][0])
    sheet1.write(5, 3, boxes[0][1])
    sheet1.write(6, 0, 'Tank-P2')
    sheet1.write(6, 2, boxes[1][0])
    sheet1.write(6, 3, boxes[1][1])
    sheet1.write(8, 0, 'x-y coordinates')
    sheet1.write(9, 0, 'Frame #')
    sheet1.write(9, 1, 'Timestamp')
    sheet1.write(9, 2, 'x')
    sheet1.write(9, 3, 'y')

    sheet1.write(2, 4, 'TotalTime:' + str(total_experiment_time))
    sheet1.write(3, 4, 'StimulusOn:' + str(stimulus_on1))
    sheet1.write(4, 4, 'StimulusOff:' + str(stimulus_off1))
    sheet1.write(5, 4, 'FishName:' + FishName1)
    sheet1.write(6, 4, 'StimulusType:' + StimulusType1)
    sheet1.write(7, 4, 'FishType:' + FishType1)
    sheet1.write(8, 4, 'FishSex:' + FishSex1)
    sheet1.write(9, 4, 'SS_made_date1:' + SS_made_date1)
    sheet1.write(0, 4, 'CurrentDate:' + CurrentDate)
    sheet1.write(1, 4, 'CurrentTime:' + CurrentTime)
    sheet1.write(0, 6, 'FrameRate:' + str(tt7))
    sheet1.write(1, 6, 'FrameWidth:' + str(FrameWidth))
    sheet1.write(2, 6, 'FrameHeight:' + str(FrameHeight))
    sheet1.write(3, 6, 'TankWidth(mm):' + str(TankWidth))
    sheet1.write(4, 6, 'ROI failed:' + str(tt811a))
    sheet1.write(5, 6, 'Global failed:' + str(tt811b))
    sheet1.write(6, 6, 'Total zeros:' + str(tt811c))
    sheet1.write(7, 6, 'Total frames:' + str(rows))
    sheet1.write(8, 6, 'Search size:' + str(SearchSize))
    sheet1.write(9, 8, 'ValveFlag')

    sheet2.write(0, 0, 'Exp Time')
    sheet2.write(0, 1, time.ctime())
    sheet2.write(1, 0, 'Tank #')
    sheet2.write(1, 1, '2')
    sheet2.write(2, 0, 'Fish Group')
    sheet2.write(3, 0, 'Fish ID')
    sheet2.write(4, 0, 'Fish DOB')
    sheet2.write(5, 0, 'Tank-P1')
    sheet2.write(5, 2, boxes[2][0])
    sheet2.write(5, 3, boxes[2][1])
    sheet2.write(6, 0, 'Tank-P2')
    sheet2.write(6, 2, boxes[3][0])
    sheet2.write(6, 3, boxes[3][1])
    sheet2.write(8, 0, 'x-y coordinates')
    sheet2.write(9, 0, 'Frame #')
    sheet2.write(9, 1, 'Timestamp')
    sheet2.write(9, 2, 'x')
    sheet2.write(9, 3, 'y')
    sheet2.write(2, 4, 'TotalTime:' + str(total_experiment_time))
    sheet2.write(3, 4, 'StimulusOn:' + str(stimulus_on2))
    sheet2.write(4, 4, 'StimulusOff:' + str(stimulus_off2))
    sheet2.write(5, 4, 'FishName:' + FishName2)
    sheet2.write(6, 4, 'StimulusType:' + StimulusType2)
    sheet2.write(7, 4, 'FishType:' + FishType2)
    sheet2.write(8, 4, 'FishSex:' + FishSex2)
    sheet2.write(9, 4, 'SS_made_date1:' + SS_made_date2)
    sheet2.write(0, 4, 'CurrentDate:' + CurrentDate)
    sheet2.write(1, 4, 'CurrentTime:' + CurrentTime)
    sheet2.write(0, 6, 'FrameRate:' + str(tt7))
    sheet2.write(1, 6, 'FrameWidth:' + str(FrameWidth))
    sheet2.write(2, 6, 'FrameHeight:' + str(FrameHeight))
    sheet2.write(3, 6, 'TankWidth(mm):' + str(TankWidth))
    sheet2.write(4, 6, 'ROI failed:' + str(tt812a))
    sheet2.write(5, 6, 'Global failed:' + str(tt812b))
    sheet2.write(6, 6, 'Total zeros:' + str(tt812c))
    sheet2.write(7, 6, 'Total frames:' + str(rows))
    sheet2.write(8, 6, 'Search size:' + str(SearchSize))
    sheet2.write(9, 8, 'ValveFlag')

    for row in range(rows):
        for col in range(1):
            sheet1.write(row+10, col+0, data1Stamp[row][col])
            sheet1.write(row+10, col+1, str(data1Stamp[row][col+1]))
            sheet1.write(row+10, col+2, data1[row][col])
            sheet1.write(row+10, col+3, data1[row][col+1])
            sheet1.write(row+10, col+4, data1[row][col+2])
            sheet1.write(row+10, col+5, data1[row][col+3])
            sheet1.write(row+10, col+6, data1[row][col+4])
            sheet1.write(row+10, col+7, data1[row][col+5])
            sheet1.write(row+10, col+8, data1[row][col+6])

            sheet2.write(row+10, col+0, data2Stamp[row][col])
            sheet2.write(row+10, col+1, str(data2Stamp[row][col+1]))
            sheet2.write(row+10, col+2, data2[row][col])
            sheet2.write(row+10, col+3, data2[row][col+1])
            sheet2.write(row+10, col+4, data2[row][col+2])
            sheet2.write(row+10, col+5, data2[row][col+3])
            sheet2.write(row+10, col+6, data2[row][col+4])
            sheet2.write(row+10, col+7, data2[row][col+5])
            sheet2.write(row+10, col+8, data2[row][col+6])


    book.save("C:\\Users\\Administrator\\Desktop\\AlarmExcel\\AlarmTrackingData_" + GroupName + "_Raw.xls")
cv2.destroyAllWindows()
ser.close()
