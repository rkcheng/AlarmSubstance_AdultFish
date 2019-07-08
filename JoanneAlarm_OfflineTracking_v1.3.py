import sys
sys.path.append("C:\\opencv\\build\\python\\2.7")
import copy
import Tkinter, tkFileDialog, Tkconstants
from Tkinter import *
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.widgets import Button
from matplotlib.ticker import MultipleLocator, FormatStrFormatter
import cv2
import math
import time
from time import strftime
from time import time
import os
import seaborn as sns
import re
from numpy import array
import xlwt
import xlrd
import datetime
import csv
import msvcrt

distance_in_seconds =[]
speed_before=[]
speed_after=[]
single_speed=[]
ExperimentID=[]

def nothing(x):
    pass

def JumpFrame():
    print 'Number of Jumped Frame is: ' + str(tempD1)
    print 'Jump threshold is: ' + str(tempD2) + " pixels/frame"

    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Jump List")
    for i in range(tempD1):
        for j in range(1):
            sheet1.write(i+1, 0, tempD[i][0])
            sheet1.write(i+1, 1, tempD[i][1])
            sheet1.write(i+1, 2, tempD[i][2])
    sheet1.write(0, 0, 'Item#')
    sheet1.write(0, 1, 'Frame#')
    sheet1.write(0, 2, 'Distance (pixels)')
    sheet1.write(0, 4, 'The threshold is ')
    sheet1.write(0, 6, tempD2)
    sheet1.write(0, 7, ' pixels')
    if tempD1 <= 1:
        if WhichTank == 1:
            if ExcelStatus == 0:
                if int(OnOff) == 2:
                    book.save(path + ExperimentID + "Tank1_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame(s).xls")
                else:
                    book.save(path + ExperimentID + "Tank1_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame(s).xls")
            else:
                book.save(path + ExperimentID + "Tank1_Online_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame(s).xls")

        else:
            if ExcelStatus == 0:
                if int(OnOff) == 2:
                    book.save(path + ExperimentID + "Tank2_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame(s).xls")
                else:
                    book.save(path + ExperimentID + "Tank2_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame(s).xls")
            else:
                book.save(path + ExperimentID + "Tank2_Online_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame(s).xls")
    else:
        if WhichTank == 1:
            if ExcelStatus == 0:
                if int(OnOff) == 2:
                    book.save(path + ExperimentID + "Tank1_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame.xls")
                else:
                    book.save(path + ExperimentID + "Tank1_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame.xls")
            else:
                book.save(path + ExperimentID + "Tank1_Online_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame.xls")
        else:
            if ExcelStatus == 0:
                if int(OnOff) == 2:
                    book.save(path + ExperimentID + "Tank2_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame.xls")
                else:
                    book.save(path + ExperimentID + "Tank2_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame.xls")
            else:
                book.save(path + ExperimentID + "Tank2_Online_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame.xls")
    return;

def Behavior1():
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Zero List")
    sheet1.write(0, 0, '** Up-Down')
    sheet1.write(1, 0, tempUp)
    sheet1.write(1, 1, tempDown)
    sheet1.write(1, 2, "{:.2%}".format(tempUp/float(total-9)))
    sheet1.write(1, 3, "{:.2%}".format(tempDown/float(total-9)))
    sheet1.write(2, 0, '** Left-Center-Right')
    sheet1.write(3, 0, tempLeft)
    sheet1.write(3, 1, tempCenter)
    sheet1.write(3, 2, tempRight)
    sheet1.write(3, 3, "{:.2%}".format(tempLeft/float(total-9)))
    sheet1.write(3, 4, "{:.2%}".format(tempCenter/float(total-9)))
    sheet1.write(3, 5, "{:.2%}".format(tempRight/float(total-9)))
    if WhichTank == 1:
        if ExcelStatus == 0:
            if int(OnOff) == 2:
                book.save(path + ExperimentID + "Tank1_Offline_Behavior Data.xls")
            else:
                book.save(path + ExperimentID + "Tank1_Offline_Behavior Data.xls")
        else:
            book.save(path + ExperimentID + "Tank1_Online_Behavior Data.xls")
    else:
        if ExcelStatus ==0:
            if int(OnOff) == 2:
                book.save(path + ExperimentID + "Tank2_Offline_Behavior Data.xls")
            else:
                book.save(path + ExperimentID + "Tank2_Offline_Behavior Data.xls")
        else:
            book.save(path + ExperimentID + "Tank2_Online_Behavior Data.xls")
    # for Up-Down analysis ===================================
    print 'Percent of time in the top of the tank: ' + "{:.2%}".format(tempUp/(float(total-9)))
    print 'Percent of time in the bottom of the tank: ' + "{:.2%}".format(tempDown/(float(total-9)))
    print '========================================'
    # for Left-Center-Right analysis ============================
    print 'Percent of time in the left of the tank: ' + "{:.2%}".format(tempLeft/(float(total-9)))
    print 'Percent of time in the center of the tank: ' + "{:.2%}".format(tempCenter/(float(total-9)))
    print 'Percent of time in the right of the tank: ' + "{:.2%}".format(tempRight/(float(total-9)))
    return;

def Plotting1():
    label_in_seconds = np.divide(array1, 60.)
    fs = plt.figure(figsize = (18,5))
    gs = plt.GridSpec(2,3,height_ratios=[2,1])
    ax1 = fs.add_subplot(gs[0,0])
    z6 = z5[1:OnFrame]
    z6 = np.asarray(z6) * float(200) / float(int(boxes[1][0])-int(boxes[0][0]))
    z6mean = np.average(z6)
    ax1.set_title('Before Stimulus Speed = ' + "{:.3}".format(str(z6mean)) + " mm/sec")
    ax2 = fs.add_subplot(gs[1, 0:3])

    x6 = x5[1:OnFrame]
    y6 = y5[1:OnFrame]

    z6x = label_in_seconds[range(1, OnFrame)]

    colors = sns.color_palette('GnBu', (np.size(x6)/int(OriginalFrameRate))+12)
    count=10
    for ii in xrange(1,np.size(x6),int(OriginalFrameRate)):
        ax1.plot(x6[ii-1:ii+int(OriginalFrameRate)],y6[ii-1:ii+int(OriginalFrameRate)],linewidth=3, color=colors[count], alpha=0.5)
        ax2.plot(z6x[ii-1:ii+int(OriginalFrameRate)],z6[ii-1:ii+int(OriginalFrameRate)],linewidth=2, color=colors[count], alpha=0.5)
        count += 1

    ax1.set_xlim(0,1)
    ax1.set_ylim(0,1)
    ax = ax1.axis()
    ax1.axis((ax[0],ax[1],ax[3],ax[2]))


    ax1 = fs.add_subplot(gs[0,1])
    z6 = z5[OnFrame:OffFrame]
    z6 = np.asarray(z6) * float(200) / float(int(boxes[1][0])-int(boxes[0][0]))
    z6mean = np.average(z6)
    ax1.set_title('During Stimulus Speed = ' + "{:.3}".format(str(z6mean)) + " mm/sec")
    x6 = x5[OnFrame:OffFrame]
    y6 = y5[OnFrame:OffFrame]
    z6x = label_in_seconds[range(OnFrame,OffFrame)]

    colors = sns.color_palette('YlGn', (np.size(x6)/int(OriginalFrameRate))+12)
    count=10
    for ii in xrange(1,np.size(x6),int(OriginalFrameRate)):
        ax1.plot(x6[ii-1:ii+int(OriginalFrameRate)],y6[ii-1:ii+int(OriginalFrameRate)],linewidth=3, color=colors[count], alpha=0.5)
        ax2.plot(z6x[ii-1:ii+int(OriginalFrameRate)],z6[ii-1:ii+int(OriginalFrameRate)],linewidth=2, color=colors[count], alpha=0.5)
        count += 1
    ax1.set_xlim(0,1)
    ax1.set_ylim(0,1)
    ax = ax1.axis()
    ax1.axis((ax[0],ax[1],ax[3],ax[2]))


    ax1 = fs.add_subplot(gs[0,2])
    z6 = z5[OffFrame:]
    z6 = np.asarray(z6) * float(200) / float(int(boxes[1][0])-int(boxes[0][0]))
    z6mean = np.average(z6)
    ax1.set_title('After Stimulus Speed = ' + "{:.3}".format(str(z6mean)) + " mm/sec")
    x6 = x5[OffFrame:]
    y6 = y5[OffFrame:]
    if int(OnOff) == 2:
        z6x = label_in_seconds[range(OffFrame,np.size(array1)-1)]
    else:
        z6x = label_in_seconds[range(OffFrame,np.size(array1))]
    colors = sns.color_palette('YlOrRd', (np.size(x6)/int(OriginalFrameRate))+12)
    count=10
    for ii in xrange(1,np.size(x6),int(OriginalFrameRate)):
        ax1.plot(x6[ii-1:ii+int(OriginalFrameRate)],y6[ii-1:ii+int(OriginalFrameRate)],linewidth=3, color=colors[count], alpha=0.5)
        ax2.plot(z6x[ii-1:ii+int(OriginalFrameRate)],z6[ii-1:ii+int(OriginalFrameRate)],linewidth=2, color=colors[count], alpha=0.5)
        count += 1
    ax1.set_xlim(0,1)
    ax1.set_ylim(0,1)
    ax = ax1.axis()
    ax1.axis((ax[0],ax[1],ax[3],ax[2]))

    ax2.set_xlabel('Time (minutes)')
    ax2.set_ylabel('Distance (mm)')

    ax2.set_ylim((0, 40))
    major_ticks = np.arange(0, 41, 8)
    ax2.set_yticks(major_ticks)
    ax2.set_xlim((0, 15))
    major_ticks = np.arange(0, 16, 5)
    ax2.set_xticks(major_ticks)
    ax2.minorticks_on()
    ax2.tick_params(axis='x', which='minor', direction='out',length=8, width = 1)
    ax2.tick_params(axis='x', which='major', direction='out',length=12, width = 1)
    class Index1(object):
        def next1(self, event):
            plt.savefig(path + ExperimentID[:-7] + '_Histogram.pdf')
            plt.close('all')
            sys.exit()
    class Index(object):
        def next(self, event):
            if ExcelStatus == 0:
                if WhichTank == 1:
                    if int(OnOff) == 2:
                        plt.savefig(path + 'Tank1_Offline.pdf')
                    else:
                        plt.savefig(path + ExperimentID + 'Tank1_Offline.pdf')
                else:
                    if int(OnOff) == 2:
                        plt.savefig(path + 'Tank2_Offline.pdf')
                    else:
                        plt.savefig(path +ExperimentID + 'Tank2_Offline.pdf')
            else:
                if WhichTank == 1:
                    plt.savefig(path +ExperimentID + 'Tank1_Online.pdf')
                else:
                    plt.savefig(path +ExperimentID + 'Tank2_Online.pdf')
            if WhichTank2 == 3:
                if iii < iiii:
                    plt.close('all')
                    if iii + 1 == iiii:
                        if int(FirstOrBoth) == 1:
                            if int(AllOrNone) == 2:
                                plt.close('all')
                                sys.exit()
                            else:
                                fs = plt.figure(figsize = (10,6))
                                gs = plt.GridSpec(2,2)
                                ax1 = fs.add_subplot(gs[0,:])
                                ax2 = ax1.twinx()
                                ax3 = fs.add_subplot(gs[1,0])
                                color = [sns.color_palette()[0], sns.color_palette()[2]]
                                label = ['Before', 'After']
                                ax1.hist(speed_before, np.size(speed_before), density=2, color = 'b', facecolor=color[0], histtype = 'bar', alpha=0.5, label=['Before'])
                                sns.distplot(speed_before, np.size(speed_before), hist=False, kde=True,ax=ax2, color = 'b')
                                ax3.hist(speed_before, np.size(speed_before), density=1, color = 'b', facecolor=color[0], histtype = 'step', cumulative = True, lw=3, label=['Before'])
                                ax1.hist(speed_after, np.size(speed_after), density=2, color = 'g', facecolor=color[1], histtype = 'bar', alpha=0.5, label=['After'])
                                sns.distplot(speed_after, np.size(speed_after), hist=False, kde=True,ax=ax2, color = 'g')
                                ax3.hist(speed_after, np.size(speed_after), density=1, color = 'g', facecolor=color[1], histtype = 'step', cumulative = True, lw=3, label=['After'])
                                legend = ax1.legend(loc='upper right', shadow=True)
                                legend = ax3.legend(loc='upper right', shadow=True)
                                ax1.set_ylabel('Number of occurrences')
                                ax1.set_xlabel('Speed (cm/s)')
                                ax1.set_xlim(-5,20)
                                ax1.set_ylim(0,1)
                                ax2.set_ylabel('Probability')
                                ax3.set_ylabel('Cumulative Probability')
                                ax3.set_xlabel('Speed(cm/s)')
                                ax3.set_xlim(0.3,30)
                                plt.tight_layout()
                                callback = Index1()
                                axnext = plt.axes([0.8, 0.01, 0.15, 0.05])
                                bnext = Button(axnext, 'Save and Close')
                                bnext.on_clicked(callback.next1)
                                plt.show()
                else:
                    if int(AllOrNone) == 2:
                        plt.close('all')
                        sys.exit()
                    else:
                        plt.close('all')
                        fs = plt.figure(figsize = (10,6))
                        gs = plt.GridSpec(2,2)
                        ax1 = fs.add_subplot(gs[0,:])
                        ax2 = ax1.twinx()
                        ax3 = fs.add_subplot(gs[1,0])
                        color = [sns.color_palette()[0], sns.color_palette()[2]]
                        label = ['Before', 'After']
                        ax1.hist(speed_before, np.size(speed_before), density=2, color = 'b', facecolor=color[0], histtype = 'bar', alpha=0.5, label=['Before'])
                        sns.distplot(speed_before, np.size(speed_before), hist=False, kde=True,ax=ax2, color = 'b')
                        ax3.hist(speed_before, np.size(speed_before), density=1, color = 'b', facecolor=color[0], histtype = 'step', cumulative = True, lw=3, label=['Before'])
                        ax1.hist(speed_after, np.size(speed_after), density=2, color = 'g', facecolor=color[1], histtype = 'bar', alpha=0.5, label=['After'])
                        sns.distplot(speed_after, np.size(speed_after), hist=False, kde=True,ax=ax2, color = 'g')
                        ax3.hist(speed_after, np.size(speed_after), density=1, color = 'g', facecolor=color[1], histtype = 'step', cumulative = True, lw=3, label=['After'])
                        legend = ax1.legend(loc='upper right', shadow=True)
                        legend = ax3.legend(loc='upper right', shadow=True)
                        ax1.set_ylabel('Number of occurrences')
                        ax1.set_xlabel('Speed (cm/s)')
                        ax1.set_xlim(-5,20)
                        ax1.set_ylim(0,1)
                        ax2.set_ylabel('Probability')
                        ax3.set_ylabel('Cumulative Probability')
                        ax3.set_xlabel('Speed(cm/s)')
                        ax3.set_xlim(0.3,30)
                        plt.tight_layout()
                        callback = Index1()
                        axnext = plt.axes([0.8, 0.01, 0.15, 0.05])
                        bnext = Button(axnext, 'Save and Close')
                        bnext.on_clicked(callback.next1)
                        plt.show()
            else:
                plt.close('all')
                sys.exit()
    callback = Index()
    axnext = plt.axes([0.915, 0.01, 0.085, 0.05])
    bnext = Button(axnext, 'Save and Close')
    bnext.on_clicked(callback.next)
    cv2.destroyAllWindows()
    plt.show()

# User input below
# =======================================
tempD2 =80  # Pixel threshold of jump (jump may indicate tracking errors)
# =======================================
RGB = 0 # 1 stands for color video and 0 stands for B/W video

x0 = 0
y0 = 0
distance1 = 0
tempD1 = 0
tempD = []
tempZero =[]
tempZero1 = 0
tempUp = 0
tempDown = 0
tempLeft = 0
tempCenter = 0
tempRight = 0
##Experiment parameters
setframerate = 80
setframerate = 1/float(setframerate)
stamp = np.zeros((100,200,3))
root = Tk()
OriginalFrameRate = 0
OriginalFrameRate = 0
Downsize = 0
OnFrame = 0
OffFrame = 0
x5=[];
y5=[];
z5=[];
ExcelStatus=0
TimeStamp = []
array1=[];
boxes=[]
xlist=[]
WhichTank=[]
ValveFlag = 0
iii=0
FileName=[]
speedtime = []
count1=0
root.filename2 = tkFileDialog.askopenfilename(initialdir = "/",title = "Select the raw Excel file",filetypes = (("videos files","*.xls"),("all files","*.*")))
OnOff = raw_input("Analyze Online data (1) or commence Offline tracking (2) -----> Please enter here: ")
full_path = os.path.realpath(root.filename2)
path, filename = os.path.split(full_path)
path = path + '\\'
FileName = [ii for ii in os.listdir(path) if ii.endswith('xls')]
if int(OnOff) == 2 :  #  For offline tracking
    root.filename1 = tkFileDialog.askopenfilename(initialdir = "/",title = "Select a video file",filetypes = (("videos files","*.avi"),("all files","*.*")))
    if not root.filename1:
        sys.exit()
    root.destroy()
    print "The opened video file is ..."
    print (root.filename1)
    WhichDataSet = raw_input("Which data set does this video belong to? (1) Seetha's video (2) Ruey's old (3) Ruey's new -----> Please enter here: ")
    if int(WhichDataSet) > 1:
        WhichTank = root.filename1[len(root.filename1)-9:len(root.filename1)-8]
        WhichTank=int(WhichTank)
        WhichTank2 = 4
        if int(WhichDataSet) == 2:
            full_path = os.path.realpath(root.filename1)
            path, filename = os.path.split(full_path)
            path = path + '\\'
            ExperimentID = filename
            ExperimentID, temp1 = ExperimentID.split(".avi")

            print path
            print ExperimentID

            if WhichTank == 1:
                df = pd.read_excel(root.filename2, sheet_name='Tank1')
                column1 = df.as_matrix(columns=None)
                for i in range(17000,19000,1):
                    if column1[i][5] + column1[i+1][5] == 1:
                        OnFrame = i-7
                        count = 1
                        print OnFrame
                        while (column1[i+1][5] + column1[i+count+1][5]) != 3:
                            count += 1
                        OffFrame = i + count-7
                        i = 19000
                print OnFrame
                print OffFrame
            else:
                df = pd.read_excel(root.filename2, sheet_name='Tank2')
                column1 = df.as_matrix(columns=None)
                for i in range(17000,19000,1):
                    if column1[i][5] + column1[i+1][5] == 1:
                        OnFrame = i-7
                        count = 1
                        while (column1[i+1][5] + column1[i+count+1][5]) != 3:
                            count += 1
                        OffFrame = i + count-7
                        i = 19000
                print OnFrame
                print OffFrame

        if int(WhichDataSet) == 3:
            full_path = os.path.realpath(root.filename1)
            path, filename = os.path.split(full_path)
            path = path + '\\'
            ExperimentID = filename
            ExperimentID, temp1 = ExperimentID.split(".avi")
            print ExperimentID
            if WhichTank == 1:
                print root.filename2
                df = pd.read_excel(root.filename2, sheet_name='Tank1')
                column1 = df.as_matrix(columns=None)
                for i in range(3000,10000,1):
                    if column1[i][8] + column1[i+1][8] == 1:
                        OnFrame = i-7
                        count = 1
                        while ((int(column1[i+count][1])-int(column1[i][1]))<=3000):
                            count += 1
                        OffFrame = i + 25-7
                        i = 10000
            else:
                df = pd.read_excel(root.filename2, sheet_name='Tank2')
                column1 = df.as_matrix(columns=None)
                for i in range(3000,10000,1):
                    if column1[i][8] + column1[i+1][8] == 5:
                        OnFrame = i-7
                        count = 1
                        while ((int(column1[i+count][1])-int(column1[i][1]))<=3000):
                            count += 1
                        OffFrame = i + 25-7
                        i = 10000
    else:
        OnFrame = raw_input("Enter OnFrame -----> Please enter here: ")
        OffFrame = raw_input("Enter OffFrame -----> Please enter here: ")
        OnFrame = int(OnFrame)
        OffFrame = int(OffFrame)
        WhichTank2 = 4
        WhichTank= 1
        full_path = os.path.realpath(root.filename1)
        path, filename = os.path.split(full_path)
        path = path + '\\'
        ExperimentID = filename
        ExperimentID, temp1 = ExperimentID.split(".avi")
    cap = cv2.VideoCapture(root.filename1)
    total= cap.get(7)

    print "Video width is = " + str(cap.get(3)) # change Width
    print "Video height is = " + str(cap.get(4)) # change Height
    if cap.get(3) == 1280:
        if cap.get(4) == 960:
            Downsize = 1
            tempD2 = tempD2 / 2
            print "Need to downsize the image ========================="

    print "Video total frame number is = " + str(cap.get(7)) # change Height
    ExcelStatus = 0
    #OriginalFrameRate = raw_input("Enter original frame rate (fps) -----> Please enter here: ")
    if int(WhichDataSet) > 1:
        temp1, OriginalFrameRate1 = df.columns[6].split(":")
        OriginalFrameRate1 = OriginalFrameRate1[:2]
        OriginalFrameRate = float(OriginalFrameRate1)
        column1 = df.as_matrix(columns=None)
        for i in range(9,int(total+9),1):
            array1.append(float((int(column1[i][1])-int(column1[9][1]))/100))
    else:
        OriginalFrameRate = 30
        temp31 = float(1) /int(OriginalFrameRate)
        for i in range(1,int(total+1),1):
            array1.append(float(i*temp31))
else:
    ExcelStatus = 1  # Use Online tracking data
    root.destroy()
    AllOrNone = raw_input("Do you want to analyze all Excel files (1) or just a single Excel file (2) -----> Please enter here: ")
    if int(AllOrNone) == 2:
        WhichTank2 = raw_input("Enter which tank to analyze (1 or 2 or 3 for both tanks) -----> Please enter here: ")
        WhichTank2 = int(WhichTank2)
        FirstOrBoth = '2'
    else:
        WhichTank2 = 3
        FirstOrBoth = raw_input("Do you want to analyze the first tank of each file (1) or both tanks (2) -----> Please enter here: ")
    if WhichTank2 == 1:
        df = pd.read_excel(root.filename2, sheet_name='Tank1')
        WhichTank = 1
        ValveFlag = 1
        iiii=1
        iii=1
    elif WhichTank2 == 2:
        WhichTank = 2
        df = pd.read_excel(root.filename2, sheet_name='Tank2')
        ValveFlag = 5
        iiii=2
        iii=2
    elif WhichTank2 == 3:
        if int(AllOrNone) == 1:
            iii = 1
            iiii = 2*len(FileName)
            iiiii = 1
        else:
            iii = 1
            iiii = 2
    while (iii<= iiii):
        if iii % 2 == 1:
            if int(AllOrNone) == 1:
                df = pd.read_excel(path + FileName[iiiii-1], sheet_name='Tank1')
                ExperimentID = FileName[iiiii-1]
            else:
                df = pd.read_excel(root.filename2, sheet_name='Tank1')
                ExperimentID = root.filename2
            temp1, ExperimentID = ExperimentID.split("AlarmTrackingData_")
            ExperimentID = ExperimentID[:-7]
            print 'Now analyzing ... ' + ExperimentID + ' Tank1 ========================================'
            WhichTank = 1
            ValveFlag = 1
            temp1, OriginalFrameRate1 = df.columns[6].split(":")
            OriginalFrameRate1 = OriginalFrameRate1[:2]
            column1 = df.as_matrix(columns=None)
        elif iii % 2 == 0:
            if int(AllOrNone) == 1:
                df = pd.read_excel(path + FileName[iiiii-1], sheet_name='Tank2')
                ExperimentID = FileName[iiiii-1]
                iiiii += 1
                temp1, ExperimentID = ExperimentID.split("AlarmTrackingData_")
                ExperimentID = ExperimentID[:-7]
                temp1, OriginalFrameRate1 = df.columns[6].split(":")
                OriginalFrameRate1 = OriginalFrameRate1[:2]
                column1 = df.as_matrix(columns=None)
            else:
                df = pd.read_excel(root.filename2, sheet_name='Tank2')
                ExperimentID = root.filename2
                temp1, ExperimentID = ExperimentID.split("AlarmTrackingData_")
                ExperimentID = ExperimentID[:-7]
                temp1, OriginalFrameRate1 = df.columns[6].split(":")
                OriginalFrameRate1 = OriginalFrameRate1[:2]
                column1 = df.as_matrix(columns=None)
            if int(FirstOrBoth) == 1:
                print 'Now skipping ... ' + ExperimentID + ' Tank2 ========================================'
            else:
                print 'Now analyzing ... ' + ExperimentID + ' Tank2 ========================================'
            temp1, OriginalFrameRate1 = df.columns[6].split(":")
            OriginalFrameRate1 = OriginalFrameRate1[:2]
            column1 = df.as_matrix(columns=None)
            WhichTank = 2
            ValveFlag = 5

        for i in range(3000,4000,1):
            if column1[i][8] + column1[i+1][8] == ValveFlag:
                OnFrame = i-7
                OnTime = (int(column1[i][1]) - int(column1[9][1]))/100
                count = 1
                while ((int(column1[i+count][1])-int(column1[i][1]))<=3000):
                    count += 1
                OffFrame = i + count-7
                OffTime = (int(column1[i+count][1]) - int(column1[9][1]))/100
                i = 4000
        total = len(column1)
        temp1, FishID = column1[4][4].split(":")
        OriginalFrameRate = float(OriginalFrameRate1)
        sbox = (df['Unnamed: 2'][4], df['Unnamed: 3'][4])
        boxes.append(sbox)
        sbox = (df['Unnamed: 2'][5], df['Unnamed: 3'][5])
        boxes.append(sbox)
        for i in range(9,int(total),1):
            array1.append(float((int(column1[i][1])-int(column1[9][1]))/100))
            x = df['Unnamed: 2'][i]
            y = df['Unnamed: 3'][i]
            if (x) > (2 * (int((boxes[1][0] - boxes[0][0])/3)) + boxes[0][0]):
                tempRight = tempRight + 1
            else:
                if (x) <= (int((boxes[1][0] - boxes[0][0])/3) + boxes[0][0]):
                    tempLeft = tempLeft + 1
                else:
                    tempCenter = tempCenter + 1
            if (y) > (int((boxes[1][1] - boxes[0][1])/2) + boxes[0][1]):
                tempDown = tempDown + 1
            else:
                tempUp = tempUp + 1
            x5.append(float(x-int(boxes[0][0]))/float((int(boxes[1][0])-int(boxes[0][0]))))
            y5.append(float(y-int(boxes[0][1]))/float((int(boxes[1][1])-int(boxes[0][1]))))
            if i==9:
                distance1 = 0
                z5.append(distance1)
                x0 = x
                y0 = y
            else:
                distance1 = math.sqrt((x-x0)**2 + ((y-y0)**2))
                if round(distance1) > tempD2:
                    tempD1 = tempD1 + 1
                    tempD.append([tempD1, i, int(distance1)])
                z5.append(distance1)
                x0 = x
                y0 = y
        count1 = 1
        for dd in xrange(0,np.size(z5), int(OriginalFrameRate)):
            if WhichTank == 2:
                if int(FirstOrBoth) == 1:
                    pass
                else:
                    distance_in_seconds.append (np.sum(z5[dd:dd+int(OriginalFrameRate)]))
                    speedtime.append(count1)
                    single_speed.append (np.sum(z5[dd:dd+int(OriginalFrameRate)]))
                    count +=1
            else:
                if int(FirstOrBoth) == 1:
                    if iii % 2 == 1:
                        distance_in_seconds.append (np.sum(z5[dd:dd+int(OriginalFrameRate)]))
                        speedtime.append(count1)
                        single_speed.append (np.sum(z5[dd:dd+int(OriginalFrameRate)]))
                        count +=1
                    else:
                        pass
                else:
                    distance_in_seconds.append (np.sum(z5[dd:dd+int(OriginalFrameRate)]))
                    speedtime.append(count1)
                    single_speed.append (np.sum(z5[dd:dd+int(OriginalFrameRate)]))
                    count +=1

        for dd in xrange(0,np.size(single_speed),1):
            if WhichTank == 2:
                if int(FirstOrBoth) == 1:
                    pass
                else:
                    if dd <= OnTime:
                        speed_before.append(single_speed[dd])
                    else:
                        speed_after.append(single_speed[dd])
            else:
                if int(FirstOrBoth) == 1:
                    if iii % 2 == 1:
                        if dd <= OnTime:
                            speed_before.append(single_speed[dd])
                        else:
                            speed_after.append(single_speed[dd])
                else:
                    if dd <= OnTime:
                        speed_before.append(single_speed[dd])
                    else:
                        speed_after.append(single_speed[dd])
        single_speed = []
        if int(FirstOrBoth) == 1:
            if iii == iiii:
                pass
                #distance_in_seconds = (np.asarray(distance_in_seconds) * float(200) / float(int(boxes[1][0])-int(boxes[0][0])))/10
                #speed_before = (np.asarray(speed_before) * float(200) / float(int(boxes[1][0])-int(boxes[0][0])))/10
                #speed_after = (np.asarray(speed_after) * float(200) / float(int(boxes[1][0])-int(boxes[0][0])))/10
            else:
                if iii +1 == iiii:
                    distance_in_seconds = (np.asarray(distance_in_seconds) * float(200) / float(int(boxes[1][0])-int(boxes[0][0])))/10
                    speed_before = (np.asarray(speed_before) * float(200) / float(int(boxes[1][0])-int(boxes[0][0])))/10
                    speed_after = (np.asarray(speed_after) * float(200) / float(int(boxes[1][0])-int(boxes[0][0])))/10
            if WhichTank == 2:
                iii +=1
                pass
            else:
                JumpFrame()
                Behavior1()
                Plotting1()
                iii +=1
        else:
            JumpFrame()
            Behavior1()
            if iii == iiii:
                distance_in_seconds = (np.asarray(distance_in_seconds) * float(200) / float(int(boxes[1][0])-int(boxes[0][0])))/10
                speed_before = (np.asarray(speed_before) * float(200) / float(int(boxes[1][0])-int(boxes[0][0])))/10
                speed_after = (np.asarray(speed_after) * float(200) / float(int(boxes[1][0])-int(boxes[0][0])))/10
            Plotting1()
            iii +=1
        df=[];
        boxes=[]
        sbox=[]
        x5=[]
        y5=[]
        z5=[]
        array1=[]
        tempRight=0
        tempCenter=0
        tempLeft=0
        tempUp=0
        tempDown=0
        x0 = 0
        y0 = 0
        distance1 = 0
        tempD1 = 0
        tempD = []
        tempZero =[]
        tempZero1 = 0

if ExcelStatus == 0:  # For Offline tracking
    drawing = False
    mode = 0
    dummy= 0
    data1=[];
    data1Stamp=[];

    ret, gray4 = cap.read()
    gray4 = np.float32(gray4)
    if Downsize == 1:
        gray4 = cv2.resize(gray4, None, fx=0.4, fy=0.4, interpolation = cv2.INTER_CUBIC)

    print "Calculating background image"
    for i in range(0,int(total),1):
        ret, frame = cap.read()
        if Downsize == 1:
            gray2 = frame [:]
            gray2 = cv2.resize(gray2, None, fx=0.4, fy=0.4, interpolation = cv2.INTER_CUBIC)
        else:
            gray2 = frame [:]
        if i%int(total/100) ==0:
            cv2.accumulateWeighted(gray2,gray4,0.1)
            gray5 = cv2.convertScaleAbs(gray4)
            if i > int(total/4):
                cv2.imshow('Background Image',gray5)
                break

    cap.release()
    cap = cv2.VideoCapture(root.filename1)
    cv2.namedWindow('Tracking ...')
    ret, frame1 = cap.read()
    if Downsize == 1:
        frame1 = cv2.resize(frame1, None, fx=0.4, fy=0.4, interpolation = cv2.INTER_CUBIC)

    cap.set(0, 0) # set starting frame
    startingframe = cap.get(0)

    def ClickFish(event, x, y,flags,param):
        print 'here is', i
        if event == cv2.EVENT_LBUTTONDOWN:
            print "New Fish Position", x, y

    def set_up_ROIs(event,x,y,flags,param):
        if len(boxes) <= 4:
            global boxnumber, drawing, img, gray2, arr, arr1, arr2, arr3, dummy, mode, tank1a, tank2a, tank3a, tank4a, StimulusOnMinute, StimulusOffMinute, StimulusOnSecond, StimulusOffSecond, OldY1, OldX1
            boxnumner = 0
            if event == cv2.EVENT_LBUTTONDOWN:
                drawing = True
                sbox = [x, y]
                boxes.append(sbox)
            elif event == cv2.EVENT_MOUSEMOVE:
                if drawing == True:
                    if len(boxes) == 1:
                        arr = array(gray2)
                        cv2.rectangle(arr,(boxes[-1][0],boxes[-1][-1]),(x, y),(0,0,255),1)
                        cv2.imshow('Tracking ...',arr)
                        stamp = np.zeros((100,200,3))
                    elif len(boxes) == 3:
                        arr1 = array(gray2)
                        cv2.rectangle(arr1,(boxes[-1][0],boxes[-1][-1]),(x, y),(0,0,255),1)
                        cv2.imshow('Tracking ...',arr1)
            elif event == cv2.EVENT_LBUTTONUP:
                drawing = not drawing
                ebox = [x, y]
                boxes.append(ebox)
                cv2.rectangle(gray2,(boxes[-2][-2],boxes[-2][-1]),(boxes[-1][-2],boxes[-1][-1]), (0,0,255),1)
                cv2.imshow('Tracking ...',gray2)
                if len(boxes) == 4:
                    mode = 1
                    dummy = 1
                    i=0
    i = int(total)
    starttime = time()
    if ExcelStatus == 0:
        i = 0
    kernel = np.ones((5, 5), np.uint8)
    while(i<(total)):
        if i == 1:
            starttime = time()
        ret, frame = cap.read()
        gray2 = frame [:]
        if Downsize == 1:
            gray2 = cv2.resize(gray2, None, fx=0.4, fy=0.4, interpolation = cv2.INTER_CUBIC)
        if mode == 1:
            img1 = cv2.subtract(gray2[boxes[0][1]:boxes[1][1],boxes[0][0]:boxes[1][0]],gray5[boxes[0][1]:boxes[1][1],boxes[0][0]:boxes[1][0]]) # global view
            img1 = cv2.cvtColor(img1,cv2.COLOR_BGR2GRAY)
            img1a = cv2.GaussianBlur(img1,(25,25),0)
            img1a = cv2.dilate(img1a,kernel,iterations=1)
            ret, img1b = cv2.threshold(img1a,30,255,0)
            edged = cv2.Canny(img1b,30,200)
            cv2.imshow('Skeleton View',edged)
            (cnts, _) = cv2.findContours(edged.copy(), cv2.RETR_TREE, cv2.CHAIN_APPROX_NONE)
            cnts = sorted(cnts, key=cv2.contourArea, reverse = True)[:1]
            if len(cnts) == 0:
                x = 0
                y = 0
                w = 0
                h = 0
                data1.append([0, 0, 0, 0, 0, 0, 0, 0, 0])
                x5.append(0)
                y5.append(0)
                #z5.append("{:.2}".format(float(0)))
                z5.append(0)
                tempZero1 = tempZero1 + 1
                tempZero.append([tempZero1, i])
            else:
                [x, y, w, h] = cv2.boundingRect(cnts[0])
                rect1 = cv2.minAreaRect(cnts[0])
                box2 = np.int0(cv2.cv.BoxPoints(rect1))
                #cv2.drawContours(gray2,[box2],0,(0,0,255),2)
                M = cv2.moments(cnts[0])
                distance1 = math.sqrt(((x+w/2+boxes[0][0])-x0)**2 + ((y+h/2+boxes[0][1])-y0)**2)
                if i == 1:
                    data1.append([x+w/2+boxes[0][0], y+h/2+boxes[0][1], x, y, w, h, 0, 0, 0])
                    x5.append(float(x+w/2)/float((int(boxes[1][0])-int(boxes[0][0]))))
                    y5.append(float(y+h/2)/float((int(boxes[1][1])-int(boxes[0][1]))))
                    z5.append(distance1)
                else:
                    data1.append([x+w/2+boxes[0][0], y+h/2+boxes[0][1], x, y, w, h, 0, round(distance1),int(M['m00'])])
                    x5.append(float(x+w/2)/float((int(boxes[1][0])-int(boxes[0][0]))))
                    y5.append(float(y+h/2)/float((int(boxes[1][1])-int(boxes[0][1]))))
                    z5.append(distance1)
                    if round(distance1) > tempD2:
                        tempD1 = tempD1 + 1
                        tempD.append([tempD1, i, int(distance1)])
            cv2.rectangle(gray2,(x-2+boxes[0][0],y-2++boxes[0][1]),(x+w+2+boxes[0][0],y+h+2++boxes[0][1]),(0,255,0),2)
            cv2.circle(gray2, (x+w/2+boxes[0][0], y+h/2+boxes[0][1]), 10, (0,0,255),-1)
            cv2.imshow('Tracking ...',gray2)
            x0 = x+w/2+boxes[0][0]
            y0 = y+h/2+boxes[0][1]
            if i%1000 ==0:
                print "Processing ..." + str(i) + "/" + str(int(total))
            if int(OriginalFrameRate) == 0:
                data1Stamp.append([i,str((time() - starttime))])
                print OriginalFrameRate
            else:
                data1Stamp.append([i,i*(1/OriginalFrameRate)])

            frametime = datetime.datetime.now()
            elapsed_time = 0
            while (elapsed_time < (setframerate-0.022)):  # This is to control the frame rate 9.5 means 95 ms and this will give you about 10.45 fps(1000/95)
                elapsed_time = (datetime.datetime.now() - frametime).total_seconds()
        else:
            cv2.imshow('Tracking ...',gray2)
        i=i+1
        cv2.setMouseCallback('Tracking ...',set_up_ROIs)
        if cv2.waitKey(dummy) & 0xFF == ord('q'):
            sys.exit()
            save1=1
    endtime =  time()
    print "total processing time is ..." + str(int(endtime - starttime)) + " seconds"
    print str(int(total/(endtime - starttime))) + ' frame per second'
    cap.release()
    CurrentTime = strftime("%H",) + strftime("%M",) + strftime("%S",)
    CurrentDate = strftime("%y",) + strftime("%m",) + strftime("%d",)
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Tank1")
    rows = len(data1)
    sheet1.write(0, 0, 'Processing Date')
    sheet1.write(0, 2, CurrentDate)
    sheet1.write(0, 3, 'Processing Time')
    sheet1.write(0, 5, CurrentTime)
    sheet1.write(1, 3, 'OnFrame')
    sheet1.write(1, 4, 'OffFrame')
    sheet1.write(2, 3, OnFrame)
    sheet1.write(2, 4, OffFrame)
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
    sheet1.write(9, 4, 'RectX')
    sheet1.write(9, 5, 'RectY')
    sheet1.write(9, 6, 'RectW')
    sheet1.write(9, 7, 'RectH')
    sheet1.write(9, 8, 'ValveFlag')
    sheet1.write(9, 9, 'Distance')
    sheet1.write(9, 10, 'Area')
    sheet1.write(9, 11, 'Up(0) - Down(1)')
    sheet1.write(9, 12, 'Left(0) - Center(1) - Right (2)')

    for row in range(rows):
        for col in range(1):
            sheet1.write(row+10, col+0, data1Stamp[row][col])
            sheet1.write(row+10, col+1, str(data1Stamp[row][col+1]))
            sheet1.write(row+10, col+2, data1[row][col])
            if (data1[row][col+0] + boxes[0][0]) > (2 * (int((boxes[1][0] - boxes[0][0])/3)) + boxes[0][0]):
                tempRight = tempRight + 1
                sheet1.write(row+10, col+12, 2)
            else:
                if (data1[row][col+0] + boxes[0][0]) <= (int((boxes[1][0] - boxes[0][0])/3) + boxes[0][0]):
                    tempLeft = tempLeft + 1
                    sheet1.write(row+10, col+12, 0)
                else:
                    tempCenter = tempCenter + 1
                    sheet1.write(row+10, col+12, 1)
            sheet1.write(row+10, col+3, data1[row][col+1])
            if (data1[row][col+1]+ boxes[0][1]) > (int((boxes[1][1] - boxes[0][1])/2) + boxes[0][1]):
                tempDown = tempDown + 1
                sheet1.write(row+10, col+11, 1)
            else:
                tempUp = tempUp + 1
                sheet1.write(row+10, col+11, 0)
            sheet1.write(row+10, col+4, data1[row][col+2])
            sheet1.write(row+10, col+5, data1[row][col+3])
            sheet1.write(row+10, col+6, data1[row][col+4])
            sheet1.write(row+10, col+7, data1[row][col+5])
            sheet1.write(row+10, col+8, data1[row][col+6])
            sheet1.write(row+10, col+9, data1[row][col+7])
            sheet1.write(row+10, col+10, data1[row][col+8])
    if WhichTank ==1:
        if int(OnOff) == 2:
            print path
            print ExperimentID
            book.save(path + ExperimentID + "Tank1_OffLine_Raw.xls")
        else:
            book.save(path + ExperimentID + "Tank1_OffLine_Raw.xls")
    else:
        if int(OnOff) == 2:
            book.save(path + ExperimentID + "Tank1_OffLine_Raw.xls")
        else:
            book.save(path + ExperimentID + "Tank2_OffLine_Raw.xls")

    print 'Number of No-Fish Frame is: ' + str(tempZero1)

    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Zero List")
    for i in range(tempZero1):
        for j in range(1):
            sheet1.write(i+1, 0, tempZero[i][0])
            sheet1.write(i+1, 1, tempZero[i][1])
    sheet1.write(0, 0, 'Item#')
    sheet1.write(0, 1, 'Frame#')
    if tempZero1 <= 1:
        if WhichTank ==1:
            if ExcelStatus ==0:
                if int(OnOff) == 2:
                    book.save(path + ExperimentID + "Tank1_OffLine_Zero List_n=" + str(tempZero1) + "_frame.xls")
                else:
                    book.save(path + ExperimentID + "Tank1_OffLine_Zero List_n=" + str(tempZero1) + "_frame.xls")
            else:
                book.save(path + ExperimentID + "Tank1_OnLine_Zero List_n=" + str(tempZero1) + "_frame.xls")
        else:
            if ExcelStatus ==0:
                if int(OnOff) == 2:
                    book.save(path + ExperimentID + "Tank2_OffLine_Zero List_n=" + str(tempZero1) + "_frame.xls")
                else:
                    book.save(path + ExperimentID + "Tank2_OffLine_Zero List_n=" + str(tempZero1) + "_frame.xls")
            else:
                book.save(path + ExperimentID + "Tank2_OnLine_Zero List_n=" + str(tempZero1) + "_frame.xls")
    else:
        if WhichTank ==1:
            if ExcelStatus ==0:
                if int(OnOff) == 2:
                    book.save(path + ExperimentID + "Tank1_OffLine_Zero List_n=" + str(tempZero1) + "_frames.xls")
                else:
                    book.save(path + ExperimentID + "Tank1_OffLine_Zero List_n=" + str(tempZero1) + "_frames.xls")
            else:
                book.save(path + ExperimentID + "Tank1_OnLine_Zero List_n=" + str(tempZero1) + "_frames.xls")
        else:
            if ExcelStatus ==0:
                if int(OnOff) == 2:
                    book.save(path + ExperimentID + "Tank2_OffLine_Zero List_n=" + str(tempZero1) + "_frames.xls")
                else:
                    book.save(path + ExperimentID + "Tank2_OffLine_Zero List_n=" + str(tempZero1) + "_frames.xls")
            else:
                book.save(path + ExperimentID + "Tank2_OnLine_Zero List_n=" + str(tempZero1) + "_frames.xls")
    print '========================================'
    JumpFrame()
    Behavior1()
    Plotting1()