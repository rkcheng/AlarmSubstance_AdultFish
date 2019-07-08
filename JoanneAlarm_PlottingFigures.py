import sys
sys.path.append("C:\\opencv\\build\\python\\2.7")
import copy
import Tkinter, tkFileDialog, Tkconstants
from Tkinter import *
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.widgets import Button
from matplotlib.widgets import CheckButtons
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

# Darting threshold can be seen in line 280
# Freezing threshold can be seen in line 288
# Slow movment threshold can be seen in line 295

distance_in_seconds =[]
speed_before=[]
speed_after=[]
single_speed=[]
OffLine1=0
ExperimentID=[]
ExperimentID1=[]
ExperimentID2=[]
ExperimentID3=[]
ExperimentID4=[] # for different type of files
DartingBefore=[]
DartingAfter=[]
Darting1=[]
Darting2=[]
Darting3=[]
FreezingBefore=[]
FreezingAfter=[]
SlowBefore=[]
SlowAfter=[]
DownBefore=[]
DownAfter=[]
totaltime = []

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
                    book.save(path + "Tank1_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame(s).xls")
                else:
                    book.save(path + ExperimentID + "Tank1_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame(s).xls")
            else:
                if int(ExperimentID4[iii-1]) >= 2:
                    book.save(path + ExperimentID + "Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame(s).xls")
                else:
                    book.save(path + ExperimentID + "Tank1_Online_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame(s).xls")
        else:
            if ExcelStatus == 0:
                if int(OnOff) == 2:
                    book.save(path + "Tank2_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame(s).xls")
                else:
                    book.save(path + ExperimentID + "Tank2_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame(s).xls")
            else:
                if int(ExperimentID4[iii-1]) >= 2:
                    book.save(path + ExperimentID + "Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame(s).xls")
                else:
                    book.save(path + ExperimentID + "Tank2_Online_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame(s).xls")
    else:
        if WhichTank == 1:
            if ExcelStatus == 0:
                if int(OnOff) == 2:
                    book.save(path + "Tank1_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame.xls")
                else:
                    book.save(path + ExperimentID + "Tank1_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame.xls")
            else:
                if int(ExperimentID4[iii-1]) >= 2:
                    book.save(path + ExperimentID + "_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame.xls")
                else:
                    book.save(path + ExperimentID + "Tank1_Online_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame.xls")
        else:
            if ExcelStatus == 0:
                if int(OnOff) == 2:
                    book.save(path + "Tank2_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame.xls")
                else:
                    book.save(path + ExperimentID + "Tank2_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame.xls")
            else:
                if int(ExperimentID4[iii-1]) >= 2:
                    book.save(path + ExperimentID + "_Offline_Jump List_" + str(tempD2) + "_n=" + str(tempD1) + " frame.xls")
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
                book.save(path + "Tank1_Offline_Behavior Data.xls")
            else:
                book.save(path + ExperimentID + "Tank1_Offline_Behavior Data.xls")
        else:
            if int(ExperimentID4[iii-1]) ==1:
                book.save(path + ExperimentID + "Tank1_Online_Behavior Data.xls")
            else:
                book.save(path + ExperimentID + "_Offline_Behavior Data.xls")
    else:
        if ExcelStatus ==0:
            if int(OnOff) == 2:
                book.save(path + "Tank2_Offline_Behavior Data.xls")
            else:
                book.save(path + ExperimentID + "Tank2_Offline_Behavior Data.xls")
        else:
            if int(ExperimentID4[iii-1]) == 1:
                book.save(path + ExperimentID + "Tank2_Online_Behavior Data.xls")
            else:
                book.save(path + ExperimentID + "_Offline_Behavior Data.xls")
    # for Up-Down analysis ===================================
    print 'Percent of time in the top of the tank: ' + "{:.2%}".format(tempUp/(float(total-9)))
    print 'Percent of time in the bottom of the tank: ' + "{:.2%}".format(tempDown/(float(total-9)))
    print '========================================'
    # for Left-Center-Right analysis ============================
    print 'Percent of time in the left of the tank: ' + "{:.2%}".format(tempLeft/(float(total-9)))
    print 'Percent of time in the center of the tank: ' + "{:.2%}".format(tempCenter/(float(total-9)))
    print 'Percent of time in the right of the tank: ' + "{:.2%}".format(tempRight/(float(total-9)))
    return;

def Plotting3():
    fs = plt.figure(figsize = (18,5))
    gs = plt.GridSpec(2,5)
    ax1 = fs.add_subplot(gs[0,0])
    ax1.set_title('-50 seconds')
    x66 = x5[int(OnFrame-OriginalFrameRate*50):int(OnFrame-OriginalFrameRate*40)]
    y66 = y5[int(OnFrame-OriginalFrameRate*50):int(OnFrame-OriginalFrameRate*40)]
    colors = sns.color_palette('GnBu_d', (np.size(x66)/int(OriginalFrameRate))+12)
    count=10
    for ii in xrange(1,np.size(x66),int(OriginalFrameRate)):
        ax1.plot(x66[ii-1:ii+int(OriginalFrameRate)],y66[ii-1:ii+int(OriginalFrameRate)],linewidth=1, color=colors[count], alpha=0.5)
        count += 1
    ax1.set_xlim(0,1)
    ax1.set_ylim(0,1)
    ax = ax1.axis()
    ax1.axis((ax[0],ax[1],ax[3],ax[2]))

    ax2 = fs.add_subplot(gs[1,0])
    ax2.set_title('-40 seconds')
    x66 = x5[int(OnFrame-OriginalFrameRate*40):int(OnFrame-OriginalFrameRate*30)]
    y66 = y5[int(OnFrame-OriginalFrameRate*40):int(OnFrame-OriginalFrameRate*30)]
    colors = sns.color_palette('GnBu_d', (np.size(x66)/int(OriginalFrameRate))+12)
    count=10
    for ii in xrange(1,np.size(x66),int(OriginalFrameRate)):
        ax2.plot(x66[ii-1:ii+int(OriginalFrameRate)],y66[ii-1:ii+int(OriginalFrameRate)],linewidth=1, color=colors[count], alpha=0.5)
        count += 1
    ax2.set_xlim(0,1)
    ax2.set_ylim(0,1)
    ax2.axis((ax[0],ax[1],ax[3],ax[2]))

    ax3 = fs.add_subplot(gs[0,1])
    ax3.set_title('-30 seconds')
    x66 = x5[int(OnFrame-OriginalFrameRate*30):int(OnFrame-OriginalFrameRate*20)]
    y66 = y5[int(OnFrame-OriginalFrameRate*30):int(OnFrame-OriginalFrameRate*20)]
    colors = sns.color_palette('GnBu_d', (np.size(x66)/int(OriginalFrameRate))+12)
    count=10
    for ii in xrange(1,np.size(x66),int(OriginalFrameRate)):
        ax3.plot(x66[ii-1:ii+int(OriginalFrameRate)],y66[ii-1:ii+int(OriginalFrameRate)],linewidth=1, color=colors[count], alpha=0.5)
        count += 1
    ax3.set_xlim(0,1)
    ax3.set_ylim(0,1)
    ax3.axis((ax[0],ax[1],ax[3],ax[2]))



    ax4 = fs.add_subplot(gs[1,1])
    ax4.set_title('-20 seconds')
    x66 = x5[int(OnFrame-OriginalFrameRate*20):int(OnFrame-OriginalFrameRate*10)]
    y66 = y5[int(OnFrame-OriginalFrameRate*20):int(OnFrame-OriginalFrameRate*10)]
    colors = sns.color_palette('GnBu_d', (np.size(x66)/int(OriginalFrameRate))+12)
    count=10
    for ii in xrange(1,np.size(x66),int(OriginalFrameRate)):
        ax4.plot(x66[ii-1:ii+int(OriginalFrameRate)],y66[ii-1:ii+int(OriginalFrameRate)],linewidth=1, color=colors[count], alpha=0.5)
        count += 1
    ax4.set_xlim(0,1)
    ax4.set_ylim(0,1)
    ax4.axis((ax[0],ax[1],ax[3],ax[2]))

    ax5 = fs.add_subplot(gs[0,2])
    ax5.set_title('-10 seconds')
    x66 = x5[int(OnFrame-OriginalFrameRate*10):int(OnFrame-OriginalFrameRate*0)]
    y66 = y5[int(OnFrame-OriginalFrameRate*10):int(OnFrame-OriginalFrameRate*0)]
    colors = sns.color_palette('GnBu_d', (np.size(x66)/int(OriginalFrameRate))+12)
    count=10
    for ii in xrange(1,np.size(x66),int(OriginalFrameRate)):
        ax5.plot(x66[ii-1:ii+int(OriginalFrameRate)],y66[ii-1:ii+int(OriginalFrameRate)],linewidth=1, color=colors[count], alpha=0.5)
        count += 1
    ax5.set_xlim(0,1)
    ax5.set_ylim(0,1)
    ax5.axis((ax[0],ax[1],ax[3],ax[2]))



    ax6 = fs.add_subplot(gs[1,2])
    ax6.set_title('+10 seconds')
    x66 = x5[int(OnFrame+OriginalFrameRate*0):int(OnFrame+OriginalFrameRate*10)]
    y66 = y5[int(OnFrame+OriginalFrameRate*0):int(OnFrame+OriginalFrameRate*10)]
    colors = sns.color_palette('YlGn_d', (np.size(x66)/int(OriginalFrameRate))+12)
    count=10
    for ii in xrange(1,np.size(x66),int(OriginalFrameRate)):
        ax6.plot(x66[ii-1:ii+int(OriginalFrameRate)],y66[ii-1:ii+int(OriginalFrameRate)],linewidth=1, color=colors[count], alpha=0.5)
        count += 1
    ax6.set_xlim(0,1)
    ax6.set_ylim(0,1)
    ax6.axis((ax[0],ax[1],ax[3],ax[2]))

    ax7 = fs.add_subplot(gs[0,3])
    ax7.set_title('+20 seconds')
    x66 = x5[int(OnFrame+OriginalFrameRate*10):int(OnFrame+OriginalFrameRate*20)]
    y66 = y5[int(OnFrame+OriginalFrameRate*10):int(OnFrame+OriginalFrameRate*20)]
    colors = sns.color_palette('YlGn_d', (np.size(x66)/int(OriginalFrameRate))+12)
    count=10
    for ii in xrange(1,np.size(x66),int(OriginalFrameRate)):
        ax7.plot(x66[ii-1:ii+int(OriginalFrameRate)],y66[ii-1:ii+int(OriginalFrameRate)],linewidth=1, color=colors[count], alpha=0.5)
        count += 1
    ax7.set_xlim(0,1)
    ax7.set_ylim(0,1)
    ax7.axis((ax[0],ax[1],ax[3],ax[2]))



    ax8 = fs.add_subplot(gs[1,3])
    ax8.set_title('+30 seconds')
    x66 = x5[int(OnFrame+OriginalFrameRate*20):int(OnFrame+OriginalFrameRate*30)]
    y66 = y5[int(OnFrame+OriginalFrameRate*20):int(OnFrame+OriginalFrameRate*30)]
    colors = sns.color_palette('YlGn_d', (np.size(x66)/int(OriginalFrameRate))+12)
    count=10
    for ii in xrange(1,np.size(x66),int(OriginalFrameRate)):
        ax8.plot(x66[ii-1:ii+int(OriginalFrameRate)],y66[ii-1:ii+int(OriginalFrameRate)],linewidth=1, color=colors[count], alpha=0.5)
        count += 1
    ax8.set_xlim(0,1)
    ax8.set_ylim(0,1)
    ax8.axis((ax[0],ax[1],ax[3],ax[2]))


    ax9 = fs.add_subplot(gs[0,4])
    ax9.set_title('+40 seconds')
    x66 = x5[int(OnFrame+OriginalFrameRate*30):int(OnFrame+OriginalFrameRate*40)]
    y66 = y5[int(OnFrame+OriginalFrameRate*30):int(OnFrame+OriginalFrameRate*40)]
    colors = sns.color_palette('YlGn_d', (np.size(x66)/int(OriginalFrameRate))+12)
    count=10
    for ii in xrange(1,np.size(x66),int(OriginalFrameRate)):
        ax9.plot(x66[ii-1:ii+int(OriginalFrameRate)],y66[ii-1:ii+int(OriginalFrameRate)],linewidth=1, color=colors[count], alpha=0.5)
        count += 1
    ax9.set_xlim(0,1)
    ax9.set_ylim(0,1)
    ax9.axis((ax[0],ax[1],ax[3],ax[2]))


    ax10 = fs.add_subplot(gs[1,4])
    ax10.set_title('+50 seconds')
    x66 = x5[int(OnFrame+OriginalFrameRate*40):int(OnFrame+OriginalFrameRate*50)]
    y66 = y5[int(OnFrame+OriginalFrameRate*40):int(OnFrame+OriginalFrameRate*50)]
    colors = sns.color_palette('YlGn_d', (np.size(x66)/int(OriginalFrameRate))+12)
    count=10
    for ii in xrange(1,np.size(x66),int(OriginalFrameRate)):
        ax10.plot(x66[ii-1:ii+int(OriginalFrameRate)],y66[ii-1:ii+int(OriginalFrameRate)],linewidth=1, color=colors[count], alpha=0.5)
        count += 1
    ax10.set_xlim(0,1)
    ax10.set_ylim(0,1)
    ax10.axis((ax[0],ax[1],ax[3],ax[2]))


    if iii % 2 == 1:
        fs.suptitle(ExperimentID + 'Tank1 10-Second Blocks')
    else:
        fs.suptitle(ExperimentID + 'Tank2 10-Second Blocks')

    class Index2(object):
        def next2(self, event):
            if WhichTank == 1:
                if int(ExperimentID4[iii-1]) > 2:
                    plt.savefig(path +ExperimentID + '_Offline.pdf')
                else:
                    plt.savefig(path +ExperimentID + 'Tank1_Online_10-Second Blocks.pdf')
            else:
                if int(ExperimentID4[iii-1]) > 2:
                    plt.savefig(path +ExperimentID + '_Offline.pdf')
                else:
                    plt.savefig(path +ExperimentID + 'Tank2_Online_10-Second Blocks.pdf')
            plt.close('all')
            if iii == iiii:
                Plotting2()

    callback = Index2()
    axnext = plt.axes([0.915, 0.01, 0.085, 0.05])
    bnext = Button(axnext, 'Save and Close')
    bnext.on_clicked(callback.next2)
    cv2.destroyAllWindows()
    plt.show()  # show for plotting3

def Plotting1(ExperimentID2,iii,ExperimentID4,OffTime,iiii,array1,OriginalFrameRate):
    if int(OnOff) == 2: # for offline tracking
        iii=1
        iiii=iii
        ExperimentID2=[]
        ExperimentID4=[]
        ExperimentID2.append('True')
        ExperimentID4.append('1')
        OffTime = 600
    if ExperimentID2[iii-1] == 'True':
        if ExperimentID4[iii-1] != '3':
            label_in_seconds =[]
            for m in range (1,OffFrame+1,1):
                #label_in_seconds.append(totaltime*(1/float(OffFrame+1))*m)
                label_in_seconds.append(600*(1/float(OffFrame+1))*m)
            label_in_seconds = np.divide(label_in_seconds, 60.)
        else:
            if OffLine1 == 1:
                OriginalFrameRate = 26
                label_in_seconds =[]
                for m in range (1,OffFrame+1,1):
                    #label_in_seconds.append(totaltime*(1/float(OffFrame+1))*m)
                    label_in_seconds.append(600*(1/float(OffFrame+1))*m)
                label_in_seconds = np.divide(label_in_seconds, 60.)
                #label_in_seconds = np.divide(array1, float(OffFrame/300))
            else:
                label_in_seconds = np.divide(array1, 60.)
        #fs = plt.figure(figsize = (18,5))
        #gs = plt.GridSpec(2,2,height_ratios=[2,1])
        fs = plt.figure(figsize = (18,12))
        gs = plt.GridSpec(4,2,height_ratios=[2,1,1,1])
        ax1 = fs.add_subplot(gs[0,0])

        z6 = z5[1:OnFrame]
        z66 = z55[1:OnFrame]
        z666 = z555[1:OnFrame]
        z666c = temp555c[int(OnFrame-OriginalFrameRate*60):OnFrame]
        turningbe = int(math.fsum(z666c))
        z6 = np.asarray(z6) * float(200) / float(int(boxes[1][0])-int(boxes[0][0]))
        z66 = np.asarray(z66) * float(200) / float(int(boxes[1][0])-int(boxes[0][0]))
        z6mean = np.average(z6)
        z66mean = np.average(z66)
        z666mean = np.average(z666)
        ax1.set_title('Before Stimulus Speed = ' + "{:.3}".format(str(z6mean)) + " mm/frame")
        ax2 = fs.add_subplot(gs[1, 0:3])
        ax3 = fs.add_subplot(gs[2, 0:3])
        ax4 = fs.add_subplot(gs[3, 0:3])
        x6 = x5[1:OnFrame]
        y6 = y5[1:OnFrame]
        OnTimeSec = int(OnFrame / OriginalFrameRate)
        z6x = label_in_seconds[range(1, OnFrame)]
        z66x = label_in_seconds[range(1, OnFrame)]
        z666x = label_in_seconds[range(1, OnFrame)]
        colors = sns.color_palette('GnBu_d', (np.size(x6)/int(OriginalFrameRate))+12)
        count=10
        for ii in xrange(1,np.size(x6),int(OriginalFrameRate)):
            ax1.plot(x6[ii-1:ii+int(OriginalFrameRate)],y6[ii-1:ii+int(OriginalFrameRate)],linewidth=3, color=colors[count], alpha=0.5)
            ax2.plot(z6x[ii:ii+int(OriginalFrameRate)],z6[ii:ii+int(OriginalFrameRate)],linewidth=2, color=colors[count], alpha=0.5)
            ax3.plot(z66x[ii:ii+int(OriginalFrameRate)],z66[ii:ii+int(OriginalFrameRate)],linewidth=2, color=colors[count], alpha=0.5)
            ax4.plot(z666x[ii:ii+int(OriginalFrameRate)],z666[ii:ii+int(OriginalFrameRate)],linewidth=1, color=colors[count], alpha=0.5)
            count += 1
        ax1.set_xlim(0,1)
        ax1.set_ylim(0,1)
        ax = ax1.axis()
        ax1.axis((ax[0],ax[1],ax[3],ax[2]))


        ax1 = fs.add_subplot(gs[0,1])
        z6 = z5[OnFrame:OffFrame]
        z66 = z55[OnFrame:OffFrame]
        z666 = z555[OnFrame:OffFrame]
        z666c = temp555c[OnFrame:int(OnFrame+OriginalFrameRate*60)]
        turningaf = int(math.fsum(z666c))

        z6 = np.asarray(z6) * float(200) / float(int(boxes[1][0])-int(boxes[0][0]))
        z66 = np.asarray(z66) * float(200) / float(int(boxes[1][0])-int(boxes[0][0]))
        z6mean = np.average(z6)
        z66mean = np.average(z66)
        z666mean = np.average(z666)

        ax1.set_title('After Stimulus Speed = ' + "{:.3}".format(str(z6mean)) + " mm/frame")
        x6 = x5[OnFrame:OffFrame]
        y6 = y5[OnFrame:OffFrame]

        z6x = label_in_seconds[range(OnFrame,OffFrame)]
        z66x = label_in_seconds[range(OnFrame,OffFrame)]
        z666x = label_in_seconds[range(OnFrame,OffFrame)]

        colors = sns.color_palette('YlGn_d', (np.size(x6)/int(OriginalFrameRate))+12)
        #colors = np.flip(colors,0)
        count=10

        for ii in xrange(1,np.size(x6),int(OriginalFrameRate)):
            ax1.plot(x6[ii-1:ii+int(OriginalFrameRate)],y6[ii-1:ii+int(OriginalFrameRate)],linewidth=3, color=colors[count], alpha=0.5)
            ax2.plot(z6x[ii:ii+int(OriginalFrameRate)],z6[ii:ii+int(OriginalFrameRate)],linewidth=2, color=colors[count], alpha=0.5)
            ax3.plot(z66x[ii:ii+int(OriginalFrameRate)],z66[ii:ii+int(OriginalFrameRate)],linewidth=2, color=colors[count], alpha=0.5)
            ax4.plot(z666x[ii:ii+int(OriginalFrameRate)],z666[ii:ii+int(OriginalFrameRate)],linewidth=1, color=colors[count], alpha=0.5)
            count += 1
        ax1.set_xlim(0,1)
        ax1.set_ylim(0,1)
        ax = ax1.axis()
        ax1.axis((ax[0],ax[1],ax[3],ax[2]))


        #ax1 = fs.add_subplot(gs[0,2])
        z6 = z5[OffFrame:]
        z7=[]
        z77=[]
        z57=[]
        #z7 = z5[:OffFrame]
        z7 = z5[1:OnFrame]
        z6 = np.asarray(z6) * float(200) / float(int(boxes[1][0])-int(boxes[0][0]))
        z6mean = np.average(z6)
        for ii in xrange(1,np.size(z7),int(OriginalFrameRate)):
            z77.append(np.sum(z7[ii-1:ii+int(OriginalFrameRate)]))
        for ii in xrange(1,np.size(z5),int(OriginalFrameRate)):
            z57.append(np.sum(z5[ii-1:ii+int(OriginalFrameRate)]))
        z77 = np.asarray(z77) * float(200) / float(int(boxes[1][0])-int(boxes[0][0]))
        z57 = np.asarray(z57) * float(200) / float(int(boxes[1][0])-int(boxes[0][0]))

        #z57=z57[0:OnTime]
        #z77=z77[0:OnTime]
        z57=z57[0:OffTime]
        z77=z77[0:299]

        z7mean = np.average(z77)
        z7std = np.std(z77)
        z7darting = 0
        z7dartingbe = 0
        z7dartingaf = 0
        z7freezing1 = 0
        z7freezingbe = 0
        z7freezingaf = 0
        z7freezing2 = 0
        z7slow1 = 0
        z7slowbe = 0
        z7slowaf = 0
        Darting0 = 0
        for z in range (1,len(z57),1):  # changed from z77 to z57
            if z57[z] > z7mean + 5 * z7std:
                if z < OnTimeSec:
                    z7dartingbe +=1
                else:
                    z7dartingaf +=1
                z7darting += 1
                Darting0 += 1
                Darting1.append(z)
            elif z57[z] < z7mean/10:
                z7freezing1 +=1
                if z < OnTimeSec:
                    z7freezingbe +=1
                else:
                    z7freezingaf +=1
            else:
                if z57[z] < z7mean/2:
                    z7slow1 +=1
                    if z < OnTimeSec:
                        z7slowbe +=1
                    else:
                        z7slowaf +=1
        FreezingBefore.append(z7freezingbe)
        FreezingAfter.append(z7freezingaf)
        DartingBefore.append(z7dartingbe)
        DartingAfter.append(z7dartingaf)
        SlowBefore.append(z7slowbe)
        SlowAfter.append(z7slowaf)
        kk1 = format(tempDownBe/float(OnFrame),'.3f')
        kk2 = format(tempDownAf/float(OffFrame-OnFrame),'.3f')
        DownBefore.append(float(kk1))
        DownAfter.append(float(kk2))
        Darting2.append(z7dartingbe+z7dartingaf)

        if iii % 2 == 1:
            if int(ExperimentID4[iii-1]) ==3:
                fs.suptitle(ExperimentID + ', darting = (' + str(z7dartingbe) + '-' + str(z7dartingaf) + '), freezing = (' + str(z7freezingbe) + '-' + str(z7freezingaf) + '), bottom1/3 = (' + str(kk1) + '-' + str(kk2) + '), turning = (' + str(turningbe) + '-' + str(turningaf) + ')', fontsize=14, fontweight='bold')
            else:
                if int(OnOff) == 2: # for offline tracking
                    fs.suptitle(ExperimentID + ', darting = (' + str(z7dartingbe) + '-' + str(z7dartingaf) + '), freezing = (' + str(z7freezingbe) + '-' + str(z7freezingaf) + '), bottom1/3 = (' + str(kk1) + '-' + str(kk2) + '), turning = (' + str(turningbe) + '-' + str(turningaf) + ')', fontsize=14, fontweight='bold')
                else:
                    #fs.suptitle(ExperimentID + 'Tank1' + ', darting = (' + str(z7dartingbe) + '-' + str(z7dartingaf) + '), freezing = (' + str(z7freezingbe) + '-' + str(z7freezingaf) + '), bottom1/3 = (' + str(float(tempDownBe/OnFrame)) + '-' + str(tempDownAf) + '), turning = (' + str(turningbe) + '-' + str(turningaf) + ')', fontsize=14, fontweight='bold')
                    fs.suptitle(ExperimentID + 'Tank1' + ', darting = (' + str(z7dartingbe) + '-' + str(z7dartingaf) + '), freezing = (' + str(z7freezingbe) + '-' + str(z7freezingaf) + '), bottom1/3 = (' + str(kk1) + '-' + str(kk2) + '), turning = (' + str(turningbe) + '-' + str(turningaf) + ')', fontsize=14, fontweight='bold')
        else:
            if int(ExperimentID4[iii-1]) ==3:
                fs.suptitle(ExperimentID  + ', darting = (' + str(z7dartingbe) + '-' + str(z7dartingaf) + '), freezing = (' + str(z7freezingbe) + '-' + str(z7freezingaf) + '), bottom1/3 = (' + str(kk1) + '-' + str(kk2) + '), turning = (' + str(turningbe) + '-' + str(turningaf) + ')', fontsize=14, fontweight='bold')
            else:
                fs.suptitle(ExperimentID + 'Tank2' + ', darting = (' + str(z7dartingbe) + '-' + str(z7dartingaf) + '), freezing = (' + str(z7freezingbe) + '-' + str(z7freezingaf) + '), bottom1/3 = (' + str(kk1) + '-' + str(kk2) + '), turning = (' + str(turningbe) + '-' + str(turningaf) + ')', fontsize=14, fontweight='bold')

        #ax1.set_title('After Stimulus Speed = ' + "{:.3}".format(str(z6mean)) + " mm/frame")
        #x6 = x5[OffFrame:]
        #y6 = y5[OffFrame:]
        #if int(OnOff) == 2:
        #    z6x = label_in_seconds[range(OffFrame,np.size(array1)-1)]
        #else:
        #    z6x = label_in_seconds[range(OffFrame,np.size(array1))]

        #colors = sns.color_palette('YlOrRd', (np.size(x6)/int(OriginalFrameRate))+12)
        #count=10
        #for ii in xrange(1,np.size(x6),int(OriginalFrameRate)):
            #ax1.plot(x6[ii-1:ii+int(OriginalFrameRate)],y6[ii-1:ii+int(OriginalFrameRate)],linewidth=3, color=colors[count], alpha=0.5)
            #ax2.plot(z6x[ii-1:ii+int(OriginalFrameRate)],z6[ii-1:ii+int(OriginalFrameRate)],linewidth=2, color=colors[count], alpha=0.5)
            #count += 1
        ax1.set_xlim(0,1)
        ax1.set_ylim(0,1)
        ax = ax1.axis()
        ax1.axis((ax[0],ax[1],ax[3],ax[2]))

        #ax2.set_xlabel('Time (minutes)')
        ax2.set_ylabel('Distance (mm/fr)')
        #ax3.set_ylabel('Acceleration (mm/fr2)')
        ax4.set_xlabel('Time (minutes)')
        #ax4.set_ylabel('+1:To Right')

        ax2.set_ylim((0, 40))
        major_ticks = np.arange(0, 41, 8)
        ax2.set_yticks(major_ticks)
        ax2.set_xlim((0, 10))
        major_ticks = np.arange(0, 11, 5)
        ax2.set_xticks(major_ticks)

        ax3.set_ylim((-25, 25))
        major_ticks = np.arange(-25, 26, 10)
        ax3.set_yticks(major_ticks)
        ax3.set_xlim((0, 10))
        major_ticks = np.arange(0, 11, 5)
        ax3.set_xticks(major_ticks)

        ax4.set_ylim((-1.5, 1.5))
        major_ticks = np.arange(-1.5,1.6, 1)
        ax4.set_yticks(major_ticks)
        ax4.set_xlim((0, 10))
        major_ticks = np.arange(0, 11, 5)
        ax4.set_xticks(major_ticks)

        ax2.minorticks_on()
        ax3.minorticks_on()
        ax4.minorticks_on()
        ax2.tick_params(axis='x', which='minor', direction='out',length=8, width = 1)
        ax2.tick_params(axis='x', which='major', direction='out',length=12, width = 1)
        ax3.tick_params(axis='x', which='minor', direction='out',length=8, width = 1)
        ax3.tick_params(axis='x', which='major', direction='out',length=12, width = 1)
        ax4.tick_params(axis='x', which='minor', direction='out',length=8, width = 1)
        ax4.tick_params(axis='x', which='major', direction='out',length=12, width = 1)

    class Index1(object):
        def next1(self, event):
            #plt.savefig(path + ExperimentID[:-7] + '_Histogram.pdf')
            plt.savefig(path + '_Histogram.pdf')
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
                    if int(ExperimentID4[iii-1]) > 2:
                        plt.savefig(path +ExperimentID + '_Offline.pdf')
                    else:
                        plt.savefig(path +ExperimentID + 'Tank1_Online.pdf')
                else:
                    if int(ExperimentID4[iii-1]) > 2:
                        plt.savefig(path +ExperimentID + '_Offline.pdf')
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
                                Plotting2()
                else:
                    if int(AllOrNone) == 2:
                        plt.close('all')
                        sys.exit()
                    else:
                        plt.close('all')
            else:
                plt.close('all')
                sys.exit()
    if iii==iiii:
        if ExperimentID2[iii-1] == 'False':
            Plotting2()
            callback = Index()
            axnext = plt.axes([0.915, 0.01, 0.085, 0.05])
            bnext = Button(axnext, 'Save and Close')
            bnext.on_clicked(callback.next)
            cv2.destroyAllWindows()
            plt.show()
        else:
            callback = Index()
            axnext = plt.axes([0.915, 0.01, 0.085, 0.05])
            bnext = Button(axnext, 'Save and Close')
            bnext.on_clicked(callback.next)
            cv2.destroyAllWindows()
            plt.show()
            Plotting2()
    else:
        callback = Index()
        axnext = plt.axes([0.915, 0.01, 0.085, 0.05])
        bnext = Button(axnext, 'Save and Close')
        bnext.on_clicked(callback.next)
        cv2.destroyAllWindows()
        plt.show() # show for plotting1

def Plotting2():
    class Index1(object):
        def next1(self, event):
            #plt.savefig(path + ExperimentID[:-7] + 'Histogram.pdf')
            plt.savefig(path + 'Histogram.pdf')
            plt.close('all')
            sys.exit()

    plt.close('all')
    fs = plt.figure(figsize = (13,6))
    gs = plt.GridSpec(2,5)
    ax1 = fs.add_subplot(gs[0,:-3])
    ax2 = ax1.twinx()
    ax3 = fs.add_subplot(gs[1,:-3])
    ax4 = fs.add_subplot(gs[1,2])
    ax5 = fs.add_subplot(gs[1,3])
    ax6 = fs.add_subplot(gs[1,4])
    color = [sns.color_palette()[0], sns.color_palette()[2]]
    label = ['Before', 'After']
    ax1.hist(speed_before, np.size(speed_before), density=2, color = 'b', facecolor=color[0], histtype = 'bar', alpha=0.5, label=['Before'])
    sns.distplot(speed_before, np.size(speed_before), hist=False, kde=True,ax=ax2, color = 'b')
    ax3.hist(speed_before, np.size(speed_before), density=1, color = 'b', facecolor=color[0], histtype = 'step', cumulative = True, lw=3, label=['Before'])
    ax1.hist(speed_after, np.size(speed_after), density=2, color = 'g', facecolor=color[1], histtype = 'bar', alpha=0.5, label=['After'])
    sns.distplot(speed_after, np.size(speed_after), hist=False, kde=True,ax=ax2, color = 'g')
    ax3.hist(speed_after, np.size(speed_after), density=1, color = 'g', facecolor=color[1], histtype = 'step', cumulative = True, lw=3, label=['After'])
    label1 = ['Before','After']

    data=[]
    x=[]
    data = [DartingBefore, DartingAfter]
    ax4.boxplot(data, 0, '')
    for i in [1,2]:
        if i == 1:
            x = np.random.normal(i, 0.08, size=len(DartingBefore))
            ax4.plot(x, DartingBefore, 'b.', alpha=0.8)
        else:
            x = np.random.normal(i, 0.08, size=len(DartingAfter))
            ax4.plot(x, DartingAfter, 'r.', alpha=0.8)
    ax4.set_xlabel('Darting')
    ax4.set_ylabel('Seconds')
    ax4.set_xticklabels(label1,rotation=0, fontsize=10)
    data=[]
    x=[]
    data = [FreezingBefore, FreezingAfter]
    ax5.boxplot(data, 0, '')
    for i in [1,2]:
        if i == 1:
            x = np.random.normal(i, 0.08, size=len(FreezingBefore))
            ax5.plot(x, FreezingBefore, 'b.', alpha=0.8)
        else:
            x = np.random.normal(i, 0.08, size=len(FreezingAfter))
            ax5.plot(x, FreezingAfter, 'r.', alpha=0.8)
    ax5.set_xlabel('Freezing')
    ax5.set_ylabel('Seconds')
    ax5.set_xticklabels(label1,rotation=0, fontsize=10)

    # Change below from Slow to Down
    data=[]
    x=[]
    data = [DownBefore, DownAfter]
    ax6.boxplot(data, 0, '')
    for i in [1,2]:
        if i == 1:
            x = np.random.normal(i, 0.08, size=len(DownBefore))
            ax6.plot(x, DownBefore, 'b.', alpha=0.8)
        else:
            x = np.random.normal(i, 0.08, size=len(DownAfter))
            ax6.plot(x, DownAfter, 'r.', alpha=0.8)
    ax6.set_xlabel('Bottom 1/3')
    ax6.set_ylabel('Ratio')
    ax6.set_xticklabels(label1,rotation=0, fontsize=10)

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
    axnext = plt.axes([0.46, 0.9, 0.10, 0.05])
    bnext = Button(axnext, 'Save and Close')
    bnext.on_clicked(callback.next1)
    rax1 = plt.axes([0.6, 0.51, 0.39, 0.48])
    ExperimentID4 =[]
    NumberFish=0
    for o in range (0,len(ExperimentID2),1):
        if ExperimentID2[o] == 'True':
            NumberFish +=1
            ExperimentID3.append(ExperimentID1[o])
            ExperimentID4.append('True')
    ExperimentID3.append('n = ' + str(NumberFish) + ' fish')
    ExperimentID4.append('True')
    Behavior2()
    check1 = CheckButtons(rax1, ExperimentID3, ExperimentID4)
    plt.draw()
    check1.on_clicked(func)
    plt.show()

def Behavior2():
    NumberFish = 0
    for o in range (0,len(ExperimentID2),1):
        if ExperimentID2[o] == 'True':
            NumberFish +=1
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Behavior List")
    sheet1.write(0, 0, 'FishID')
    sheet1.write(0, 1, 'Darting Before')
    sheet1.write(0, 2, 'Darting After')
    sheet1.write(0, 3, 'Freezing Before')
    sheet1.write(0, 4, 'Freezing After')
    sheet1.write(0, 5, 'Slow Swim Before')
    sheet1.write(0, 6, 'Slow Swim After')
    sheet1.write(0, 7, 'Bottom1/3 Before')
    sheet1.write(0, 8, 'Bottom1/3 After')
    sheet1.write(0, 9, 'All darting time (sec)')
    Darting3 = 0
    for p in range (0, NumberFish,1):
        sheet1.write(1+p, 0, ExperimentID3[p])
        sheet1.write(1+p, 1, DartingBefore[p])
        sheet1.write(1+p, 2, DartingAfter[p])
        sheet1.write(1+p, 3, FreezingBefore[p])
        sheet1.write(1+p, 4, FreezingAfter[p])
        sheet1.write(1+p, 5, SlowBefore[p])
        sheet1.write(1+p, 6, SlowAfter[p])
        sheet1.write(1+p, 7, DownBefore[p])
        sheet1.write(1+p, 8, DownAfter[p])
        for q in range (0,Darting2[p],1):
            sheet1.write(1+p, 9+q, Darting1[Darting3])
            Darting3 +=1

    #book.save(path + ExperimentID[:-7] + "Histogram_Data.xls")
    book.save(path + "Histogram_Data.xls")

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
tempDownBe = 0
tempDownAf = 0
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
z55=[];
z555=[];
temp555b=0;
temp555c=[];
z6=[];

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
OnOff = raw_input("Analyze Existing (both Online and Offline) Excel data (1) or commence Offline tracking (2) -----> Please enter here: ")
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
    ExperimentID = root.filename1
    ExperimentID = ExperimentID[len(ExperimentID)-21:len(ExperimentID)-8]

    WhichTank = root.filename1[len(root.filename1)-9:len(root.filename1)-8]
    WhichTank=int(WhichTank)
    print WhichTank
    WhichTank2 = 4
    if WhichTank == 1:
        df = pd.read_excel(root.filename2, sheet_name='Tank1')
        column1 = df.as_matrix(columns=None)
        for i in range(3000,9000,1):
            if column1[i][8] + column1[i+1][8] == 1:
                OnFrame = i-7
                count = 1
                while ((int(column1[i+count][1])-int(column1[i][1]))<=30000):
                    count += 1
                OffFrame = i + count-7
                i = 9000
    else:
        df = pd.read_excel(root.filename2, sheet_name='Tank2')
        column1 = df.as_matrix(columns=None)
        for i in range(3000,9000,1):
            if column1[i][8] + column1[i+1][8] == 5:
                OnFrame = i-7
                count = 1
                while ((int(column1[i+count][1])-int(column1[i][1]))<=30000):
                    count += 1
                OffFrame = i + count-7
                i = 9000

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
    temp1, OriginalFrameRate1 = df.columns[6].split(":")
    OriginalFrameRate1 = OriginalFrameRate1[:2]
    OriginalFrameRate = float(OriginalFrameRate1)
    column1 = df.as_matrix(columns=None)
    for i in range(9,int(total+9),1):
        array1.append(float((int(column1[i][1])-int(column1[9][1]))/100))
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
        fs1 = plt.figure(figsize = (8,6))
        rax1 = plt.axes([0.01, 0.1, 0.9, 0.9])
        tempk=0
        for j in range (0, len(FileName),1):
            df = pd.read_excel(path + FileName[j], sheet_name='Tank1')
            column1 = df.as_matrix(columns=None)
            if column1[9][5] == 0:
                ExperimentID4.append('2')  # Old Online tracking data
                ExperimentID4.append('2')
            elif column1[0][3] == 'OnFrame':
                ExperimentID4.append('3')  # Offline tracking data
            elif column1[8][4] == 'RectX':
                OffLine1=1
                ExperimentID4.append('3')  # Offline tracking data
            else:
                ExperimentID4.append('1')  # New online tracking data
                ExperimentID4.append('1')
            if ExperimentID4[tempk] == '3':
                ExperimentID = FileName[j]
                tempk=tempk+1
                ExperimentID, temp1 = ExperimentID.split("OffLine_Raw")
                temp66 = ExperimentID[-3:-1]
                if int(temp66) % 2 == 1:
                    ExperimentID = str(1) + str(tempk).zfill(2) + '_' + ExperimentID + 'Tank-1'
                else:
                    ExperimentID = str(1) + str(tempk).zfill(2) + '_' + ExperimentID + 'Tank-2'
                ExperimentID1.append(ExperimentID)
                ExperimentID2.append('True')
            else:
                for k in range (0,2,1):
                    tempk=tempk+1
                    ExperimentID = FileName[j]
                    temp1, ExperimentID = ExperimentID.split("AlarmTrackingData_")
                    ExperimentID = ExperimentID[:-7]
                    if k == 0:
                        ExperimentID = str(1) + str(tempk).zfill(2) + '_' + ExperimentID + 'Tank-1'
                    else:
                        ExperimentID = str(1) + str(tempk).zfill(2) + '_' + ExperimentID + 'Tank-2'
                    ExperimentID1.append(ExperimentID)
                    ExperimentID2.append('True')

        check = CheckButtons(rax1, ExperimentID1, ExperimentID2)
        def func(label):
            for m in range (0,len(ExperimentID2),1):
                if label == ExperimentID1[m]:
                    if ExperimentID2[m] == 'True':
                        ExperimentID2[m] = 'False'
                    else:
                        ExperimentID2[m] = 'True'
            plt.draw()
        check.on_clicked(func)
        plt.show()
        #FirstOrBoth = raw_input("Do you want to analyze the first tank of each file (1) or both tanks (2) -----> Please enter here: ")
        FirstOrBoth = 2
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
            #iiii = 2*len(FileName)
            iiii = len(ExperimentID2)
            iiiii = 1
        else:
            iii = 1
            iiii = 2
    while (iii<= iiii):
        if iii % 2 == 1:
            if int(AllOrNone) == 1:
                if ExperimentID2[iii-1] == 'True':
                    df = pd.read_excel(path + FileName[iiiii-1], sheet_name='Tank1')
                    column1 = df.as_matrix(columns=None)
                    ExperimentID = FileName[iiiii-1]
                    if ExperimentID4[iii-1] == '3':
                        ExperimentID, temp1 = ExperimentID.split("_OffLine_Raw")
                        print 'Now analyzing ... ' + ExperimentID + ' ========================================'
                        iiiii += 1
                        if column1[1][3] > 10000:
                            OriginalFrameRate1 = 60
                        else:
                            OriginalFrameRate1 = 30
                    else:
                        temp1, ExperimentID = ExperimentID.split("AlarmTrackingData_")
                        ExperimentID = ExperimentID[:-7]
                        print 'Now analyzing ... ' + ExperimentID + ' Tank1 ========================================'
                        temp1, OriginalFrameRate1 = df.columns[6].split(":")
                        OriginalFrameRate1 = OriginalFrameRate1[:2]
                    WhichTank = 1
                    ValveFlag = 1
                else:
                    ExperimentID = FileName[iiiii-1]
                    if ExperimentID4[iii-1] == '3':
                        ExperimentID, temp1 = ExperimentID.split("OffLine_Raw")
                        print 'Now skipping ... ' + ExperimentID + '  ========================================'
                        iiiii += 1
                    else:
                        temp1, ExperimentID = ExperimentID.split("AlarmTrackingData_")
                        ExperimentID = ExperimentID[:-7]
                        print 'Now skipping ... ' + ExperimentID + ' Tank1 ========================================'
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
            if int(FirstOrBoth) == 1:
                print 'Now skipping ... ' + ExperimentID + ' Tank2 ========================================'
            else:
                if int(AllOrNone) == 1:
                    if ExperimentID2[iii-1] == 'True':
                        if ExperimentID4[iii-1] != '3':
                            df = pd.read_excel(path + FileName[iiiii-1], sheet_name='Tank2')
                            ExperimentID = FileName[iiiii-1]
                            iiiii += 1
                            temp1, ExperimentID = ExperimentID.split("AlarmTrackingData_")
                            ExperimentID = ExperimentID[:-7]
                            temp1, OriginalFrameRate1 = df.columns[6].split(":")
                            OriginalFrameRate1 = OriginalFrameRate1[:2]
                            column1 = df.as_matrix(columns=None)
                            print 'Now analyzing ... ' + ExperimentID + ' Tank2 ========================================'
                        else:
                            df = pd.read_excel(path + FileName[iiiii-1], sheet_name='Tank1')
                            ExperimentID = FileName[iiiii-1]
                            iiiii += 1
                            ExperimentID,temp1 = ExperimentID.split("_OffLine_Raw")
                            column1 = df.as_matrix(columns=None)
                            print 'Now analyzing ... ' + ExperimentID + '  ========================================'
                            if column1[1][3] > 10000:
                                OriginalFrameRate1 = 60
                            else:
                                OriginalFrameRate1 = 30

                        WhichTank = 2
                        ValveFlag = 5
                    else:
                        ExperimentID = FileName[iiiii-1]
                        iiiii += 1
                        if ExperimentID4[iii-1] == '3':
                            ExperimentID, temp1 = ExperimentID.split("OffLine_Raw")
                        else:
                            temp1, ExperimentID = ExperimentID.split("AlarmTrackingData_")
                            ExperimentID = ExperimentID[:-7]
                            print 'Now skipping ... ' + ExperimentID + ' Tank2 ========================================'
#                        if iii==iiii:
#                            print 'ending'
#                            print iii
#                            print iiii
#                            Plotting1()
                else:
                    df = pd.read_excel(root.filename2, sheet_name='Tank2')
                    ExperimentID = root.filename2
                    temp1, ExperimentID = ExperimentID.split("AlarmTrackingData_")
                    ExperimentID = ExperimentID[:-7]
                    temp1, OriginalFrameRate1 = df.columns[6].split(":")
                    OriginalFrameRate1 = OriginalFrameRate1[:2]
                    column1 = df.as_matrix(columns=None)

                    WhichTank = 2
                    ValveFlag = 5

        if ExperimentID2[iii-1] == 'True':
            if ExperimentID4[iii-1] == '1':   # New Online data
                if int(OriginalFrameRate1) > 30:
                    LowerLimit = 9000
                    UpperLimit = 11000
                    total = len(column1)
                    totaltime = (int(column1[total-1][1]) - int(column1[9][1]))
                    totaltime = int(totaltime/100)
                else:
                    LowerLimit = 3000
                    UpperLimit = 9000
                    total = len(column1)
                for i in range(LowerLimit,UpperLimit,1):
                    if column1[i][8] + column1[i+1][8] == ValveFlag:
                        OnFrame = i-7
                        OnTime = (int(column1[i][1]) - int(column1[9][1]))/100
                        count = 1
                        while ((int(column1[i+count][1])-int(column1[i][1]))<=30000):
                            if count + OnFrame < total-9:
                                count += 1
                            else:
                                break
                        OffFrame = i + count-7
                        OffTime = (int(column1[i+count][1]) - int(column1[9][1]))/100
                        i = 9000
            elif ExperimentID4[iii-1] == '2':
                for i in range(17000,20000,1):
                    if column1[i][5] + column1[i+1][5] == 1:
                        OnFrame = i-7
                        OnTime = (int(column1[i][1]) - int(column1[9][1]))/100
                    if column1[i][5] + column1[i+1][5] == 3:
                        OffFrame = i-7
                        OffTime = (int(column1[i][1]) - int(column1[9][1]))/100
                        #while ((int(column1[i+count][1])-int(column1[i][1]))<=3000):
                        #    count += 1
            elif ExperimentID4[iii-1] == '3':
                if OffLine1 ==0:  # old Offline data
                    OnFrame = column1[1][3]
                    OffFrame = column1[1][4]
                    OnTime = column1[int(OnFrame)+9][1]
                    OffTime = column1[int(OffFrame)+9][1]
                else:  # new Offline data
                    if int(ExperimentID[-2:])%2 == 1:
                        OnTime = 300
                        OffTime = 600
                        OnFrame = 7858
                        OffFrame = 15723
                    else:
                        OnTime = 309
                        OffTime = 600
                        OnFrame = 8112
                        OffFrame = 15979



            total = len(column1)
            if ExperimentID4[iii-1] != '3':
                temp1, FishID = column1[4][4].split(":")
            else:
                FishID = 'OffLine'
            OriginalFrameRate = float(OriginalFrameRate1)
            #sbox = (df['Unnamed: 2'][4], df['Unnamed: 3'][4])
            sbox = (int(column1[4][2]), int(column1[4][3]))

            boxes.append(sbox)
            boxesA = boxes[0][0]
            #sbox = (df['Unnamed: 2'][5], df['Unnamed: 3'][5])
            sbox = (int(column1[5][2]), int(column1[5][3]))
            boxes.append(sbox)
            boxesB = boxes[1][0]
            for i in range(9,int(total),1):
                if ExperimentID4[iii-1] != '3':
                    array1.append(float((int(column1[i][1])-int(column1[9][1]))/100))
                else:
                    array1.append(float(column1[i][1]))
                #x = df['Unnamed: 2'][i]
                #y = df['Unnamed: 3'][i]
                x = int(column1[i][2])
                y = int(column1[i][3])
                if (x) > (2 * (int((boxes[1][0] - boxes[0][0])/3)) + boxes[0][0]):
                    tempRight = tempRight + 1
                else:
                    if (x) <= (int((boxes[1][0] - boxes[0][0])/3) + boxes[0][0]):
                        tempLeft = tempLeft + 1
                    else:
                        tempCenter = tempCenter + 1
                if (y) > (int((boxes[1][1] - boxes[0][1])/3)*2 + boxes[0][1]):
                    tempDown = tempDown + 1
                    if i <= OnFrame:
                        tempDownBe = tempDownBe + 1
                    elif i <= OffFrame:
                        tempDownAf = tempDownAf + 1
                else:
                    tempUp = tempUp + 1
                x5.append(float(x-int(boxes[0][0]))/float((int(boxes[1][0])-int(boxes[0][0]))))
                y5.append(float(y-int(boxes[0][1]))/float((int(boxes[1][1])-int(boxes[0][1]))))
                if i==9:
                    distance1 = 0
                    distance0 = 0
                    z5.append(distance1)
                    z55.append(distance0)
                    z555.append(0)
                    x0 = x
                    y0 = y
                else:
                    if i>9:
                        distance0 = distance1
                    distance1 = math.sqrt((x-x0)**2 + ((y-y0)**2))
                    if round(distance1) > tempD2:
                        tempD1 = tempD1 + 1
                        tempD.append([tempD1, i, int(distance1)])
                    z5.append(distance1)
                    z55.append(distance1-distance0)
                    if x > x0:
                        if x - 3 > x0:
                            z555.append(1)
                            temp555a=1
                        else:
                            temp555a = temp555b
                            z555.append(temp555a)
                    elif x0 > x:
                        if x0 - 3 > x:
                            z555.append(-1)
                            temp555a=-1
                        else:
                            temp555a = temp555b
                            z555.append(temp555a)
                    else:
                        temp555a = temp555b
                        z555.append(temp555a)
                    if temp555a != temp555b:
                        temp555c.append(1)
                    else:
                        temp555c.append(0)
                    temp555b = temp555a

                    x0 = x
                    y0 = y
            count1 = 1
            if int(len(array1)/1000) == 17:
                array2=[]
                for m in range(0,len(array1),1):
                    array2.append(float(1/30.) * m)
                array1=array2
            if int(len(array1)/1000) == 18:
                array2=[]
                for m in range(0,len(array1),1):
                    array2.append(float(1/30.) * m)
                array1=array2

            if int(len(array1)/1000) == 27:
                array2=[]
                for m in range(0,len(array1),1):
                    array2.append(float(1/OriginalFrameRate) * m)
                array1=array2

            if int(len(array1)/1000) == 26:
                array2=[]
                for m in range(0,len(array1),1):
                    array2.append(float(1/OriginalFrameRate) * m)
                array1=array2

            for dd in xrange(0,np.size(z5), int(OriginalFrameRate)):
                if WhichTank == 2:
                    if int(FirstOrBoth) == 1:
                        pass
                    else:
                        distance_in_seconds.append (np.sum(z5[dd:dd+int(OriginalFrameRate)]))
                        speedtime.append(count1)
                        single_speed.append (np.sum(z5[dd:dd+int(OriginalFrameRate)]))
                        count1 +=1
                else:
                    if int(FirstOrBoth) == 1:
                        if iii % 2 == 1:
                            distance_in_seconds.append (np.sum(z5[dd:dd+int(OriginalFrameRate)]))
                            speedtime.append(count1)
                            single_speed.append (np.sum(z5[dd:dd+int(OriginalFrameRate)]))
                            count1 +=1
                        else:
                            pass
                    else:
                        distance_in_seconds.append (np.sum(z5[dd:dd+int(OriginalFrameRate)]))
                        speedtime.append(count1)
                        single_speed.append (np.sum(z5[dd:dd+int(OriginalFrameRate)]))
                        count1 +=1
            OnTimeSec = int(OnFrame / OriginalFrameRate)
            for dd in xrange(0,np.size(single_speed),1):
                if WhichTank == 2:
                    if int(FirstOrBoth) == 1:
                        pass
                    else:
                        if dd <= OnTimeSec:
                            speed_before.append(single_speed[dd])
                        else:
                            speed_after.append(single_speed[dd])
                else:
                    if int(FirstOrBoth) == 1:
                        if iii % 2 == 1:
                            if dd <= OnTimeSec:
                                speed_before.append(single_speed[dd])
                            else:
                                speed_after.append(single_speed[dd])
                    else:
                        if dd <= OnTimeSec:
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
                    print 'ccc'

                    Plotting1()
                    iii +=1
            else:
                JumpFrame()
                Behavior1()
                if iii == iiii:
                    distance_in_seconds = (np.asarray(distance_in_seconds) * float(200) / float(int(boxes[1][0])-int(boxes[0][0])))/10
                    speed_before = (np.asarray(speed_before) * float(200) / float(int(boxes[1][0])-int(boxes[0][0])))/10
                    speed_after = (np.asarray(speed_after) * float(200) / float(int(boxes[1][0])-int(boxes[0][0])))/10
                Plotting1(ExperimentID2,iii,ExperimentID4,OffTime,iiii,array1,OriginalFrameRate)
                #Plotting3()
                if iii == iiii:
                    Plotting1(ExperimentID2,iii,ExperimentID4,OffTime,iiii,array1,OriginalFrameRate)
                    Plotting2
                iii +=1
            df=[];
            boxes=[]
            sbox=[]
            x5=[]
            y5=[]
            z5=[]
            z55=[]
            z555=[]
            temp555c = []
            array1=[]
            tempRight=0
            tempCenter=0
            tempLeft=0
            tempUp=0
            tempDown=0
            tempDownBe=0
            tempDownAf=0
            x0 = 0
            y0 = 0
            distance1 = 0
            tempD1 = 0
            tempD = []
            tempZero =[]
            tempZero1 = 0
        else:
            if iii == iiii:
                distance_in_seconds = (np.asarray(distance_in_seconds) * float(200) / float(int(boxesB)-int(boxesA)))/10
                speed_before = (np.asarray(speed_before) * float(200) / float(int(boxesB)-int(boxesA)))/10
                speed_after = (np.asarray(speed_after) * float(200) / float(int(boxesB)-int(boxesA)))/10
                Plotting1(ExperimentID2,iii,ExperimentID4,OffTime,iiii,array1,OriginalFrameRate)
                Plotting3()
            iii+=1
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
            img1a = cv2.morphologyEx(img1a,cv2.MORPH_OPEN,kernel)   # cv2.MORPH_OPEN = erode and then dilate while cv2.MORPH_CLOSE = dilate and then erode
            img1a = cv2.morphologyEx(img1a,cv2.MORPH_OPEN,kernel)   # run it again
            #img1a = cv2.dilate(img1a,kernel,iterations=1)
            ret, img1b = cv2.threshold(img1a,65,255,0)
            # change the above threshold between 30 and 90 if you encounter the tail issue again, the higher the value, the more strict.
            # Please note that if the threshold is too high, the tracking codes may lose the fish.
            # When the tracking codes lose the fish, it gives x=0 and y=0 and you will see many lines connecting to the top left corner.
            # Check the output Excel file name "Tank2_OffLine_Zero List_n=1_frame" This number should be zero.
            # If the number is not zero, reduce the treshold and do it again

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
                z55.append(0)
                z555.append(0)
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
                    z5.append(0)
                    z55.append(0)
                    z555.append(0)
                    x0 = x
                    y0 = y
                else:
                    distance0=distance1
                    data1.append([x+w/2+boxes[0][0], y+h/2+boxes[0][1], x, y, w, h, 0, round(distance1),int(M['m00'])])
                    x5.append(float(x+w/2)/float((int(boxes[1][0])-int(boxes[0][0]))))
                    y5.append(float(y+h/2)/float((int(boxes[1][1])-int(boxes[0][1]))))
                    z5.append(distance1)
                    z55.append(distance1-distance0)
                    if x > x0:
                        if x - 3 > x0:
                            z555.append(1)
                            temp555a=1
                        else:
                            temp555a = temp555b
                            z555.append(temp555a)
                    elif x0 > x:
                        if x0 - 3 > x:
                            z555.append(-1)
                            temp555a=-1
                        else:
                            temp555a = temp555b
                            z555.append(temp555a)
                    else:
                        temp555a = temp555b
                        z555.append(temp555a)
                    if temp555a != temp555b:
                        temp555c.append(1)
                    else:
                        temp555c.append(0)
                    temp555b = temp555a
                    x0 = x
                    y0 = y

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
            book.save(path + "Tank1_OffLine_Raw.xls")
        else:
            book.save(path + ExperimentID + "Tank1_OffLine_Raw.xls")
    else:
        if int(OnOff) == 2:
            book.save(path + "Tank2_OffLine_Raw.xls")
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
                    book.save(path + "Tank1_OffLine_Zero List_n=" + str(tempZero1) + "_frame.xls")
                else:
                    book.save(path + ExperimentID + "Tank1_OffLine_Zero List_n=" + str(tempZero1) + "_frame.xls")
            else:
                book.save(path + ExperimentID + "Tank1_OnLine_Zero List_n=" + str(tempZero1) + "_frame.xls")
        else:
            if ExcelStatus ==0:
                if int(OnOff) == 2:
                    book.save(path + "Tank2_OffLine_Zero List_n=" + str(tempZero1) + "_frame.xls")
                else:
                    book.save(path + ExperimentID + "Tank2_OffLine_Zero List_n=" + str(tempZero1) + "_frame.xls")
            else:
                book.save(path + ExperimentID + "Tank2_OnLine_Zero List_n=" + str(tempZero1) + "_frame.xls")
    else:
        if WhichTank ==1:
            if ExcelStatus ==0:
                if int(OnOff) == 2:
                    book.save(path + "Tank1_OffLine_Zero List_n=" + str(tempZero1) + "_frames.xls")
                else:
                    book.save(path + ExperimentID + "Tank1_OffLine_Zero List_n=" + str(tempZero1) + "_frames.xls")
            else:
                book.save(path + ExperimentID + "Tank1_OnLine_Zero List_n=" + str(tempZero1) + "_frames.xls")
        else:
            if ExcelStatus ==0:
                if int(OnOff) == 2:
                    book.save(path + "Tank2_OffLine_Zero List_n=" + str(tempZero1) + "_frames.xls")
                else:
                    book.save(path + ExperimentID + "Tank2_OffLine_Zero List_n=" + str(tempZero1) + "_frames.xls")
            else:
                book.save(path + ExperimentID + "Tank2_OnLine_Zero List_n=" + str(tempZero1) + "_frames.xls")
    print '========================================'
    JumpFrame()
    Behavior1()
    Plotting1()