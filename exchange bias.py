# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import win32com.client
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d

origin = win32com.client.Dispatch("Origin.ApplicationSI")
rtfd = origin.rootfolder
fds = rtfd.folders
output=[]
for i in range(fds.count):
    if fds[i].name[-1]=='K':
        Temp = int(fds[i].name[0:-1])
        fds[i].activate
        for j in range(fds[i].pagebases.count) :
            print i,j
            if fds[i].pagebases[j].longname[0:2]=='Bo' :
                data_temp=origin.GetWorksheet(fds[i].pagebases[j].name)
                if len(data_temp[0])<=3 and len(data_temp)>10:
                    data=data_temp
                    data=list(data)
                    for k in range(len(data)):
                        data[k]=list(data[k])
                    data=np.array(data)[:,0:2]
                    interp1=interp1d(data[0:len(data)/2-1,1],data[0:len(data)/2-1,0])
                    interp2=interp1d(data[len(data)/2+1:,1],data[len(data)/2+1:,0])
                    zero_point1=interp1(0)
                    zero_point2=interp2(0)
                    #zero_point1=np.interp(0,data[0:len(data)/2-1,1],data[0:len(data)/2-1,0])
                    #zero_point2=np.interp(0,data[len(data)/2+1:,1],data[len(data)/2+1:,0])
                    average=(zero_point1+zero_point2)/2
                    length=abs(zero_point1-zero_point2)
                    fig = plt.figure(i)
                    plt.plot(data[0:len(data)/2-1,0],data[0:len(data)/2-1,1],'.')
                    plt.plot(data[len(data)/2+1:,0],data[len(data)/2+1:,1],'.')
                    plt.plot(zero_point1,0,'r.')
                    plt.plot(zero_point2,0,'r.')
                    plt.title(fds[i].name)
                    fig.savefig(fds[i].name+'.jpg')
        output.append([Temp,zero_point1,zero_point2,average,length])
output=np.array(output)
plt.figure()
plt.plot(output[:,0],output[:,3],'.')
outputfile=rtfd.folders.add
outputfile.name="Qi"
outputfile.activate
pagename=origin.createpage(2,"Qi","origin")
origin.putworksheet(pagename,output)

    
        
        
