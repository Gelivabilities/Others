# -*- coding: utf-8 -*-
"""
Created on Fri Nov  1 13:43:12 2019

@author: Administrator
"""
import os
import numpy as np
import traceback
import matplotlib.pyplot as plt
import pandas as pd

def to_energy(file_msg):
    text=separate(file_msg)#每个不同的谱面
    title=[]
    BPM=-1
    level=[]
    measure=1
    energy=[]
    note_counts=[]
    total_ts=[]
    densities=[]
    for i in range(len(text)):
        flag=False
        
        for line in text[i][0]:#获取信息
            if ':' in line:[name,value]=line.split(':')
            else:continue
            if name=='TITLE':title.append(value)
            if name=='BPM':BPM=float(value)
            if name=='LEVEL':level.append(int(value))
            
        for line in text[i][0]:#仅限Oni
            if (line.lower() in('course:oni','course:3','course:4')):flag=True

        flag2=True
        for line in text[i][0]:#没写oni的也当oni（不然一堆没法统计）
            if ('course' in line.lower()): flag2=False

        if not flag and not flag2:continue
        
        count=0
        notes=[]#缓存要计算的音符值
        bpms=[]#缓存要计算的音符BPM
        measures=[]#缓存要计算的音符measure
        counts=[]#缓存要计算的音符当前小节count有多少个
        for line in text[i][1]:
            line=line.replace('#BPMCHANGE','#BPMCHANGE ')
            line=line.replace('#MEASURE','#MEASURE ')
            line=line.replace('  ',' ')
            line_split=line.split()
            if '#' in line and ' ' in line:
                name=line_split[0]
                value=line_split[1]
            #本问题只对BPM、measure和逗号前面的音符个数敏感。分歧谱面去掉
            if name=='#BRANCHSTART':return [],[],[],[]
            if name=='#BPMCHANGE':BPM=float(value)
            if name=='#MEASURE':
                [num,den]=[float(t) for t in value.split('/')]
                measure=num/den
            if '#' not in line:#是纯谱面数据
                line=line.replace(' ','')#去空格
                if ',' in line:#有逗号
                    if line==',':line='0,'
                    temp=line.index(',')
                    count+=temp
                    notes+=list(line[:temp])
                    bpms+=[BPM for i in range(temp)]
                    measures+=[measure for i in range(temp)]
                    counts+=[count for i in range(count)]
                    count=0#遇到逗号count归零
                else:
                    count+=len(line)
                    notes+=list(line)
                    bpms+=[BPM for i in range(len(line))]
                    measures+=[measure for i in range(len(line))]
                    #counts+=[count for i in range(len(line))]

        #头尾去除非1234的音符
        min_head_index=65535
        if '1' in notes:min_head_index=min([min_head_index,notes.index('1')])
        if '2' in notes:min_head_index=min([min_head_index,notes.index('2')])
        if '3' in notes:min_head_index=min([min_head_index,notes.index('3')])
        if '4' in notes:min_head_index=min([min_head_index,notes.index('4')])
        head=min_head_index
        tail=len(notes)-1
        while notes[tail] not in ['1','2','3','4']: tail-=1
        notes,bpms,measures,counts=notes[head:tail+1],bpms[head:tail+1],measures[head:tail+1],counts[head:tail+1]
        #开始计算能量
        energy_temp,note_count,i,dt,total_t=0,1,0,0,0
        while i<len(notes)-1:

            dt=240/bpms[i]*measures[i]/counts[i]
            j=i+1
            while j<len(notes) and notes[j] not in ['1','2','3','4']:
                dt+=240/bpms[j]*measures[j]/counts[j]
                j+=1
            energy_temp+=1/dt**2
            i=j
            note_count+=1
            total_t+=dt
        energy.append(energy_temp/note_count)
        note_counts.append(note_count)
        total_ts.append(total_t)
        densities.append((note_count-1)/total_t)
    print(title,level,note_counts,total_ts,densities,energy)
    return title,level,note_counts,total_ts,densities,energy
    
def separate(file_msg):#谱面和谱面信息分割开
    i,l,i_temp,part_count,l_temp=0,[],0,0,[]
    while True:
        while i<len(file_msg) and file_msg[i]!='#START' and file_msg[i]!='#END':i+=1
        if i>=len(file_msg):break
        if file_msg[i]=='#END':i+=1
        l_temp_part=[]
        for j in range(i_temp,i):l_temp_part.append(file_msg[j])
        l_temp.append(l_temp_part)
        part_count+=1
        if part_count%2==0: 
            l.append(l_temp)
            l_temp=[]
        i_temp=i
        i+=1
    return l

#读文件输出结果
file_dir='C:\\Users\\Administrator\\Desktop\\10星\\'
files = os.listdir(file_dir)
l=[]
for file in files:
    if not file.split('.')[1]=='tja':continue
    file_msg=open(file_dir+file,'r').readlines()
    for i in range(len(file_msg)):file_msg[i]=file_msg[i].strip()#去换行符
    try:
        s=list(to_energy(file_msg))
    except:
        print(traceback.format_exc())
        continue
    for i in range(len(s[2])):
        l.append([file,s[0][0],s[1][0],s[2][i],s[3][i],s[4][i],s[5][i]])
l_np=np.array(l)
l_np= l_np[l_np[:,6].astype(np.float32).argsort()]
l_np_stars=l_np[l_np[:,2].astype(np.int).argsort(),2].astype(np.int)    
s=pd.Series(l_np_stars)
s=s.groupby(by=s.values).count()
s.to_csv('C:\\Users\\Administrator\\Desktop\\stars.csv')
plt.plot(pd.DataFrame(l_np[:,[6]]))
df=pd.DataFrame(l_np)
df.to_csv('C:\\Users\\Administrator\\Desktop\\results.csv')
plt.figure()
plt.xlim([0,225])
plt.hist(pd.DataFrame(l_np[:,[6]]).astype(np.float),bins=20,alpha=0.5)