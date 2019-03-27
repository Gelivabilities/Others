#整个程序分为以下几个部分：
#1. 爬取数据集，设定读取谱面文件的规则，并找出对应音频文件，对之进行采样
#2. 用爬取的数据集训练做谱器模型
#3. 用作谱器对新的采样音频进行作谱

#基本工具
import wave
import numpy as np
import os


#用相似法作谱，相似度用距离表示
# 取众数时由于不同颜色比例有差异，因此加个权重p会比较好
def composer_KNN(training_wave,training_notes,compose_wave,training_message,p,k=4):
    n=np.size(training_wave,0)#训练样本个数
    d=[]
    for i in range(n):
        d.append(np.sqrt(np.sum(np.power(training_wave[i]-compose_wave,2))))
    # 找到最相似的k个声波对应的谱面，每个16分音符取众值。特别地，当k=1时，直接取最相似谱面
    # 按相似度排序，得到排序后的索引
    sorted_index=np.array(d).argsort()
    print('作谱来源：')
    origin=[training_message[x].tolist() for x in sorted_index[0:k]]
    notes=np.array([training_notes[x,:] for x in sorted_index[0:k]],dtype='int64').T
    # 将前k个最相似谱面取众数，作为要返回的谱面
    mode_notes=[]
    for i in range(16):
        v=np.unique(notes[i])#0、1、2取了哪些值
        bin_count=np.bincount(notes[i]).astype(float)
        m=np.array([0,0,0])
        for value in v: m[value]=bin_count[value]
        count_note=m*p
        mode_notes.append(np.argmax(count_note))
    for i in range(k):
        print(origin[i][0],origin[i][1],notes.T[i])
    return mode_notes

class FileError(ValueError):
    pass

class TextError(ValueError):
    pass

#读取配置信息，合法返回相关信息，没文件则报错
def config_read(tja_config_data,files):
    config_dict = {}
    for txt in tja_config_data:
        if txt.replace(' ','').replace('\n','') == '': continue #跳过空行
        if txt.index(':') > 0:
            [key, value] = txt.split(':')
            config_dict[key.lower()] = value.lower()
    wav_name=config_dict['wave'].split('.')[0]+'.wav'
    if wav_name not in [x.lower() for x in files]:#没这个音频文件则报错
        raise FileError('Cannot find file:',wav_name)
    #判断是否魔王，没有写就默认是
    try: is_oni=config_dict['course'].replace('\n','') == 'oni' or config_dict['course'].replace('\n','') == '3'
    except Exception as e: is_oni=True
    return float(config_dict['bpm']), abs(float(config_dict['offset'])), \
           is_oni,wav_name

def change_txt(txt,measure):
    if len(txt)==0:return np.zeros(4*measure).astype(int).tolist()
    if txt.count('0')+txt.count('1')+txt.count('2')+txt.count('3')+txt.count('4')!=len(txt):
        raise TextError('遇到黄条、气球或不合法音符')#谱面只能是：空，红，蓝，大红，大蓝
    step=int(4*measure/len(txt))
    new_txt=np.zeros(4*measure)
    count=0
    dict=[0,1,2,1,2]
    for i in [x for x in range(4*measure)][0:4*measure:step]:
        new_txt[i]=dict[int(txt[count])]#大红蓝变小红蓝，其他正常映射
        count+=1
    return new_txt.reshape([measure,4]).astype(int).tolist()

def song_read(bpm,start_time,tja_text_data,wav_address):#读取谱面部分，返回转换成的小节和结束时间
    measure=4
    flag_measure=True
    txt_line=''
    sect_start_time=start_time
    sect_time_length=60/bpm#暂定一拍长度
    sect_end_time=start_time+measure*sect_time_length#暂定小节结束时间
    #要导出小节信息（wav名+第几小节）
    sect_count=0
    note=[]
    sect_wave=[]
    for txt in tja_text_data:
        try:
            if txt.find('#')==0:#1. 关键字类型
                keyword=txt.split('#')[1].lower()
                if txt.find(' ')>0:#关键字后面带数值（BPM改变，长度改变等）
                    [key,value]=keyword.split(' ')
                    #分情况处理：Python太垃圾了没有switch语句
                    if key=='branchstart':
                        raise TextError('存在分歧谱面，本曲不再读取')
                    if key=='bpmchange':
                        bpm=float(value)
                        sect_time_length = 60 / bpm
                        print('BPM变为：',bpm)
                    if key=='measure':
                        [value1,value2]=value.split('/')
                        if int(value2)!=4:#切成8份的不考虑，不然要分成不同的训练集来面对复杂的谱面作谱
                            print('分母不是4的小节长度')
                            flag_measure=False
                        else:
                            flag_measure=True
                        measure=int(4*int(value1)/int(value2))#该小节的节拍个数
                        print('Measure变为：',measure)
                        sect_end_time = sect_start_time + measure * sect_time_length  # 更新当前小节的结束时间
            else: #2. 谱面类型
                # 有的只有一个逗号，跳过，时间要加上整个小节的时长。小节的时长和bpm与节拍个数有关
                if txt.replace('\n','')==',':
                    print('空小节，时间：',sect_start_time,sect_end_time)
                    [sect_start_time, sect_end_time] = sect_time_change(sect_end_time,bpm,measure)
                    sect_count+=1
                    continue
                # 遇到分母不是4的小节，避免谱面因存在复杂情况而无法计算时间，直接报错
                if not flag_measure:
                    raise TextError('遇到分母不是4的小节，本曲不再读取')
                # 跳过空行，时间不用变
                if txt.replace(' ', '') .replace('\n','')== '': continue
                #看看一行里有没有逗号，没有的话直接报错，以此避免复杂情况
                if txt.find(',')==-1:
                    raise TextError('一小节中没有逗号，本曲不再读取')
                    #txt_line+=txt
                    #continue
                else:txt_line+=txt.split(',')[0]#去掉逗号剩下数字
                len_note=len(txt_line)#读取到的这一行音符有多少个
                if (4*measure)%len_note!=0:#不能变换成16分音符的情况，直接报错避免问题
                    print(txt_line, measure)
                    raise TextError('不能转换成16分音符，本曲不再读取')
                    #txt_line=''
                    #continue
                else:#通过了上面那么多关，基本可以确定是正常谱面信息（应该没有人那么无聊作死乱写吧）
                    #谱面、谱面对应的采样音频存入数组中
                    note_temp_4=change_txt(txt_line, measure)  # 一个line对应一个小节和measure个拍
                    note_temp=[]
                    for i in range(measure):
                        note_temp+=note_temp_4[i]
                    print(txt_line, note_temp, measure)

                    if not measure == 4:  # 跳过measure不是4的情况
                        print('Measure不是4，跳过。时间：', sect_start_time, sect_end_time)
                        [sect_start_time, sect_end_time] = sect_time_change(sect_end_time, bpm, measure)
                        txt_line=''
                        sect_count += 1
                        continue
                    #这时measure也是4了
                    print(note_temp,' 时间：', sect_start_time, sect_end_time)
                    note += [note_temp]
                    sect_wave+=get_wav(wav_address,sect_start_time,sect_end_time)
                    [sect_start_time,sect_end_time]=sect_time_change(sect_end_time,bpm,measure)
                    sect_count += 1
                    print('')
                    txt_line=''#清空缓存
        except Exception as e:
            print('Song_read exception:',e)
            break
    return sect_wave,note,[x+1 for x in range(sect_count)]

def sect_time_change(sect_end_time,bpm,beat=1):
    # 新的起始时间为上一小节结束时间
    sect_start_time=sect_end_time
    # 新的结束时间为新起始时间加小节长度
    sect_time_length=60/bpm
    sect_end_time=sect_start_time+beat*sect_time_length
    return sect_start_time,sect_end_time

def get_wav(wav_address,start,end,cuts=64):#获取小节对应音乐段，采样和转换音频
    f = wave.open(wav_address, "rb")
    params = f.getparams()
    nchannels, sampwidth, framerate, nframes = params[:4]
    str_data = f.readframes(nframes)
    f.close()
    wave_data = np.fromstring(str_data, dtype=np.short)
    wave_data.shape = -1, 2
    wave_data = wave_data.T[0]
    wave_data=wave_data[start*framerate:end*framerate]#截取
    #每拍采样成若干段
    beat=np.zeros(cuts).astype(int)
    beat_length=int((end-start)*framerate)
    for cut in range(cuts):
        beat[cut] = float(round(np.mean(abs(
            wave_data[int(cut/cuts*beat_length):int((cut+1)/cuts*beat_length)]))))#按声强取平均值
    '''#选择性归一化
    max=np.max(beat)
    min=np.min(beat)
    beat = (beat-min)/(max-min) # 采样的音频归一化
    '''
    return [beat.tolist()]

def tja_process(load_dir,save_dir):
    #数据处理：tja转音频data和谱面data
    # 读文件
    files = os.listdir(load_dir)
    waves = []  # 保存小节音乐
    notes = []  # 保存小节音符
    sect_message=[]# 保存小节信息
    for file in files:
        lst=file.split('.')
        if lst[len(lst)-1] != 'tja': continue  # 只读取tja文件
        tja_txt = open(load_dir + file, 'r').readlines()
        try:
            while True:#非魔王则继续，是魔王则break，开始读取谱面部分
                start_row=tja_txt.index('#START\n')#没#start的直接当不合法
                end_row=len(tja_txt)-1 if '#END\n' not in tja_txt else tja_txt.index('#END\n')
                #print(start_row,end_row)
                [bpm,start_time,is_oni,wav_name]=config_read(tja_txt[0:start_row-1],files)
                if is_oni:break
                else: tja_txt=tja_txt[end_row+1:]
            #到这里，音乐文件是存在的，定位的谱面也是魔王的
            #因为不存在文件或没魔王谱都会出错，进入except，跳过这个文件
            #可以开始读取谱面，返回小节的音乐与音符部分
            print(bpm, start_time, is_oni, wav_name)
            [wave_get,note_get,sect_message_get]=song_read(bpm,start_time,tja_txt[start_row+1:end_row-1],load_dir+wav_name)
            if len(wave_get)==len(note_get):#丢弃因为出错而导致谱面和音频小节数目不同的情况
                waves+=wave_get
                notes+=note_get
                name=wav_name.split('.')[0].replace(' ','')
                name_repeat=[]
                for i in range(len(sect_message_get)):
                    name_repeat.append(name)
                name_np=np.array(name_repeat)
                sect_message_np=np.array([str(x) for x in sect_message_get]).astype('U13')
                sect_message+=np.vstack([name_np,sect_message_np]).T.tolist()
            print(wav_name+'读入完成\n')
        except Exception as e:
            print('Main:exception',e)
            continue

    print(np.size(notes,0),np.size(notes,1))
    print(np.size(waves,0),np.size(waves,1))
    #每行为一个小节对应的音符
    file_notes = open(save_dir+'notes.txt','w')
    #每行为一个小节对应的声强
    file_waves=open(save_dir+'waves.txt','w')
    #每行为一个小节对应的歌曲信息
    file_message=open(save_dir+'message.txt','w')
    for i in range(np.size(notes,0)):
        file_notes.write(' '.join([str(x) for x in notes[i]])+'\n')
        file_waves.write(' '.join([str(x) for x in waves[i]])+'\n')
        file_message.write(' '.join([str(x) for x in sect_message[i]]) + '\n')
    file_notes.close()
    file_waves.close()
    file_message.close()

#这段代码用来生成训练集
load_dir='C:/Users/Administrator/Desktop/神经网络作谱/'
save_dir=load_dir
tja_process(load_dir,save_dir)


#这段代码用来作谱
#读取训练集数据
load_dir='C:/Users/Administrator/Desktop/神经网络作谱/'
save_dir=load_dir#保存位置
#读取音符数据
file_notes=np.loadtxt(load_dir+'notes.txt')
#读取训练集音频数据和信息
file_waves = np.loadtxt(load_dir+'waves.txt')
file_message_load=open(load_dir + 'message.txt', 'r').readlines()
num_of_message=len(file_message_load)
file_message=np.zeros([num_of_message,2]).astype('str')
for i in range(num_of_message):
    message_str=file_message_load[i].replace('\n', '').split(' ')
    file_message[i,:]=message_str
#读取测试音频，截取小节，转换格式
wave_address=load_dir+'wave.wav'
#作谱当然还要提供音乐相关逻辑信息(给出起点，节拍个数根据实际情况设定)
bpm=154
start=2.415
num_of_sect=65

#计算平均音量
wave_processed=get_wav(wave_address, start, start+num_of_sect*240/bpm)
wave_mean_vol=np.mean(np.abs(wave_processed))#当前乐曲的平均音量
training_mean_vol=np.mean(np.abs(file_waves))#训练集乐曲的平均音量
print('全曲平均音量：',wave_mean_vol,'训练集平均音量：',training_mean_vol)
vol_diff=wave_mean_vol-training_mean_vol
#开始作谱
all_notes=[]
for i in range(num_of_sect):
    sect_time=240/bpm
    sect_start_time=start+i*sect_time
    sect_end_time=start+(i+1)*sect_time
    #将节拍波形音频采样成标准化的数据
    sect_wave_processed=get_wav(wave_address, sect_start_time, sect_end_time)
    #音量取平均
    sect_wave_meaned=sect_wave_processed-vol_diff
    get_notes=composer_KNN(file_waves,file_notes,sect_wave_processed,file_message,[0.375,1,0.75],8)
    print('第'+str(i+1)+'小节：',end='')
    print(get_notes,end='')
    print('，时间：'+str(sect_start_time)+'到'+str(sect_end_time))
    all_notes.append(str(get_notes))

#保存作谱
save_notes = open(save_dir+'save.txt','w')
for i in range(len(all_notes)):
    save_notes.write(all_notes[i]+',\n')
save_notes.close()
