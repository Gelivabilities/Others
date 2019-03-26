import wave
import numpy as np
import tensorflow as tf
import os
#作谱器模型
class composer():
    def __init__(self,W1,W2,X,Y_,A,Y):
        self._W1,self._W2,self._X,self._Y_,\
        self._A,self._Y=[W1,W2,X,Y_,A,Y]
    #作谱函数
    def create_text(self,wave):
        with tf.Session() as sess:
            init_op = tf.global_variables_initializer()
            sess.run(init_op)
            return sess.run(self._Y, feed_dict={self._X: wave})
    #训练作谱器函数
    def train(beats_wav,notes,wav_size,note_density,hidden_layer=50,batch_size=8,steps=5000):
        w1 = tf.Variable(tf.random_normal([wav_size, hidden_layer], stddev=1, seed=1))
        w2 = tf.Variable(tf.random_normal([hidden_layer, note_density], stddev=1, seed=1))
        #x是输入样本，y_是真实值
        x = tf.placeholder(tf.float32, shape=(None, wav_size), name='x-input')
        y_ = tf.placeholder(tf.float32, shape=(None, note_density), name='y-input')
        #定义神经网络函数
        a = tf.sigmoid(tf.matmul(x, w1))
        y = tf.sigmoid(tf.matmul(a, w2))
        #目标函数和优化器
        cross_entropy = -tf.reduce_mean(y_ * tf.log(tf.clip_by_value(y, 1e-10, 1.0)))
        train_step = tf.train.AdamOptimizer(0.001).minimize(cross_entropy)
        with tf.Session() as sess:
            init_op = tf.global_variables_initializer()
            sess.run(init_op)
            for i in range(steps):
                start = (i * batch_size) % num_of_beats
                end = min(start + batch_size, num_of_beats)
                sess.run(train_step, feed_dict={
                    x: beats_wav[start:end], y_: notes[start:end]})
            return composer(w1,w2,x,y_,a,y)
#乐曲
class song():
    def __init__(self,wave,bpm,start,end,beat_per_measure,notes=[]):
        self._wave=wave
        self._bpm=bpm
        self._start=start
        self._end=end
        self._beat_per_measure=beat_per_measure
        self.notes=notes
class FileError(ValueError):
    pass

class TextError(ValueError):
    pass

#读文件
dir='C:/Users/Administrator/Desktop/神经网络作谱/'
files=os.listdir(dir)
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

def change_txt(txt,measure):#得到measure个拍的4个音符
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
    flag_branch=False
    txt_line=''
    sess_start_time=start_time
    sess_time_length=60/bpm
    sess_end_time=start_time+sess_time_length
    note=[]
    beat_wave=[]
    for txt in tja_text_data:
        try:
            if txt.find('#')==0:#1. 关键字类型
                keyword=txt.split('#')[1].lower()
                if txt.find(' ')>0:#关键字后面带数值（BPM改变，长度改变等）
                    [key,value]=keyword.split(' ')
                    if key=='branchstart':
                        flag_branch=True
                        raise TextError('存在分歧谱面，本曲不再读取')
                    if key=='bpmchange':bpm=float(value)
                    if key=='measure':
                        [value1,value2]=value.split('/')
                        if int(value2)!=4:
                            print('分母不是4的小节长度')
                            flag_measure=False
                        else:
                            flag_measure=True
                        measure=int(4*int(value1)/int(value2))
            else: #2. 谱面类型
                if txt_line=='' and txt.replace('\n','')==',':continue#有的只有一个逗号，跳过
                if flag_branch:continue#出现分歧，后面所有谱面都跳过
                if not flag_measure:continue#跳过分母不是4的所有小节
                if txt.replace(' ', '') .replace('\n','')== '': continue  # 跳过空行
                if txt.find(',')==-1:
                    txt_line+=txt
                    continue
                else:txt_line+=txt.split(',')[0]
                len_note=len(txt_line)
                if 4*measure%len_note!=0:#不能变换成16分音符的情况跳过，并清空缓存
                    txt_line=''
                    continue
                else:
                    #谱面、谱面对应的采样音频存入数组中
                    note+=change_txt(txt_line,measure)#一个line对应一个小节和measure个拍
                    for m in range(measure):#measure个拍对应measure个音符段
                        beat_wave+=get_wav(wav_address,sess_start_time,sess_end_time)
                        #新的起始时间为上一小节结束时间
                        sess_start_time=sess_end_time
                        #新的结束时间为新起始时间加小节长度
                        sess_time_length = 60 / bpm
                        sess_end_time = sess_start_time + sess_time_length
                    txt_line=''#清空缓存
        except Exception as e:
            print('Song_read exception:',e)
            continue
    return beat_wave,note

def get_wav(wav_address,start,end,cuts=64):#获取小节对应音乐段
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
    beat = beat/np.max(abs(beat)) # 采样的音频归一化
    return [beat.tolist()]

waves = []  # 保存小节音乐
notes = []  # 保存小节音符
for file in files:
    lst=file.split('.')
    if lst[len(lst)-1] != 'tja': continue  # 只读取tja文件
    tja_txt = open(dir + file, 'r').readlines()
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
        [wave_get,note_get]=song_read(bpm,start_time,tja_txt[start_row+1:end_row-1],dir+wav_name)
        if len(wave_get)==len(note_get):#丢弃因为出错而导致谱面和音频小节数目不同的情况
            waves+=wave_get
            notes+=note_get
        print(wav_name+'读入完成\n')
    except Exception as e:
        print('Main:exception',e)
        continue

print(np.size(notes,0),np.size(notes,1))
print(np.size(waves,0),np.size(waves,1))

file_notes = open('c:/users/administrator/desktop/神经网络作谱/notes.txt','w')
file_waves=open('c:/users/administrator/desktop/神经网络作谱/waves.txt','w')
for i in range(np.size(notes,0)):
    file_notes.write(' '.join([str(x) for x in notes[i]])+'\n')
    file_waves.write(' '.join([str(x) for x in waves[i]])+'\n')
file_notes.close()
file_waves.close()
