import wave
import numpy as np
waves=[[1,2,3,4],[5,6,7,8]]
notes=[[1,0,2,0],[1,1,1,2]]
file_notes = open('c:/users/administrator/desktop/神经网络作谱/notes.txt','w');
file_waves=open('c:/users/administrator/desktop/神经网络作谱/waves.txt','w');
for i in range(np.size(notes,0)):
    file_notes.write(' '.join([str(x) for x in notes[i]])+'\n')
    file_waves.write(' '.join([str(x) for x in waves[i]])+'\n')
file_notes.close();
file_waves.close();

'''
def config_read(tja_config_data,files):
    config_dict = {}
    for txt in tja_config_data:
        if txt.replace(' ','') == '': continue #跳过空行
        if txt.index(':') > 0:
            [key, value] = txt.split(':')
            config_dict[key.lower()] = value.lower()
    wav_name=config_dict['wave'].split('.')[0]+'.wav'
    if wav_name not in files:#没这个音频文件则报错
        raise FileError('Cannot find file:',wav_name)
    #判断是否魔王，没有写就默认是
    try: is_oni=config_dict['course'] == 'oni'
    except: is_oni=True
    return float(config_dict['bpm']), -float(config_dict['offset']), \
           is_oni,wav_name

tja_config_data=['TITLE:又埼玉2000','LEVEL:10','BPM:240','WAVE:madarot.ogg',
                 'OFFSET:-0.879','SONGVOL:100','  ','SEVOL:100','BALLOON:100,6',
                 'SCOREINIT:370','SCOREDIFF:90','','DEMOSTART:0.7']
files=['a','madarot.wav','']
print(config_read(tja_config_data,files))
'''
'''
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
    beat=np.zeros(cuts)
    beat_length=(end-start)*framerate
    for cut in range(cuts):
        beat[cut] = int(round(np.mean(abs(
            wave_data[int(cut/cuts*beat_length):int((cut+1)/cuts*beat_length)]))))#按声强取平均值
    beat = beat/np.max(abs(beat)) # 采样的音频归一化
    return beat
'''
'''
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
    beat=np.zeros(cuts)
    beat_length=int((end-start))*framerate
    for cut in range(cuts):
        beat[cut] = int(round(np.mean(abs(
            wave_data[int(cut/cuts*beat_length):int((cut+1)/cuts*beat_length)]))))#按声强取平均值
    beat = beat/np.max(abs(beat)) # 采样的音频归一化
    return [beat.tolist()]
print(get_wav('C:/users/administrator/desktop/rotter tarmination.wav',0,5))
'''
'''
import numpy as np
def change_txt(txt,measure):
    if len(txt)==0:return np.zeros(4*measure).astype(int).tolist()
    if txt.count('0')+txt.count('1')+txt.count('2')+txt.count('3')+txt.count('4')!=len(txt):
        raise TextError#谱面只能是：空，红，蓝，大红，大蓝
    step=int(4*measure/len(txt))
    new_txt=np.zeros(4*measure)
    count=0
    dict=[0,1,2,1,2]
    for i in [x for x in range(4*measure)][0:4*measure:step]:
        new_txt[i]=dict[int(txt[count])]#大红蓝变小红蓝，其他正常映射
        count+=1
    return new_txt.reshape([measure,4]).astype(int).tolist()

txt='101324'
measure=3
print(change_txt(txt,measure))
'''
'''
notes=np.zeros([num_of_beats,note_density])
for i in range(num_of_beats):
    for j in range(note_density):
        notes[i,j]=random.randint(0,2)#现在随便生成些鼓点吧
'''
'''
import random
import numpy as np

Y=np.zeros([356,4])
for i in range(356):
    for j in range(4):
        Y[i,j]=random.randint(0,2)#现在随便生成些鼓点吧
print(Y)
'''

'''
#乐曲属性
bpm=200
start=3.09
end=109.89
beat_per_measure=4
#路径
dst='C:/users/administrator/desktop/'
src=dst+'rotter tarmination.wav'
f = wave.open(src, "rb")
params = f.getparams()
nchannels, sampwidth, framerate, nframes = params[:4]
str_data = f.readframes(nframes)
f.close()
wave_data = np.fromstring(str_data, dtype=np.short)
wave_data.shape = -1, 2
wave_data = wave_data.T[0]
time = np.arange(0, nframes) * (1.0 / framerate)

#分节拍，截取每拍，采样成k段，便于作为神经网络的输入
k=64
max_beat=16#最密16分音符
p=int(max_beat/beat_per_measure)#每拍多少个音符位
num_of_beats=8#round((end-start)*bpm/60)
beats=np.zeros([num_of_beats,k])
for beat in range(num_of_beats):
    beat_wave=wave_data[round(framerate*(start+beat*60/bpm)):#按时间定位节拍
                        round(framerate*(start+(beat+1)*60/bpm))]
    beat_length=len(beat_wave)#每拍的帧数
    for cut in range(k):
        beats[beat,cut]=int(round(np.mean(abs(beat_wave[int(cut/k*beat_length):
                               int((cut+1)/k*beat_length)]))))#按声强取平均值
beats=beats/np.max(abs(beats))#采样的音频归一化


#谱面：现在随便生成些鼓点吧（0,1,2）
notes=np.array([[1,0,0,1],
                [2,0,1,0],
                [1,0,0,1],
                [2,0,1,0],
                [1,0,0,1],
                [0,0,1,0],
                [1,1,1,2],
                [1,1,2,2]])
notes=notes/2#谱面数值归一化
#训练作谱器
Cmp=composer.train(beats,notes,k,p,16,2,10000)
new_notes=Cmp.create_text(beats)
print(new_notes)
for i in range(np.size(new_notes,0)):
    for j in range(np.size(new_notes,1)):
        new_notes[i,j]=(0 if new_notes[i,j]<1/3
                        else (2 if new_notes[i,j]>=2/3 else 1))
    print(new_notes[i,:],end=' ' if (i+1)%4!=0 else '\n')
    if (i+1)%4==0:
        print(2*notes[i-3],2*notes[i-2],2*notes[i-1],2*notes[i])
        print('')
'''