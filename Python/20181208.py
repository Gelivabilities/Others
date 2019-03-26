import wave
bpm=200
start=3.09
end=109.89
beats_per_measure=4
'''
def downsampleWav(src, dst, inrate=44100, outrate=4096, inchannels=2, outchannels=1):
    import os, wave, audioop
    if not os.path.exists(src):
        print ('Source not found!')
        return False

    if not os.path.exists(os.path.dirname(dst)):
        os.makedirs(os.path.dirname(dst))

    try:
        s_read = wave.open(src, 'r')
        s_write = wave.open(dst, 'w')
    except:
        print ('Failed to open files!')
        return False

    n_frames = s_read.getnframes()
    data = s_read.readframes(n_frames)

    try:
        converted = audioop.ratecv(data, 2, inchannels, inrate, outrate, None)
        if outchannels == 1:
            converted = audioop.tomono(converted[0], 2, 1, 0)
    except:
        print ('Failed to downsample wav')
        return False

    try:
        s_write.setparams((outchannels, 2, outrate, 0, 'NONE', 'Uncompressed'))
        s_write.writeframes(converted)
    except:
        print ('Failed to write wav')
        return False

    try:
        s_read.close()
        s_write.close()
    except:
        print ('Failed to close wav files')
        return False

    return True
'''

dst='C:/users/administrator/desktop/'
src=dst+'rotter tarmination.wav'

wav=wave.open(src,'rb')
frame_rate=wav.getparams()[2]
print(frame_rate)
temp=str(wav.readframes(5))
print(temp)

frames=str(wav.readframes(int(frame_rate*end))).split('\'\\')[1].split('\\')
print(len(frames)/frame_rate)
valid_frames=frames[int(start*frame_rate):]
print(len(valid_frames))
#for i in range(100):
#    print(valid_frames[i],end='\n' if (i+1)%10==0 else ' ')

#s = file.read(4)
#print(s)
#s = file.read(44)
#print(s)

#downsampleWav(src,dst,44100,int(4096/beats_per_measure*60/bpm))
