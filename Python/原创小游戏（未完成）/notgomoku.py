import numpy as np

arr=np.zeros([15,15])

def fill(arr):
    x,y=np.size(arr[:,0]),np.size(arr[0,:])
    v=np.random.randint(x+y-1)
    x0,y0=0 if v<=y-1 else v-y+1,v if v<=y-1 else y-1
    x1,y1=x-1-x0,y-1-y0
    for i in range(x):
        for j in range(y):
            if y1==y0:arr[i,j]=1 if j>y1 else 2 if j<y1 else -1
            else:
                if i-x1<(x1-x0)/(y1-y0)*(j-y1):arr[i,j]=1 
                else: 
                    arr[i,j]=2 if i-x1>(x1-x0)/(y1-y0)*(j-y1) else -1
    return

fill(arr)
print(arr,np.sum(arr==1)==np.sum(arr==2))
'''
x1=np.random.randint(11)
y1=np.random.randint(11)

while True:
    x2=np.random.randint(11)
    y2=np.random.randint(11)
    if not(x2==x1 and y2==y1):break

arr[x1,y1]=1
arr[x2,y2]=2
'''

'''
depth=0
def spread(arr,x1,y1,x2,y2,depth):
    max_x=np.size(arr[:,0])
    max_y=np.size(arr[0,:])
    #上下左右
    bool1=np.random.rand(4)+1
    bool2=np.random.rand(4)+1

    if y1-1<0 or not arr[x1,y1-1]==0:bool1[0]=0
    if y1+1>=max_y or not arr[x1,y1+1]==0:bool1[1]=0
    if x1-1<0 or not arr[x1-1,y1]==0:bool1[2]=0
    if x1+1>=max_x or not arr[x1+1,y1]==0:bool1[3]=0

    if y2-1<0 or not arr[x2,y2-1]==0:bool2[0]=0
    if y2+1>=max_y or not arr[x2,y2+1]==0:bool2[1]=0
    if x2-1<0 or not arr[x2-1,y2]==0:bool2[2]=0
    if x2+1>=max_x or not arr[x2+1,y2]==0:bool2[3]=0

    #扩散一格
    act_1=np.argmax(bool1)
    act_2=np.argmax(bool2)
    new_x1=x1-1 if act_1==2 else x1+1 if act_1==3 else x1
    new_y1=y1-1 if act_1==0 else y1+1 if act_1==1 else y1
    new_x2=x2-1 if act_2==2 else x2+1 if act_2==3 else x2
    new_y2=y2-1 if act_2==0 else y2+1 if act_2==1 else y2
    arr[new_x1,new_y1]=1
    arr[new_x2,new_y2]=2
    #print(act_1,act_2,new_x1,new_y1,new_x2,new_y2,np.sum(bool1),np.sum(bool2))
    #递归扩散
    if np.sum(bool1)==0 or np.sum(bool2)==0: return
    else: 
        depth+=1
        spread(arr,new_x1,new_y1,new_x2,new_y2,depth)
    
    return
    
spread(arr,x1,x2,y1,y2,depth)
'''