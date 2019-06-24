import numpy as np
import pygame
from pygame.locals import *
import os

def fill(arr):#生成大小相同区域
    x,y=np.size(arr[:,0]),np.size(arr[0,:])
    v=np.random.randint(x+y-1)
    x0,y0=0 if v<=y-1 else v-y+1,v if v<=y-1 else y-1
    x1,y1=x-1-x0,y-1-y0
    for i in range(x):
        for j in range(y):
            if y1==y0:arr[i,j]=1 if j>y1 else 2 if j<y1 else -1
            elif i-x1<(x1-x0)/(y1-y0)*(j-y1):arr[i,j]=1 
            else: arr[i,j]=2 if i-x1>(x1-x0)/(y1-y0)*(j-y1) else -1
    return
	
def judge_link(arr,x,y,n):#判断某点的棋是否连成n
    max_x,max_y=np.size(arr[:,0]),np.size(arr[0,:])
    #竖
    i,counter=-n+1,1
    while i<n-1:
        if x+i<0 or x+i+1>=max_x:
            i+=1
            continue
        if arr[x+i,y]==arr[x+i+1,y]:
            counter+=1
            if counter==n:return True
        else: counter=1
        i+=1
    #横
    i,counter=-n+1,1
    while i<n-1:
        if y+i<0 or y+i+1>=max_y:
            i+=1
            continue
        if arr[x,y+i]==arr[x,y+i+1]:
            counter+=1
            if counter==n:return True
        else: counter=1
        i+=1
    #捺
    i,counter=-n+1,1
    while i<n-1:
        if x+i<0 or x+i+1>=max_x:
            i+=1
            continue
        if y+i<0 or y+i+1>=max_y:
            i+=1
            continue
    
        if arr[x+i,y+i]==arr[x+i+1,y+i+1]:
            counter+=1
            if counter==n:return True
        else: counter=1
        i+=1
    #撇
    i,counter=-n+1,1
    while i<n-1:
        if x+i<0 or x+i+1>=max_x:
            i+=1
            continue
        if y-i-1<0 or y-i>=max_y:
            i+=1
            continue
        if arr[x+i,y-i]==arr[x+i+1,y-i-1]:
            counter+=1
            if counter==n:return True
        else: counter=1
        i+=1
        
    return False

def put_chess(arr,x,y,note):arr[x,y]=note

cur_dir=os.getcwd()
pygame.init()
screen=pygame.display.set_mode((900,500))
screen.fill((200,200,0))

pygame.display.set_caption("谁先连五谁输")
image = pygame.image.load(cur_dir+'\\background.png')
new_game_button=pygame.image.load(cur_dir+'\\black_chess.png')


ball_x=np.random.randint(900)
ball_y=np.random.randint(500)
speed_x=18
speed_y=16

while True:
    #screen.blit(image,(0,0))
    p=pygame.mouse.get_pos()
    print(p)
    #if p[1]>=27 and p[1]<=78 and p[0]>=728 and p[0]<=836:
        #screen.blit(new_game_button,(728,27))
    screen.fill((200,200,0))
    screen.blit(new_game_button,(ball_x,ball_y))
    ball_x+=speed_x
    ball_y+=speed_y
    if ball_x>900 or ball_x<0:speed_x=-speed_x
    if ball_y>500 or ball_y<0:speed_y=-speed_y
    
    pygame.display.flip()
    pygame.time.wait(16)
    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            pygame.quit()




'''
size_x=8
size_y=8
n=5
arr=np.zeros([size_x,size_y])
fill(arr)
print(arr)
conditions=[[],[]]
flag=True
while True:
    a=int(input('Please put a chess (x): '))
    b=int(input('Please put a chess (y): '))
    put_chess(arr,a,b,3 if flag else 4)
    print(arr)
    if judge_link(arr,a,b,n):
        print('Black' if flag else 'White'+' lose')
        break
    flag=not flag
'''