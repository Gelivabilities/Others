import numpy as np
'''
arr=np.array([2,3,1,0,5,4,7,8,6])
#目标是输出2和3，因为2和3有重复数字
def find(arr):
    for i in range(np.size(arr)):
        while arr[i]!=i:
            if arr[arr[i]]==arr[i]:
                return arr[i]
            else:
                t=arr[i]
                arr[i]=arr[t]
                arr[t]=t
    return -1
'''
'''
arr=np.array([2,3,5,4,3,2,6,7])
def find(arr):
    n=np.size(arr)
    hash_table=np.zeros(n)
    for i in range(n):
        t=arr[i]
        if hash_table[t]==1:
            return t
        else:
            hash_table[t] += 1
    return -1
'''
arr=np.array([[ 1, 2, 3, 5, 6, 7, 9,10],
              [ 2, 3, 5, 6, 7, 9,11,13],
              [ 4, 5, 6, 8, 9,12,14,16],
              [ 6, 8, 9,10,12,14,16,18],
              [ 7, 9,10,12,13,15,18,20],
              [ 9,11,12,14,15,16,18,21],
              [11,13,15,16,17,18,20,22],
              [12,14,16,18,19,21,24,25]])
def find(arr,num):#m行n列，复杂度O(m+n)
    #从右上角开始定位
    row=0
    col=np.size(arr[0,:])-1
    while arr[row][col]!=num:
        if arr[row][col]>num:
            if col == 0:
                return False
            col-=1
        else:
            if row == np.size(arr[0, :]) - 1:
                return False
            row+=1
    return True
'''
#试下能不能二分
def quick_find(arr,num):#复杂度O(logm*logn)
    # 从右上角开始定位
    row = 0
    col = np.size(arr[0, :]) - 1
    range_col=[0,col]
    range_row=[0,np.size(arr[:,0])-1]
    UDRL=np.array([False,False,False,False])
    flag=False
    while arr[row][col] != num:
        # 给指针定位
        if arr[row][col] > num:#指针行往上，列往左

        else:#指针行往下，列往右

    return True
'''
for i in range(25):
    print(i+1,quick_find(arr,i+1))
print(np.unique(arr))