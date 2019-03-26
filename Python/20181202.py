import numpy as np
def cycle(arr):
    n = np.size(arr, 0)
    if n<=1: return [] if n<=0 else [arr[0,0]]
    lst1=[x for x in arr[0,:]]+[x for x in arr[1:n,n-1]]
    lst2=[x for x in arr[n-1,n-2::-1]]+[x for x in arr[n-2:0:-1,0]]
    return lst1+lst2+cycle(arr[1:n-1,1:n-1])
n=5
arr=np.reshape([i+1 for i in range(pow(n,2))],[n,n]).astype(np.int)
print(arr)
print("")
print(np.reshape(cycle(arr),[n,n]))