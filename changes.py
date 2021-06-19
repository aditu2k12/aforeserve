import numpy as np
import pandas as pd
import time

l=[]
t1= time.time()
for i in range(10):
  if i %2==0:
    l.append(i)
    
print(time.time()-t1)

l=[]
t1= time.time()

l =[i for i in range(10) if i%2==0]
print(time.time-t1)
    
