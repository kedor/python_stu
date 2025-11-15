l = [1,2,3,4]
ll =['a','b','c','d']
lll=['e','f','g','h']

outfile = './zcj.txt'
with open(outfile,'a') as f:
    for i in zip(l,ll):
        print(i,end='\n',file = f)
        
        
