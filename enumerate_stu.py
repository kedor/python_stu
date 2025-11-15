l = [1,2,3,4]
ll =['a','b','c','d']
lll=['e','f','g','h']

outfile = './zcj.txt'
with open (outfile ,'a') as f:
    print('-----------enumerate study-----------------',file = f)
    for i in enumerate(ll):
        print(i,file= f)
    print(file = f)
        