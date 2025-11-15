""" 
iter学习：
iter（）函数一般和next（）函数同时使用。用于将 iterable列表中的元素一个一个的取出来。

"""

l = [1,2,3,4]
ll =['a','b','c','d']
lll=['e','f','g','h']

outfile = './zcj.txt'  
with open(outfile, 'a') as f:
    print('-----------------iter-next-stu----------------',file =f)
    it = iter(l)
    print(next(it),end='##',file=f)
    print(next(it),end='##',file=f)
    print(next(it),end='##',file=f)
    print(next(it),end='##',file=f)
    print(file=f)