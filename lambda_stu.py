#近日学习函数式编程中的map、filter、lambda内容。
# 首先学习lambda表达式
l = [1,2,3,4]
ll =['a','b','c','d']
lll=['e','f','g','h']

outfile = './zcj.txt'
with open(outfile,'a') as f:
    print('-----------------lambda study-----------------',file = f)
    print(tuple(map(lambda x:x**2,l)),file = f)