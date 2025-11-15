#filter+lambda学习

""" 
    filter函数用于过滤符合可迭代对象l（即filter函数的iterables参数）中
     符合lambda（即filter函数中func参数）条件的结果 
"""
l = [1,2,3,4]
ll =['a','b','c','d']
lll=['e','f','g','h']

outfile = './zcj.txt'

with open(outfile, 'a') as f:
    print('--------------------filter stu -------------------------',file =f)
    print(list(filter(lambda x:x%2==0,l)),file =f)