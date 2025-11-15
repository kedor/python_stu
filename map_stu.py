# map+lambda学习

""" 
     map函数用于将可迭代对象l（即map函数的iterables参数）中的元素
     交给lambda（即map函数中func参数）函数依次运行并返回结果 

"""
     
l = [1,2,3,4]
ll =['a','b','c','d']
lll=['e','f','g','h']

outfile = './zcj.txt'

with open(outfile, 'a') as f:
    print('--------------------map stu -------------------------',file =f)
    print(list(map(lambda x:pow(x,x),l)),file =f)