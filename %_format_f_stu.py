""" l=[1,2,3,4]
ll=['a','b',1,3,'c',0]


outfile = './zcj.txt'

with open(outfile,'a') as f:
    print('------------------格式化符号：%_format_f_stu-------------------------',file=f)
    print(l,file=f)
    print(ll,file=f)
    print(f'f符号返回的any函数值:{any(l)},{any(ll)},这是f符号返回的all函数值:{all(l)},{all(ll)}',file =f)
    print('format函数格式化返回的any函数值：{},{},这是format函数格式化返回的all函数值{},{}'.format(any(l),any(ll),all(l),all(ll)),file=f)
    #print("%符号格式化返回的any函数值 %s,%符号格式化返回的all函数值 %s." % (str(any(l)),str(all(ll))),file=f)
 """


name = 'zcj'
age =18
print('my name is %s,my age is %d years old'%(name,age))