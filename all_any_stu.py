""" 
all和any学习：
     all和any都是用于判断元素是否为真。
     all判断的是iteratable list只要有一个为假，则all返回false,iteratable list为空或没有假值元素则返回True.

     any判断的是iteratable list中只要有一个为真值，则any返回True,否则返回False

"""

l=[1,2,3,4]
ll=['a','b',1,3,'c',0]

outfile = './zcj.txt'

with open(outfile,'a') as f:
    print('------------------any_all_stu-------------------------',file=f)
    print(l,file=f)
    print(ll,file=f)
    print(f'any函数返回l和ll:{any(l)},{any(ll)},all函数返回l和ll:{all(l)},{all(ll)}',file =f)