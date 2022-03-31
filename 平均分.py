import pandas as pd

#将DataFrame转换为list类型，将论文有价值的信息存储在list中
def praselist(z):
    n= []
    for index, row in z.iterrows():
        l = []        
        for index1, row1 in row.items():
            l.append(row1)
        if(l[0]<10):
            n.append([index,"0"+str(l[0]),'否',(l[1]+l[7]+l[13])/3,(l[2]+l[8]+l[14])/3,(l[3]+l[9]+l[15])/3,(l[4]+l[10]+l[16])/3,min(l[5],l[11],l[17]),(l[5]+l[11]+l[17])/3])
        else:
            n.append([index,str(l[0]),'否',(l[1]+l[7]+l[13])/3,(l[2]+l[8]+l[14])/3,(l[3]+l[9]+l[15])/3,(l[4]+l[10]+l[16])/3,min(l[5],l[11],l[17]),(l[5]+l[11]+l[17])/3])
    return n

#读取附件1.xlsx
data = pd.read_excel(r"附件1.xlsx",index_col=0,usecols=range(20))

#存储问题论文
m = []

#转换data为list类型,并存储所有论文
z = praselist(data)

#论文按照学科门类为key，论文为value分类存储在字典中
dic = {}
for i in z:
    if i[1] not in dic:
        dic[i[1]] = []
    dic[i[1]].append(i)

#论文分类后排序
for i in dic:
    dic[i] = sorted(dic[i], key=lambda paper: paper[7],reverse=True)
    #除去各论文信息中的无用值
    for d in range(int(len(dic[i]))):
        dic[i][d].pop(-2)
    #存储问题论文
    for d in range(int(len(dic[i])*0.05)):
        m.append(dic[i].pop())

#清空z,再次添加所有论文
z=[]
for i in range(len(m)):
    #修改各个淘汰论文是否淘汰的list对应值
    m[i][2] = "是"
    z.append(m[i])
for i in dic:
    for d in range(len(dic[i])):
        z.append(dic[i][d])

#论文按照bm排序
z = sorted(z, key=lambda paper: paper[0])

#list转化为DataFrame
df = pd.DataFrame(z,columns=['bm', 'Tag', '是否淘汰','选题与综述平均分','创新性及论文价值平均分','科研能力与基础知识平均分','论文规范性平均分','论文总分平均分']) 

#设置索引为bm
df = df.set_index('bm')

#填入附件2.xlsx
df.to_excel("附件2.xlsx")

#清空所有论文，重新输入
z = []
for i in range(len(m)):
    #修改各个淘汰论文是否淘汰的list对应值
    m[i][2] = "问题论文"
    z.append(m[i])

#填写附件2.xlsx
df.to_excel("附件2.xlsx")

