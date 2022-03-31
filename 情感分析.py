import pandas as pd
import jieba
import math

#情感分析主控函数
def analysecontrol(text):
  try:
    lists  = text.split('\n')
    al_senti = ['无','积极','消极','消极','中性','消极','积极','消极','积极','积极','积极',
                 '无','积极','积极','中性','积极','消极','积极','消极','积极','消极','积极',
                '无','中性','消极','中性','消极','积极','消极','消极','消极','消极','积极'
                ]
    for list in lists:
        if list  != '':
            sentiments = round(getscore(list),2)
            return sentiments
  except:
      return 0

#基于波森情感词典计算情感值
def getscore(text):
    df = pd.read_table(r"dictionary.txt", sep=" ", names=['key', 'score'])
    key = df['key'].values.tolist()
    score = df['score'].values.tolist()
    # jieba分词
    segs = jieba.lcut(text,cut_all = False) #返回list
    # 计算得分
    score_list = [score[key.index(x)] for x in segs if(x in key)]
    return sum(score_list)

#读取文件
def read_txt(filename):
    with open(filename,'r',encoding='utf-8')as f:
        txt = f.read()
    return txt

#写入文件
def write_data(filename,data):
    with open(filename,'a',encoding='utf-8')as f:
        f.write(data)

#将DataFrame转换为list类型，将论文有价值的信息存储在list中
def praselist(z):
    n= []
    for index, row in z.iterrows():
        l = []        
        for index1, row1 in row.items():
            l.append(row1)
        if(l[0]<10):            
            n.append([index,"0"+str(l[0]),analysecontrol(l[6]),analysecontrol(l[12]),analysecontrol(l[18])])
        else:
            n.append([index,str(l[0]),analysecontrol(l[6]),analysecontrol(l[12]),analysecontrol(l[18])])
    return n

#读取附件1.xlsx
data = pd.read_excel(r"附件1.xlsx",index_col=0,usecols=range(20))

#转换data为list类型,并存储所有论文
z = praselist(data)

#list转化为DataFrame
z = pd.DataFrame(z,columns=['1', '2', '3', '4', '5'])

#填写data.xlsx
z.to_excel("data.xlsx")
