import jieba
import jieba.analyse
import xlwt,xlrd
from xlutils.copy import copy
import os


search_key_words_path = "F:/淘宝关键字/宝贝无线APP搜索关键词表-530723471382-20180803.csv"
business_staf_platform = "F:/淘宝关键字/【生意参谋平台】2018-08-04_2018-08-04防水箱包－热搜搜索词－所有终端.xls"
save_result_file = 'F:/淘宝关键字/对比结果.xls'

#this code block's function is cut the word from csv  
wf = open('F:/clean_title.txt','w+')
first_line = True
count = 0
for line in open(search_key_words_path):
    if first_line:
        first_line = False
        continue
    count = count +1
    item = line.strip('\n\r').split('\t') #制表格切分
    item = item[0].split(',')
    #print(item[0])
    #print(count)
    tags = jieba.analyse.extract_tags(item[0]) #jieba分词
    #print(tags)
    tagsw = ",".join(tags) #逗号连接切分的词
    #print(tagsw)
    wf.write(tagsw)


#this code block's fuction is coutet the frequence of word
wf.seek(0,0)

 
article = wf.read()
#print(article)
dele = {'。','！','？','的','“','”','（','）',' ','》','《','，'}
jieba.add_word('大数据')
words = list(jieba.cut(article))
articleDict = {}
articleSet = set(words)-dele
for w in articleSet:
    if len(w)>1:
        articleDict[w] = []
        articleDict[w].append(words.count(w))
        sum = 0
        first_line = True
        for line in open('F:/宝贝无线APP搜索关键词表-530723471382-20180803.csv'):
                if first_line:
                    first_line = False
                    continue
                item = line.strip('\n\r').split('\t') #制表格切分
                item = item[0].split(',')
                if w in item[0]:
                    sum = sum + int(item[1].strip('"'))
        articleDict[w].append(sum)
articlelist = sorted(articleDict.items(),key = lambda x:x[1], reverse = True)
wf.close()

##for i in articlelist:
##    print(i[0],i[1][0],i[1][1])


if os.path.exists(save_result_file):
    os.remove(save_result_file)
    
n_wb = xlwt.Workbook()
sh = n_wb.add_sheet('sheet1')
n_wb.save(save_result_file)

    

rb = xlrd.open_workbook(business_staf_platform) #打开文件
#r_sheet = rb.sheet_by_index(0)
wb = copy(rb)
w_sheet = wb.get_sheet(0)
w_sheet.write(4,9,'关键字')
w_sheet.write(4,10,'词频')
w_sheet.write(4,11,'访客数')
i = 0
for xxx in articlelist:
    i = i+1
    w_sheet.write(4+i,9,xxx[0])
    w_sheet.write(4+i,10,xxx[1][0])
    w_sheet.write(4+i,11,xxx[1][1])
wb.save(save_result_file)

    

