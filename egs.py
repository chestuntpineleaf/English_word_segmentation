# -*- coding: utf-8 -*-
"""
Created on Wed Mar 28 16:49:38 2018
@author: 47899

modified on Sat April 11
@author PineLeaf
"""
import codecs
import os
import nltk
import math
import operator
import pandas as pd
from nltk.tokenize import WordPunctTokenizer

def excel_one_line_to_list(path,col):
    df = pd.read_excel(path,usecols=[col],names=None)  # 读取项目名称列,不要列名
    df_li = df.values.tolist()
    result = []
    for s_li in df_li:
        result.append(s_li[0])
        # print(s_li[0])
    return result


def participles(text):  # 分词函数
    pattern = r"""(?x)               # set flag to allow verbose regexps 
              (?:[A-Z]\.)+           # abbreviations, e.g. U.S.A. 
              |\d+(?:\.\d+)?%?       # numbers, incl. currency and percentages 
              |\w+(?:[-']\w+)*       # words w/ optional internal hyphens/apostrophe 
              |\.\.\.                # ellipsis 
              |(?:[.,;"'?():-_`])    # special characters with meanings 
            """
    t = nltk.regexp_tokenize(text, pattern)
    length = len(t)
    for i in range(length):
        t[i] = t[i].lower()
    return t


def getridofsw(lis, swlist):  # 去除文章中的停用词
    afterswlis = []
    for i in lis:
        if str(i) in swlist:
            continue
        else:
            afterswlis.append(str(i).lower())
    return afterswlis


def fun(filepath):  # 遍历文件夹中的所有文件，返回文件list
    arr = []
    for root, dirs, files in os.walk(filepath):
        for fn in files:
            arr.append(root + "\\" + fn)
    return arr


def read(path):  # 读取txt文件，并返回list
    with codecs.open(path, 'r', 'ANSI') as f:
        data = f.read()
    return data


def readstop(path):  # 读取txt文件，并返回list
    f = open(path, encoding='utf-8')
    data = []
    for line in f.readlines():
        data.append(line)
    return data


def getstopword(path):  # 获取停用词表
    swlis = []
    for i in readstop(path):
        outsw = str(i).replace('\n','').lower()
        swlis.append(outsw)
    return swlis


def freqword(wordlis):  # 统计词频，并返回字典
    freword = {}
    for i in wordlis:
        if str(i) in freword:
            count = freword[str(i)]
            freword[str(i)] = count + 1
        else:
            freword[str(i)] = 1
    return freword


def corpus(filelist, swlist):  # 建立语料库
    alllist = []
    for i in filelist:
        afterswlis = getridofsw(participles(str(i)), swlist)
        alllist.append(afterswlis)
    return alllist


def tf_idf(wordlis, filelist, corpuslist):  # 计算TF-IDF,并返回字典
    outdic = {}
    tf = 0
    idf = 0
    dic = freqword(wordlis)
    # outlis = []
    for i in set(wordlis):
        tf = dic[str(i)] / len(wordlis)  # 计算TF：某个词在文章中出现的次数/文章总词数
        # 计算IDF：log(语料库的文档总数/(包含该词的文档数+1))
        idf = math.log(len(filelist) / (wordinfilecount(str(i), corpuslist) + 1))
        tfidf = tf * idf  # 计算TF-IDF
        outdic[str(i)] = tfidf
    orderdic = sorted(outdic.items(), key=operator.itemgetter(1), reverse=True)  # 给字典排序
    return orderdic


def wordinfilecount(word, corpuslist):  # 查出包含该词的文档数
    count = 0  # 计数器
    for i in corpuslist:
        for j in i:
            if word in set(j):  # 只要文档出现该词，这计数器加1，所以这里用集合
                # if j.__contains__(word):
                count = count + 1
            else:
                continue
    return count


def befwry(lis):  # 写入预处理，将list转为string
    outall = ''
    for i in lis:
        ech = str(i).replace("('", '').replace("',", '\t').replace(')', '')
        outall = outall + '\t' + ech + '\n'
    return outall


# def wry(txt, path):  # 写入txt文件
#     f = codecs.open(path, 'a', 'utf-8')
#     f.write(txt)
#     f.close()
#     return path

# 追加数据
import xlrd
from xlutils.copy import copy
path = r'/Users/pineleaf/Desktop/OLED简化 2.xls'
def write_excel_xls_append(path, i,value):
    # index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    # sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    # worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    # rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
#     for i in range(0, index):
#         for j in range(0, len(value[i])):
#         new_worksheet.write(i+rows_old, j, value[7][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_worksheet.write(i,5,value)
    new_workbook.save(path)  # 保存工作簿
    print("xls格式表格【追加】写入数据成功！")

def main():
    j = 1;
    swpath = r'/Users/pineleaf/PycharmProjects/English_word_segmentation/English_stop_words.txt'  # 停用词表路径
    swlist = getstopword(swpath)  # 获取停用词表列表
    # print(swlist)
    filelist = excel_one_line_to_list("/Users/pineleaf/Desktop/OLED简化.xls",4)
    corpuslist = corpus(filelist, swlist)
    # print(corpuslist)
    outall = ''
    # wrypath = r'/Users/pineleaf/Desktop/TFIDF2.txt'
    for i in filelist:
        afterswlis = getridofsw(participles(str(i)), swlist)  # 获取每一篇已经去除停用的词表
        tfidfdic = tf_idf(afterswlis, filelist, corpuslist)  # 计算TF-IDF
        titleary = str(i).split('\\')
        title = str(titleary[-1]).replace('utf8.txt', '')
        echout = '\n' + befwry(tfidfdic)
        print(title + ' is ok!')
        outall = outall + echout
        write_excel_xls_append(path,j,outall)
        j = j+1
    # print(wry(outall, wrypath) + ' is ok!')

main();