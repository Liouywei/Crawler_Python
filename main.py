import os
import sys
import time
#爬蟲部分
import pandas as pd
import requests
from requests_html import HTML
import re
import numpy as np
import matplotlib.pyplot as plt
import xlrd
import xlwt
import openpyxl
from sklearn.cluster import KMeans
from sklearn import cluster, datasets, metrics
#install xlrd
#install xlwt
#install openpyxl


#每次執行完爬蟲，重啟程式
def restart_program():
    python = sys.executable
    os.execl(python, python, * sys.argv)

##########################################################
#----------------------爬蟲函式部分----------------------#
##########################################################
def CreateEmptyExcel():
  emptyE = pd.DataFrame()
  emptyE.to_excel('Travel.xlsx', index=False)

#讀取csv檔，並更新latest_ID和newest_ID
def UpdateValue():
  global newest_ID
  global latest_ID
  if len(read.index)==0:
    newest_ID=' '
    latest_ID=' '
  else:
    newest_ID = str(read.loc[0,'ID'])            # 找出爬出資料的第一筆ID
    latest_ID = str(read.loc[len(read)-1,'ID'])  # 找出爬出資料的最後一筆ID
    print('Excel表目前最新文章ID : ' + newest_ID)
    print('Excel表目前最舊文章ID : ' + latest_ID)

# 透過輸入文章ID，就輸出文章的資料
def Crawl(ID):
    link = 'https://www.dcard.tw/_api/posts/' + str(ID)
    requ = requests.get(link)
    rejs = requ.json()
    return(pd.DataFrame(
        data=
        [{'ID':rejs['id'],
          'title':rejs['title'],
          'createdAt':rejs['createdAt'],
          'updatedAt':rejs['updatedAt'],
          'commentCount':rejs['commentCount'],
          'gender':rejs['gender'],
          'likeCount':rejs['likeCount'],
          'topics':rejs['topics']}],
        columns=['ID','title','createdAt','updatedAt','commentCount','gender','likeCount','topics']))

# 抓取前100篇新文章
def GetNewData():
  url = 'https://www.dcard.tw/_api/forums/travel/posts?popular=false&limit=100'
  resq = requests.get(url)
  rejs = resq.json()

  global df
  global newest_ID
  num_of_newdata = 0
  for i in range(len(rejs)):
    if str(rejs[i]['id']) == newest_ID:
      if i>0:
        print('搜尋停止於 ' + str(rejs[i-1]['id']) + ' ' + str(rejs[i]['title']) + ', 總共爬取 ' + str(num_of_newdata) + ' 筆資料.')
        break 
      else:
        print('沒有搜尋到新的文章.')
        break
    else:
      df = df.append(Crawl(rejs[i]['id']))
      num_of_newdata+=1
      print('i=' + str(i) + '---文章ID : ' + str(rejs[i]['id']))

  if num_of_newdata > 0:
    final = df.append(read)  #將excel表的資料與新抓取的資料合併
    final.to_excel('Travel.xlsx', index=False)
    print('===> Excel儲存完成.')

# 抓取資料表最後一筆文章的後100篇文章
def GetOldData():
  global df
  global latest_ID
  print('最後一筆 ID : ' + str(latest_ID))
  if len(read.index) > 0:
    url = 'https://www.dcard.tw/_api/forums/travel/posts?popular=false&limit=100&before=' + latest_ID
    resq = requests.get(url)
    rejs = resq.json()
    for i in range(len(rejs)):
      df = df.append(Crawl(rejs[i]['id']))
      print('i=' + str(i) + '---文章ID : ' + str(rejs[i]['id']))

    final = read.append(df) #,ignore_index=True
    final.to_excel('Travel.xlsx', index=False)
    print('===> Excel儲存完成.')
  else:
    print('請先「爬取新的文章」, 方能搜尋舊的文章.')

#更新excel資料用爬蟲
def UpdateCrawl(ID):
    link = 'https://www.dcard.tw/_api/posts/' + str(ID)
    requ = requests.get(link)
    rejs = requ.json()
    if str(requ) == '<Response [404]>':
        print('ID : ' + str(ID) + ' 原始文章已被刪除.')
        return 'Null', 'Null', 'Null'
    else:
        print('ID : ' + str(ID) + ' 更新資料.')
        return rejs['updatedAt'], rejs['commentCount'], rejs['likeCount']

#更新excel表資料
def UpdateData():
    global read

    print('1. 更新整個資料表.')
    print('2. 更新部分資料.')
    choose=str(input('選擇要執行的模式 : '))
    if choose =='1':
      for i in range(len(read)):
          if i%100==0 and i!=0:
            keep = str(input("是否要繼續更新? (Y/y) : "))
            if(keep=='Y' or keep=='y'):
              print('等待10秒後繼續...' )
              time.sleep(10)
            else:
              print('更新結束，正在儲存至Excel檔.')
              read.to_excel('Travel.xlsx', index=False)
              return
          print('第 ' + str(i) + ' 筆, 建立日期:' + str(read.loc[i,'createdAt']) + ', 最後更新日期:' +  str(read.loc[i,'updatedAt']) + '-->', end=' ')
          x,y,z = UpdateCrawl(read.loc[i,'ID'])
          read.loc[i,'updatedAt'] = x if x!='Null' else read.loc[i,'updatedAt']
          read.loc[i,'commentCount'] = y if y!='Null' else read.loc[i,'commentCount']
          read.loc[i,'likeCount'] = z if z!='Null' else read.loc[i,'likeCount']
      print('更新結束，正在儲存至Excel檔.')
      read.to_excel('Travel.xlsx', index=False)
    elif choose =='2':
      pos = int(input('從第幾筆開始更新 : '))
      cnt = int(input('要更新多少筆資料 : '))
      total = pos + cnt
      if total >= len(read):
        total = len(read)  #避免輸入值超過資料數量
      for j in range(pos, total):
        print('第 ' + str(j) + ' 筆, 建立日期:' + str(read.loc[j,'createdAt']) + ', 最後更新日期:' +  str(read.loc[j,'updatedAt']) + '-->', end=' ')
        x,y,z = UpdateCrawl(read.loc[j,'ID'])
        read.loc[j,'updatedAt'] = x if x!='Null' else read.loc[j,'updatedAt']
        read.loc[j,'commentCount'] = y if y!='Null' else read.loc[j,'commentCount']
        read.loc[j,'likeCount'] = z if z!='Null' else read.loc[j,'likeCount']
      print('更新結束，正在儲存至Excel檔.')
      read.to_excel('Travel.xlsx', index=False)


##########################################################
#----------------------分析函式部分----------------------#
##########################################################
def ClusterFun():
  name = str(input('輸入檔名 : '))  #Travel.xlsx
  if os.path.isfile(name):
    k = pd.read_excel(name, sheet_name='Sheet1')
    x = k[['commentCount', 'likeCount']].values
    while True:
      os.system('CLS')
      print('1. 產生績效圖.')
      print('2. 使用 K-Means 進行分群.')
      print('*. 輸入其他值回到選單畫面.')
      print('---------------------------')
      num = int(input('輸入選項 : '))

      if num==1:
        Bar(x)
      elif num==2:
        KMeansFun(x,k)
        break
      else:
        break

def Bar(x):
  silhouette_avgs = []
  ks = range(2, 11)
  for k in ks:
    kmeans_fit = cluster.KMeans(n_clusters = k).fit(x)
    cluster_labels = kmeans_fit.labels_
    silhouette_avg = metrics.silhouette_score(x, cluster_labels)
    silhouette_avgs.append(silhouette_avg)
  plt.bar(ks, silhouette_avgs) # 作圖並印出 k = 2 到 10 的績效
  plt.show()

def KMeansFun(x,k):
  cnt = int(input('要分成幾群 : '))
  km = KMeans(n_clusters=cnt)  #K= cnt 群
  y_pred = km.fit_predict(x)
  plt.figure(figsize=(10, 6))
  plt.xlabel('commentCount')
  plt.ylabel('likeCount')
  plt.scatter(x[:, 0], x[:, 1], c=y_pred) #C是第三維度 已顏色做維度
  plt.show()

  s = str(input('是否儲存至Excel? (Y/y) : '))
  if s=='Y' or s=='y':
    name = str(input('檔案命名 (檔名.xlsx) : '))
    print('將分群結果儲存至檔名 : '+ name + ' 中.')
    #read.insert(8, column='Class',value='Null')
    col = 'Class' + str(len(k.columns) - 8)
    k.insert(len(k.columns), column=col ,value='Null')
    for i in range(len(y_pred)):
      k.loc[i,col] = y_pred[i]
    k.to_excel(name, index=False)
    print('儲存完成.')
    #print(km.cluster_centers_) #各群中心點(X,Y)的位置


##########################################################
#-----------------------主程式部分-----------------------#
##########################################################
latest_ID=''  #紀錄最舊一筆資料的ID
newest_ID=''  #紀錄最新一筆資料的ID
read = pd.DataFrame()
df = pd.DataFrame()
isfile = True
if os.path.isfile('Travel.xlsx'):
  read = pd.read_excel('Travel.xlsx', sheet_name='Sheet1')  #讀取excel紀錄 
  print('Excel表大小 : ' + str(read.shape))
  UpdateValue()
  print('======================================')
  print('1. 爬取新的文章 (每次100筆).')
  print('2. 爬取舊的文章 (每次100筆).')
  print('3. 更新已蒐集的文章資料.')
  print('4. Clustering  分析.')
  isfile = True
else:
  print('======================================')
  print('0. 建立Excel表 (檔名：Travel.xlsx).')
  print(' ')
  print(' ')
  print(' ')
  isfile = False
print('5. 關閉程式.')
print('======================================')

choose = str(input('Select The Number Of Option : '))
if choose=='1':
    print('爬蟲程式執行中...')
    GetNewData()        #抓最新
    os.system('pause')
    os.system('CLS')
    #restart_program()   #重啟程式
elif choose=='2':
    print('爬蟲程式執行中...')
    GetOldData()        #抓最舊
    os.system('pause')
    os.system('CLS')
    #restart_program()   #重啟程式
elif choose=='3':
    print('爬蟲程式執行中...')
    UpdateData()        #更新Excel表內容
    os.system('pause')
    os.system('CLS')
    #restart_program()   #重啟程式 
elif choose=='4':
    print('分析程式執行中...')
    ClusterFun()        #執行分群分析KMeans 
    os.system('pause')
    os.system('CLS')
    #restart_program()   #重啟程式
elif choose=='0' and isfile==False:
    print('建立Excel表中...')
    CreateEmptyExcel()  #建立Excel表
    os.system('pause')
    os.system('CLS')
    #restart_program()   #重啟程式




''' 計算 topic 的數量
read = pd.read_excel('Travel.xlsx', sheet_name='Sheet1')
x = read['topics'].values

lis=[]
cnt=[]
for i in range(len(x)):
  y = str(x[i]).split(",")
  for j in range(len(y)):
    if y[j] in lis:
      cnt[lis.index(y[j])] +=1
    else :
      lis.append(y[j])
      cnt.append(1)

f = list(zip(lis, cnt))
dfobj = pd.DataFrame(f, columns = ['topic', 'count'])
dfobj.to_excel('123.xlsx', index=False)
'''
