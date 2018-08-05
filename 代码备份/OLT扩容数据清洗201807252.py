#import numpy as np
import pandas as pd
import datetime
import os
os.chdir("D:/学习/Python/MyCodes/20180723OLT扩容数据清洗/")   #设定工作目录

# 读取原始数据
df0 = pd.read_excel('OLT基础信息20180725.xls', engine = 'xlrd',
                    sheet_name='OLT基础信息', skiprows=[0,1], usecols="B:DY")  #表格顶部两行是表头，A列是空列，都不读取


## 判断“完成目标网扩容”和“扩容施工进度”字段的准确性
#df_mbw1 = df0[['OLT名称' , '上联扩容完成情况' , '扩容施工进度' ,'上行扩容规划情况','规划扩容目标局向']]
#df_mbw1['上联扩容完成情况_修改为'] = ''   # 这两列用于记录准确的数据，以便系统后台修改
#df_mbw1['扩容施工进度_修改为'] = ''
#df_mbw1 = df_mbw1[['OLT名称' , '上联扩容完成情况' , '扩容施工进度' ,'上联扩容完成情况_修改为','扩容施工进度_修改为','上行扩容规划情况','规划扩容目标局向']] #调整顺序，方便Excel表操作
#df_mbw2 = df_mbw1[((df_mbw1['上联扩容完成情况'] == '完成目标网扩容') ^ (df_mbw1['扩容施工进度'] == '10：已完成目标扩容'))]  #判断逻辑：两个字段表达的内容不一致
#df_mbw3 = pd.merge(df_mbw2 , df0 , on='OLT名称',suffixes=('', '_y'))
#df_mbw3.to_excel('OLT扩容异常数据1完成情况%s.xls' % datetime.datetime.now().strftime("%y-%m-%d-%H-%M"), sheet_name='完成目标网扩容字段异常数据',index = 0)



## 完成目标网扩容后，将历史峰值'上行扩容规划情况'和'(历史)峰值带宽利用率(%)'按照最近的值进行重置
#df_reset1 = df0[['OLT名称' ,'(历史)上联口峰值流量(M)', '(历史)峰值带宽利用率(%)',
#                '(最近)上联口峰值流量(M)','(最近)峰值带宽利用率(%)','上行扩容规划情况', '上联扩容完成情况','规划扩容目标局向','网络域类型']]
#
#bins = [  0  ,  700 ,  1300  ,  8000 , 12000 , 22000]    #通过流量和利用率数据计算得出的端口带宽，将其归类为具体的级别
#labels=['异常1GE','1GE','异常1-10GE','10GE','异常10GE']
#df_reset1['OLT上行带宽计算_历史'] = df_reset1['(历史)上联口峰值流量(M)'] / df_reset1['(历史)峰值带宽利用率(%)'] * 100
#df_reset1['OLT上行带宽计算_最近'] = df_reset1['(最近)上联口峰值流量(M)'] / df_reset1['(最近)峰值带宽利用率(%)'] * 100
#df_reset1['OLT上行带宽级别_历史'] = pd.cut(df_reset1['OLT上行带宽计算_历史'], bins = bins, labels = labels)
#df_reset1['OLT上行带宽级别_最近'] = pd.cut(df_reset1['OLT上行带宽计算_最近'], bins = bins, labels = labels)
#df_reset1['处理意见'] = ''
#df_reset1 = df_reset1[['OLT名称' ,'(历史)上联口峰值流量(M)', '(历史)峰值带宽利用率(%)','(最近)上联口峰值流量(M)','(最近)峰值带宽利用率(%)',
#                       'OLT上行带宽计算_历史','OLT上行带宽计算_最近','OLT上行带宽级别_历史','OLT上行带宽级别_最近',
#                       '处理意见','上行扩容规划情况', '上联扩容完成情况','规划扩容目标局向','网络域类型']]
#df_reset2 = df_reset1[df_reset1['OLT上行带宽级别_历史'] != df_reset1['OLT上行带宽级别_最近']]  #判断逻辑：两个字段表达的内容不一致
#df_reset3 = pd.merge(df_reset2 , df0 , on='OLT名称', suffixes=('', '_y'))
#df_reset3.to_excel('OLT扩容异常数据2峰值重置%s.xls' % datetime.datetime.now().strftime("%y-%m-%d-%H-%M"),sheet_name='峰值流量和利用率字段异常数据',index = 0)


# 梳理网络域的对应信息
##未实现的想法：根据端口1到6提取出来的名称，看看归属BNG信息是否正确
df_wly = df0[['OLT名称' , '上联扩容完成情况' ,'上行扩容规划情况','规划扩容目标局向','原上联','网络域类型',
               '下接盒式OLT','归属框式OLT','归属BNG','归属BNG2','归属SW',
               '城域网端口1','端口速率(Mb/s)1','下行峰值利用率1',
               '城域网端口2','端口速率(Mb/s)2','下行峰值利用率2',
               '城域网端口3','端口速率(Mb/s)3','下行峰值利用率3',
               '城域网端口4','端口速率(Mb/s)4','下行峰值利用率4',
               '城域网端口5','端口速率(Mb/s)5','下行峰值利用率5',
               '城域网端口6','端口速率(Mb/s)6','下行峰值利用率6']]
df_wly['下接盒式OLT'].fillna('缺失数据', inplace = True)
df_wly['归属框式OLT'].fillna('缺失数据', inplace = True)
#df_wlyhs1 = df_wly[(df_wly['下接盒式OLT'] != '缺失数据')][['OLT名称','下接盒式OLT']]  #去除不必要的字段，增加结果的可读性
#df_wlyhs2 = df_wly[(df_wly['归属框式OLT'] != '缺失数据')][['OLT名称','归属框式OLT']]
#df_wlyhs3 = pd.merge(df_wlyhs1,df_wlyhs2,how = 'outer',left_on='OLT名称',right_on='归属框式OLT')
#df_wlyhs3.to_excel('OLT扩容异常数据3网络域%s.xls' % datetime.datetime.now().strftime("%y-%m-%d-%H-%M"),sheet_name='盒式OLT级联异常数据',index = 0)

# 利用率数据是否异常，或者没有监控到

# 梳理C16及以前完成扩容情况的信息
df_wly['上行扩容规划情况'].fillna('缺失数据', inplace = True)
#df_c16 = df_wly[df_wly['上行扩容规划情况'].str.contains('C16') | df_wly['上行扩容规划情况'].str.contains('C15')]
#df_c16a = df_c16[(df_c16['规划扩容目标局向'] != df_c16['归属BNG']) & (df_c16['规划扩容目标局向'] != df_c16['归属SW'])]
#df_c16a.to_excel('OLT扩容异常数据4C16局向标准化%s.xls' % datetime.datetime.now().strftime("%y-%m-%d-%H-%M"),sheet_name='扩容目标局向异常数据',index = 0)

df_c18 = df_wly[df_wly['上行扩容规划情况'].str.contains('C17') | df_wly['上行扩容规划情况'].str.contains('C18')]
#df_c18a = df_c18[(df_c18['规划扩容目标局向'] != df_c18['归属BNG']) & (df_c18['规划扩容目标局向'] != df_c18['归属SW'])]

df_q = pd.read_excel('（待办）规划扩容目标局向问题数据.xls', engine = 'xlrd',sheet_name=0)
df_c18a = pd.merge(df_q,df_c18,left_on='OLT_NAME',right_on='OLT名称')

df_c18a.to_excel('OLT扩容异常数据5C18局向标准化%s.xls' % datetime.datetime.now().strftime("%y-%m-%d-%H-%M"),sheet_name='扩容目标局向异常数据',index = 0)



'''
# 导出需要新增扩容规划的OLT清单
##从原始数据筛选相关字段
df_ghxq1 = df0[['OLT名称', '状态', '式样', '设备型号',
           '(历史)上联口峰值流量(M)', '(最近)上联口峰值流量(M)',
          '(历史)峰值带宽利用率(%)', '(最近)峰值带宽利用率(%)',
          '上行扩容规划情况', '规划扩容目标局向', '原上联',
          '上联扩容完成情况','完成目标网扩容时间', '扩容施工进度']]  #这里默认所选的列数据都是准确的，异常数据另行处理

##根据历史带宽利用率和上行规划情况筛选：
###①历史带宽利用率≥30%, 且未有规划
df_ghxq2 = df_ghxq1[(df_ghxq1['(历史)峰值带宽利用率(%)'] >= 30) & (df_ghxq1['上行扩容规划情况'] == '--')]

###计算OLT上行带宽级别，辅助分析
df_ghxq2['OLT上行带宽计算'] = df_ghxq2['(历史)上联口峰值流量(M)'] / df_ghxq2['(历史)峰值带宽利用率(%)'] *100
df_ghxq2['OLT上行带宽级别'] = pd.cut(df_ghxq2['OLT上行带宽计算'], bins = bins, labels = labels)

##②针对盒式OLT，虽然已有规划2GE扩容，但是仍然需要提升至扩容10GE
# 保存数据到Excel
df_ghxq2.to_excel('OLT扩容异常数据4规划需求%s.xls' % datetime.datetime.now().strftime("%y-%m-%d-%H-%M"),sheet_name='未有规划异常数据',index = 0)
'''





print('finish!')