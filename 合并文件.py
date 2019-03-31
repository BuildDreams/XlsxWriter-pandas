#!/usr/bin/env python
# coding: utf-8

# In[122]:


import pandas as pd 
import os
import sys
import time


# # <font color='blue'>查找目标文件夹下面的目标文件<font/>
# ### <font color='red'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;dirname:目标文件夹<font/>

# In[80]:


def all_path(dirname):
    result = []#所有的文件
    for maindir, subdir, file_name_list in os.walk(dirname):
#         print("1:",maindir) #当前主目录
#         print("2:",subdir) #当前主目录下的所有目录
#         print("3:",file_name_list)  #当前主目录下的所有文件
        for filename in file_name_list:
            apath = os.path.join(maindir, filename)#合并成一个完整路径
            result.append(apath)
    return result


# # <font color='blue'>合并文件逻辑 </font>
# ### <font color='red'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;files：目标文件下面的所有文件. <br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;result_name：生成的文件名字</font>

# In[125]:


def getTab(files,result_name):
    t0 = time.time()
    df = pd.read_excel(r'C:\Users\zq\Desktop\Project\人力成本表格式汇总.xlsx')
    df = df.T
    for i, file in enumerate(files) :
        nums = len(files) - i -1
        space = " "*10
        
        print('\033[34m加载文件%s,\n %s剩余：\033[32m %s 文件'%(file, space, nums))
        t1  =  time.time()
        sys.stdout.write('\r')
        new_df = pd.read_excel(file,sheet_name='人力明细').T
        t2 = time.time()
        
        file_size = os.path.getsize('./2017.xlsx')/float(1024*1024)
        print('\033[31m加载完成，文件大小%s M,用时%s秒'%(file_size, t2-t1))
        df = pd.merge(df,new_df,how='left',sort=True,left_index=True,right_index=True)
        print("\033[36m合并完成，用时%s秒"%(time.time()-t2))
        sys.stdout.write('\r')
        sys.stdout.flush()
    df = df.T
    df0 = pd.read_excel(r'C:\Users\zq\Desktop\Project\人力成本表格式汇总.xlsx')
    use_cols =df0.columns.tolist()
    df4=df.loc[:,use_cols]#调整输出的顺序
    df4 = df4.reset_index(drop=True)
    writerTab(df4,result_name)
    print('\033[31m全部运行完成，总计耗时%s 秒 '% (time.time()-t0))
    


# # <font color='blue'>指定写入格式<font/>
# ### <font color='red'> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;df:要写入的dataFrame类型。<br/> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;result_name:生成文件的名字。<font/>

# In[120]:


def writerTab(df, result_name):
    print('\033[35m正在写入文件....')
    writer = pd.ExcelWriter('%s.xlsx'%result_name, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False)
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
#     cell_format = workbook.add_format({'bold': True })
    worksheet.set_row(0, 30)#设置行的高度
    header_format = workbook.add_format({
        'bold': True,
        'font_color': 'red',
        'text_wrap': True,
        'valign': 'vcenter',
        'align': 'center',
        'fg_color': '#D7E4BC',
        'border': 3})
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num + 1, value, header_format)
    writer.save()
    print('\033[35m写入完成！')


# # <font color='blue'>程序主入口</font>

# In[127]:


if __name__ == "__main__":
    files= all_path(r'./19年华为群1-2月成本表定稿')
    result_name = '2017'
    getTab(files,result_name)

