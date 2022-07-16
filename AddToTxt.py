import pandas as pd
import openpyxl
import os
import time
# def filter()
Excel_name = 'C:/Users/xinyyang/Desktop/UK-US_Selection_Health/UNYANK/Pre.xlsx'
Excel_file = 'C:/Users/xinyyang/Desktop/UK-US_Selection_Health/UNYANK'
#目前处理的source只有4个，以后如果出现了变多的情况再加，加的时候注意SA（KSA）
source_id =['US','UK','JP','DE']
target_id = ['CN','US','UK','AU','MX','SG','SA','AE']
def assgin_Arc(source,target):
    list=[]
    for i in range(len(source)):
        for j in range(len(target)):
            list.append(source[i] + '-' + target[j])
    return list
def add_arc(source_df):
    for i in range(len(source_df.index)):
        source_df['ARC'] = source_df['Source'] + '-' + source_df['Destination']
    return source_df
def choose_asins(asins):
    list=[]
    for i in range(len(asins.index)):
        list.append(asins.iloc[i])
    return list

def FindArc(inflow):
    list=[]
    for arc in inflow['ARC']:
        if arc not in list:
            list.append(arc)
    return list
def MatchDesiredArc(todayarc,standardarc):
    list = []
    for arc in todayarc:
        if arc in standardarc:
            list.append(arc)
    return list
def choose_file_and_add(file,arclist,data):
    for arc in arclist:
        # print(arc,file,file.find(arc,0,5))
        if file.find(arc,0,5) != -1:
            file_address = os.path.join(Excel_file,file)
            #这里必须用file的绝对地址，相对地址电脑识别不出来，因为在函数里面，只传文件名字它会在本地文件夹找
            open_file = open(file_address,mode='w')
            asin_list_o = data[data['ARC'] == arc]
            asin_list = choose_asins(asin_list_o['ASIN'])
            for a in asin_list:
                a = a + '\n'
                open_file.writelines(a)
            print(arc,asin_list,asin_list_o)
            open_file.close()
def DestinationChanger(data):
    for i in range(len(data.index)):
        if (data.iloc[i][2] == 'KSA'):
            data.iloc[i][2] = 'SA'
    print(data)
    return data
all_arc_list = assgin_Arc(source_id,target_id)
print(all_arc_list)
result_df = pd.read_excel(Excel_name,sheet_name='Result')
#一个把KSA改成SA的功能
result_df = DestinationChanger(result_df)
result_df_arc_added = add_arc(result_df)
result_df_arc_added = DestinationChanger(result_df_arc_added)
# final = result_df_arc_added[result_df_arc_added['ARC'] == 'JP-CN']
# final_asin_series = final['ASIN']
# final_asin = choose_asins(final_asin_series)
arc_list_of_today_asins = FindArc(result_df_arc_added)
print(arc_list_of_today_asins)
Checked_arc = MatchDesiredArc(arc_list_of_today_asins,all_arc_list)
print(Checked_arc)
path = os.listdir(Excel_file)
for pa in path:
    choose_file_and_add(pa,Checked_arc,result_df_arc_added)

