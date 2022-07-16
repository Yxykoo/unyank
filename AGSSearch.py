#!/usr/bin/env python3
# -*- coding: UTF-8  -*-
from selenium import webdriver
from selenium.webdriver.common.by import By

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
import pandas as pd
import getpass
from selenium.webdriver.chrome.options import Options
import os
import pdb
import warnings
from pandas.core.common import SettingWithCopyWarning
warnings.simplefilter(action="ignore", category=SettingWithCopyWarning)
class Yanked(object):
    def __init__(self,yanked_df):
        self.yanked_df = yanked_df
        self.yanked_df.rename(columns={'Destination':'TargetMarketplace','Source':'SourceMarketplace'},inplace=True)
        self.yanked_df = yanked_df.drop_duplicates(['ASIN','SourceMarketplace','TargetMarketplace'])
        self.split_pattern = re.compile('\s+|\t')
        chrome_options = Options()
        # chrome_options.add_argument("--headless")
        if getpass.getuser() =='xinyyang':
            chrome_options.binary_location = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
        chrome_options.add_argument('disable-infobars')
        # self.user_pattern = re.compile('user:(\w+)')
        chrome_path = os.environ['USERPROFILE'] + r'\AppData\Local\Google\Chrome\User Data'
        chrome_options.add_argument("user-data-dir=" + chrome_path)
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 30)
    def get_yanked_url(self,source,target,asin):
        #ASIN	SourceMarketplace	TargetMarketplace

        return 'https://123.amazon.com/123?123='+str(source)+'&123='+str(target)+'&asin='+str(asin)

    def get_yanked_user(self):
        asin_list = self.yanked_df[['ASIN','SourceMarketplace','TargetMarketplace']].values
        # print(asin_list)
        user_list =[]
        count = 0
        for (asin,source,target) in asin_list:
            count +=1
            print(count)
            yanked_url = self.get_yanked_url(source,target,asin)
            try:
                self.driver.get(yanked_url)
            except:
                print('Yank页面加载异常，正在刷新页面')
                self.driver.refresh()

            try:
                table = self.wait.until(
                    EC.presence_of_element_located((
                        By.TAG_NAME, 'table'
                    )))
                table_text = table.text
            except:
                table_text = ''
            if 'OfferRemoved' not in table_text:
                user_list.append('')
            else:
                table_content = table_text.split('\n')
                # print(table_content)
                for  tr in table_content:
                    # pdb.set_trace()
                    # if 'OfferRemoved' in tr:
                    #     project_id = self.split_pattern.split(tr)[2]
                    #     remove_url = 'https://ags.amazon.com/request?sourceMarketplace='+str(source)+'&id='+str(project_id)
                    #     self.driver.get(remove_url)
                    #     body = self.wait.until(
                    #         EC.presence_of_all_elements_located((
                    #             By.TAG_NAME, 'dd'
                    #         )))
                    #     body_text = body[-1].text
                    #     user_list.append(body_text)
                    #     pdb.set_trace()
                    #     break
                    # pdb.set_trace()
                    if ('OfferBlockRemoved' in tr) or ('SyndicatedAndBuyable' in tr):
                        #如果offerblockremoved出现在offerremoved前面，则判断为coupon yank，这里append空字符串的原因是后面的程序会将空字符串判断为coupon yank
                        user_list.append('')
                        break
                    else:
                        if 'OfferRemoved' in tr:
                            # pdb.set_trace()
                            #offerblockremoved出现在offerremoved后面，则把offerremoved 后面的字符串加到user_list中
                            # user_list.append(tr.split('OfferRemoved ',1)[1])
                            # pdb.set_trace()

                            project_id = self.split_pattern.split(tr)[2]
                            remove_url = 'https://ags.amazon.com/request?sourceMarketplace=' + str(
                                source) + '&id=' + str(project_id)
                            self.driver.get(remove_url)
                            body = self.wait.until(
                                EC.presence_of_all_elements_located((
                                    By.TAG_NAME, 'dd'
                                )))
                            body_text = body[-1].text
                            if ('gspromotiongateway' in body_text) or ('globalstorecxdealspikeandkillplugin' in body_text):
                                user_list.append(body_text)
                            else:
                                # pdb.set_trace()
                                user_list.append(tr[44:])
                            break

                            # print(tr.split('OfferRemoved ',1)[1])


        # self.yanked_df.loc[self.yanked_df.index,'Blocked'] = user_list
        self.yanked_df['Blocked'] = user_list
        self.driver.quit()
        self.yanked_df.rename(columns={'TargetMarketplace':'Destination','SourceMarketplace':'Source'},inplace=True)
        self.yanked_df.fillna('',inplace=True)

if __name__ == '__main__':
    full_yanked = pd.read_excel('yank user check.xlsx')
    # full_yanked = full[full['Blocked']=="[Blocked]"]#改成blocked的情况
    yank = Yanked(full_yanked)
    yank.get_yanked_user()
    yank.yanked_df.to_csv('yanked_result.csv',index=False)



