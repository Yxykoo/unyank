import time
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import os
from selenium.webdriver.support import expected_conditions as EC
class AgsOperation:
    def __init__(self):
        self.source_market = [{'DE':'dropdown1_0'},{'AE':'dropdown1_1'},{'JP':'dropdown1_2'},{'GB':'dropdown1_3'},{'US':'dropdown1_4'}]
        self.destination_market = [{'CN':'dropdown2_0','GB':'dropdown2_1','TR':'dropdown2_2'},{'SA':'dropdown2_0'},{'CN':'dropdown2_0'},
                      {'CN':'dropdown2_0','US':'dropdown2_1','AU':'dropdown2_2','DE':'dropdown2_3','AE':'dropdown2_4','SA':'dropdown2_8'},
                      {'CN':'dropdown2_0','MX':'dropdown2_1','SA':'dropdown2_2','GB':'dropdown2_4','AU':'dropdown2_5','SG':'dropdown2_8','AE':'dropdown2_9','DE':'dropdown2_7'}]
        self.url = "https://123.amazon.com/123"
        self.source_id = {'US':'1','UK':'3','JP':'6','DE':'4','AE':'338801'}
        self.target_id = {'CN':'3240','US':'1','UK':'3','AU':'111172','MX':'771770','SG':'104444012','SA':'338811','AE':'338801','DE':'4','TR':'338851'}
        self.path = os.listdir(r'C:\Users\xinyyang\Desktop\UK-US_Selection_Health\UNYANK\New folder')
        self.root = r'C:\Users\xinyyang\Desktop\UK-US_Selection_Health\UNYANK\New folder'
        self.chrome_options = Options()
        self.chrome_options.binary_location = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
        self.chrome_path = os.environ['USERPROFILE'] + r'\AppData\Local\Google\Chrome\User Data'
        self.chrome_options.add_argument("user-data-dir=" + self.chrome_path)
        self.driver = webdriver.Chrome(options=self.chrome_options)

    def ChooseSourceDestination(self,source, destination):
        for i in range(len(self.source_market)):
            if source in self.source_market[i]:
                # 查看source的按钮的ID: print(source_market[i][source])
                source_button = self.source_market[i][source]
                index = i
                break
        # import pdb
        # pdb.set_trace()
        destination_button = self.destination_market[index][destination]
        # 查看destination的按钮的ID: print(destination_market[index][destination])
        # print(self.destination_market[index][destination])
        return [source_button, destination_button]

    def FindRequestId(self,source, target):
        print('Source: ' + source, 'Destination: ' + target,'Request ID: ' + 'https://123.amazon.com/request?id=' + self.driver.find_element_by_xpath("//a[@target='_blank']").text + '&sourceMarketplace=' + self.source_id[self.ConversionBack(source)])
        list = [source,target,'https://123.amazon.com/request?id=' + self.driver.find_element_by_xpath("//a[@target='_blank']").text + '&sourceMarketplace=' + self.source_id[self.ConversionBack(source)]]
        return list
    def Conversion(self,name):
        if name == 'UK':
            return 'GB'
        else:
            return name

    def ConversionBack(self,name):
        if name == 'GB':
            return 'UK'
        else:
            return name

    def CheckName(self,name):
        source_name = 0
        target_name = 0
        for sou in self.source_id:
            if name.find(sou, 0, 2) != -1:
                source_name = sou

        for tar in self.target_id:
            if name.find(tar, 2) != -1:
                target_name = tar

        if source_name != 0 and target_name != 0:
            return [source_name, target_name]
        else:
            return [0, 0]

    def CheckListTime(self,name, source, target):

        file_address = os.path.join(self.root, name)
        # 检查时间
        file_time = time.localtime(os.path.getmtime(file_address))
        system_time = time.localtime(time.time())

        if source != 0 and target != 0 and file_time.tm_year == system_time.tm_year and file_time.tm_mon == system_time.tm_mon and file_time.tm_mday == system_time.tm_mday:
            return [file_address, source, target]
        else:
            return [0, 0, 0]
    def Operation(self):
        self.driver.get(self.url)
        WebDriverWait(self.driver, 90).until(EC.title_contains("AmazonGlobalStore"))
        infolist=[]
        for pa in self.path:
            source_name, target_name = self.CheckName(pa)
            address, checked_source, checked_target = self.CheckListTime(pa, source_name, target_name)
            if (address != 0 and checked_source != 0 and checked_target != 0):
                source_button, destination_button = self.ChooseSourceDestination(self.Conversion(checked_source),self.Conversion(checked_target))
                source = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Source Marketplace']")))
                self.driver.execute_script("arguments[0].click();", source)
                try:
                    source_click = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.ID, source_button)))
                except:
                    self.driver.refresh()
                    source = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Source Marketplace']")))
                    self.driver.execute_script("arguments[0].click();", source)
                    source_click = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.ID, source_button)))
                self.driver.execute_script("arguments[0].click();", source_click)
                market = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Destination Marketplace']")))
                self.driver.execute_script("arguments[0].click();", market)
                market_click = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.ID, destination_button)))
                self.driver.execute_script("arguments[0].click();", market_click)
                input_button = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.ID, 'file')))
                input_button.send_keys(address)
                reason = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.ID, 'reason')))
                time.sleep(1)
                reason.send_keys("aee-gs-order-spike-report@amazon.com")
                time.sleep(1)
                tags = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.ID, 'tags')))
                tags.send_keys("aee-gs-order-spike-report@amazon.com")
                time.sleep(2)
                submit_button = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit']")))
                self.driver.execute_script("arguments[0].click();", submit_button)
                time.sleep(2)
                #把request id 的信息用append放到一个列表里，把Operation这个函数加个返回值，然后传给操作excel的class
                request_id = self.FindRequestId(checked_source, checked_target)
                print(request_id)
                infolist.append(request_id)
                time.sleep(2)
                self.driver.get(self.url)
            else:
                pass
            time.sleep(1)
        self.driver.quit()
        print('Upload to AGS Done')
        return infolist
if __name__ == '__main__':
    ags = AgsOperation()
    info = ags.Operation()