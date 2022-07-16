import pandas as pd
import win32com.client as win32
import openpyxl
class Excel:
    def __init__(self):
        Excel_name = 'C:/Users/xinyyang/Desktop/UK-US_Selection_Health/UNYANK/New folder/Pre.xlsx'
        self.info = pd.read_excel(Excel_name, sheet_name='Pretreatment')

    def ChangeFormat(self):
        body = []
        temp = []
        for i in range(len(self.info.columns)):
            temp.append(self.info.columns[i])
        # print(temp)
        body.append(temp)
        for i in range(len(self.info.index)):
            temp = []
            for j in range(len(self.info.columns)):
                temp.append(self.info.iloc[i][j])
            # print(temp)
            body.append(temp)
            # print(body)
        return body




class Email:
    def __init__(self,data):
        self.outlook = win32.Dispatch('outlook.application')
        self.data = data
        # self.messages = self.outlook.GetDefaultFolder(6).Folders['Unyank1'].Items
    def Send(self, html):
        message = self.outlook.CreateItem(0)
        message.Display()
        #邮箱和主题要改
        message.To = "xinyyang@amazon.com"
        message.Subject = "Please view"
        message.HTMLBody = html


    def html(self):
        count =0
        html_body = '<html><head><style>table,th,td{border: 2px solid black; border-collapse:collapse;}p{margin-top: 0.1em;margin-bottom: 0.1em;}p.a{text-indent: 35px;margin-bottom:20px;display:inline-block;}p.b{margin-top:20px;display:inline-block;}</style></head><body><p>Dears,</p><p class="a">Un-yankable offers for following ASINs have been completed.</p>'
        html_body += '<body><table>'
        html_body +='<tr style = "background-color:rgb(213,250,210);">'
        for i in range(6):
            html_body += '<th>'
            html_body +=  self.data[0][i]
            html_body += '</th>'
        html_body += '</tr>'

        for j in range(len(self.data)-1):
            html_body += '<tr>'
            for i in range(6):
                html_body+='<td>'
                html_body+= str(self.data[j+1][i])
                html_body +='</td>'
            html_body += '</tr>'
        html_body +='</table>'
        html_body+='<p class ="b" >Best Regards,</p>'
        html_body+='<p>Xinyue Yang</p>'
        html_body+='<p>RBS AGS Team</p>'
        html_body+='</body>'
        return html_body
if __name__ == '__main__':
    Emailbody = Excel()
    text = Emailbody.ChangeFormat()
    outlook = Email(text)
    html = outlook.html()
    outlook.Send(html)