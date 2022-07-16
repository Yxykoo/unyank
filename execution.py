from reReadOutlook import *
from reArcChanger import *
from reAddToTxt import *
from AgsOperation import *
from Record import  *
from SendEmail import *
if __name__ =='__main__':
    email = Outlook()
    email.Operation()
    arc_changer = ArcChanger()
    arc_changer.AddExcel()
    # pdb.set_trace()
    txt = AddtoTxt()
    txt.TxtAdd()
    # import pdb
    # pdb.set_trace()
    ags = AgsOperation()
    info = ags.Operation()
    # info = [['UK', 'US', 'https://ags.amazon.com/request?id=92bae9f0-0c0e-8e19-2aaa-92ebbb464f0e&sourceMarketplace=3'], ['US', 'MX', 'https://ags.amazon.com/request?id=42bae9f0-258f-4494-72d1-e22150024185&sourceMarketplace=1']]
    print(info)
    if(info):
        record = Record(info)
        record.Operation()
        Emailbody = Excel()
        text = Emailbody.ChangeFormat()
        outlook = Email(text)
        html = outlook.html()
        outlook.Send(html)
    #给Operation这个函数加个返回值，返回一个带request id 的列，然后把它传给Record的类中
    #若
