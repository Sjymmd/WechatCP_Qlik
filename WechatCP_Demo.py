from WechatCP_API import *

def Qlik_Push_Excel():

    import pandas as pd
    weChatEnterprise = WeChatEnterprise(corpsecret=CORPSECRET2)
    DataInf = pd.read_excel("push.xlsx")
    for x in range(len(DataInf)):
        userid = DataInf.iloc[x, 0]
        kwargs = {
            "articles": [
                {
                    "title": DataInf.iloc[x, 2],
                    "description": DataInf.iloc[x, 3],
                    "url": DataInf.iloc[x, 4],
                    "picurl": DataInf.iloc[x, 5]
                }
            ]
        }
        content = 'Hello！%s,以上您收到的是一份Qlik报表测试，请查阅！' % DataInf.iloc[x, 1]
        weChatEnterprise.send_msg_to_user(touser=[userid], content=content, msgtype="news", **kwargs)

def Qlik_Department():

    weChatEnterprise = WeChatEnterprise(corpsecret=CORPSECRET)
    state, res = weChatEnterprise.get_users_in_department(department_id)
    UserId = []
    UserName = []
    # print(len(res['userlist']))
    for x in range(len(res['userlist'])):
        # print(tourer)
        UserId.append(res['userlist'][x]['userid'])
        UserName.append(res['userlist'][x]['name'])

    weChatEnterprise_Push = WeChatEnterprise(corpsecret=CORPSECRET2)
    kwargs = {
        "articles": [
            {
                "title": "Qlik_报表推送",
                "description": "",
                "url": "https://sense-demo.qlik.com/single/?appid=06d53b15-2692-46ff-aaf6-651d7e1fa605&sheet=GahB",
                "picurl": "https://www.qlik.com/us/-/media/images/qlik/global/qlik-logo-2x.png?h=104&w=336&la=en&hash=39A8170194871041E4D0613C94693254962A941F"
            }

        ]
    }

    for x in range(len(UserId)):
        content = 'Hello！%s,以上您收到的是一份Qlik报表测试，请查阅！' % UserName[x]
        print(content)
        # weChatEnterprise_Push.send_msg_to_user(touser=[UserId[x]], content=content)
        weChatEnterprise_Push.send_msg_to_user(touser=[UserId[x]],content=content, msgtype = "news" ,**kwargs )



if __name__ == '__main__':

    from apscheduler.schedulers.blocking import BlockingScheduler
    import time

    sched = BlockingScheduler()

    while True:

        # sched.add_job(job, 'interval', seconds=10)
        sched.add_job(Qlik_Push_Excel, 'cron', day = 1 , hour = 10)     #定时每月1号十点出发

        try:
            sched.start()

        except:
            print('定时任务出错')
            time.sleep(10)
            continue

    # Qlik_Push_Excel()     #Excel导入用户信息推送
    # Qlik_Department()     #获取企业微信内部门信息推送