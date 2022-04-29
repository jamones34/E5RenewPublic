# -*- coding: UTF-8 -*-
import os
import xlsxwriter
import requests as req
import json,sys,time,random

#reload(sys)
#sys.setdefaultencoding('utf-8')
emailaddress=os.getenv('EMAIL')
app_num=os.getenv('APP_NUM')
###########################
# config option description
# 0：off  ， 1：on
# allstart：Whether to open the call for all APIs, and close the default random extraction call. Default 0 off
# rounds: The number of rounds, that is, how many rounds are run at each start.
# rounds_delay: Whether to enable random delay between rounds, the last two parameters represent the delay interval. Default 0 off
# api_delay: Whether to enable the delay between APIs, the default is 0 to disable
# app_delay: Whether to open the delay between accounts, the default is 0 to close
########################################
config = {
         'allstart': 0,
         'rounds': 1,
         'rounds_delay': [0,0,5],
         'api_delay': [0,0,5],
         'app_delay': [0,0,5],
         }        
if app_num == '':
    app_num = '1'
city=os.getenv('CITY')
if city == '':
    city = 'Beijing'
access_token_list=['wangziyingwen']*int(app_num)

#Microsoft refresh_token acquisition
def getmstoken(ms_token,appnum):
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={'grant_type': 'refresh_token',
        'refresh_token': ms_token,
        'client_id':client_id,
        'client_secret':client_secret,
        'redirect_uri':'http://localhost:53682/'
        }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    if 'refresh_token' in jsontxt:
        print(r'Account/App '+str(appnum)+' Microsoft key obtained successfully')
    else:
        print(r'Account/App '+str(appnum)+' The Microsoft key acquisition failed\n'+'Please check whether the format and content of CLIENT_ID , CLIENT_SECRET , MS_TOKEN in secret are correct, and then reset')
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    return access_token

#api delay
def apiDelay():
    if config['api_delay'][0] == 1:
        time.sleep(random.randint(config['api_delay'][1],config['api_delay'][2]))
        
def apiReq(method,a,url,data='QAQ'):
    apiDelay()
    access_token=access_token_list[a-1]
    headers={
            'Authorization': 'bearer ' + access_token,
            'Content-Type': 'application/json'
            }
    if method == 'post':
        posttext=req.post(url,headers=headers,data=data)
    elif method == 'put':
        posttext=req.put(url,headers=headers,data=data)
    elif method == 'delete':
        posttext=req.delete(url,headers=headers)
    else :
        posttext=req.get(url,headers=headers)
    if posttext.status_code < 300:
        print('        Successful operation')
    else:
        print('        operation failed')
#    if posttext.status_code > 300:
#        print('        operation failed')
#        #No prompt for success
    return posttext.text
          

#Upload files to onedrive (less than 4M)
def UploadFile(a,filesname,f):
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/content'
    apiReq('put',a,url,f)
    
        
# Send mail to custom mailbox
def SendEmail(a,subject,content):
    url=r'https://graph.microsoft.com/v1.0/me/sendMail'
    mailmessage={'message': {'subject': subject,
                             'body': {'contentType': 'Text', 'content': content},
                             'toRecipients': [{'emailAddress': {'address': emailaddress}}],
                             },
                 'saveToSentItems': 'true'}            
    apiReq('post',a,url,json.dumps(mailmessage))	
	
#Modify excel (this function separation does not seem to make much sense)
#api-get itemid: https://graph.microsoft.com/v1.0/me/drive/root/search(q='.xlsx')?select=name,id,webUrl
def excelWrite(a,filesname,sheet):
    print('    add worksheet')
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/workbook/worksheets/add'
    data={
         "name": sheet
         }
    apiReq('post',a,url,json.dumps(data))
    print('    Add form')
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/workbook/worksheets/'+sheet+r'/tables/add'
    data={
         "address": "A1:D8",
         "hasHeaders": False
         }
    jsontxt=json.loads(apiReq('post',a,url,json.dumps(data)))
    print('    add line')
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/workbook/tables/'+jsontxt['id']+r'/rows/add'
    rowsvalues=[[0]*4]*2
    for v1 in range(0,2):
        for v2 in range(0,4):
            rowsvalues[v1][v2]=random.randint(1,1200)
    data={
         "values": rowsvalues
         }
    apiReq('post',a,url,json.dumps(data))
    
def taskWrite(a,taskname):
    print("    Create a task list")
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists'
    data={
         "displayName": taskname
         }
    listjson=json.loads(apiReq('post',a,url,json.dumps(data)))
    print("    Create a task")
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']+r'/tasks'
    data={
         "title": taskname,
         }
    taskjson=json.loads(apiReq('post',a,url,json.dumps(data)))
    print("    delete task")
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']+r'/tasks/'+taskjson['id']
    apiReq('delete',a,url)
    print("    delete task list")
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']
    apiReq('delete',a,url)    
    
def teamWrite(a,channelname):
    #newteam
    print('    newteam')
    url=r'https://graph.microsoft.com/v1.0/teams'
    data={
         "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
         "displayName": channelname,
         "description": "My Sample Team’s Description"
         }
    apiReq('post',a,url,json.dumps(data))
    print("    Get team information")
    url=r'https://graph.microsoft.com/v1.0/me/joinedTeams'
    teamlist = json.loads(apiReq('get',a,url))
    for teamcount in range(teamlist['@odata.count']):
        if teamlist['value'][teamcount]['displayName'] == channelname:
            #Create a channel
            print("    Create a team channel")
            data={
                 "displayName": channelname,
                 "description": "This channel is where we debate all future architecture plans",
                 "membershipType": "standard"
                 }
            url=r'https://graph.microsoft.com/v1.0/teams/'+teamlist['value'][teamcount]['id']+r'/channels'
            jsontxt = json.loads(apiReq('post',a,url,json.dumps(data)))
            url=r'https://graph.microsoft.com/v1.0/teams/'+teamlist['value'][teamcount]['id']+r'/channels/'+jsontxt['id']
            print("    delete team channel")
            apiReq('delete',a,url)
            #delete teams
            print("    delete team")
            url=r'https://graph.microsoft.com/v1.0/groups/'+teamlist['value'][teamcount]['id']
            apiReq('delete',a,url)  
            
def onenoteWrite(a,notename):
    print('    Create a notebook')
    url=r'https://graph.microsoft.com/v1.0/me/onenote/notebooks'
    data={
         "displayName": notename,
         }
    notetxt = json.loads(apiReq('post',a,url,json.dumps(data)))
    print('    Create notebook sections')
    url=r'https://graph.microsoft.com/v1.0/me/onenote/notebooks/'+notetxt['id']+r'/sections'
    data={
         "displayName": notename,
         }
    apiReq('post',a,url,json.dumps(data))
    print('    delete notebook')
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/Notebooks/'+notename
    apiReq('delete',a,url)
    
#Obtain access_token at one time to reduce the acquisition rate
for a in range(1, int(app_num)+1):
    client_id=os.getenv('CLIENT_ID_'+str(a))
    client_secret=os.getenv('CLIENT_SECRET_'+str(a))
    ms_token=os.getenv('MS_TOKEN_'+str(a))
    access_token_list[a-1]=getmstoken(ms_token,a)
print('')    
#get weather
headers={'Accept-Language': 'zh-CN'}
weather=req.get(r'http://wttr.in/'+city+r'?format=4&?m',headers=headers).text

#Actual operation
for a in range(1, int(app_num)+1):
    print('account '+str(a))
    print('Send mail (mailbox runs alone, only send once per run, to prevent bans)')
    if emailaddress != '':
        SendEmail(a,'weather',weather)
print('')
#other APIs
for _ in range(1,config['rounds']+1):
    if config['rounds_delay'][0] == 1:
        time.sleep(random.randint(config['rounds_delay'][1],config['rounds_delay'][2]))     
    print('Round '+str(_)+' \n')        
    for a in range(1, int(app_num)+1):
        if config['app_delay'][0] == 1:
            time.sleep(random.randint(config['app_delay'][1],config['app_delay'][2]))        
        print('account '+str(a))    
        #Generate random names
        filesname='QAQ'+str(random.randint(1,600))+r'.xlsx'
        #Create a new random xlsx file
        xls = xlsxwriter.Workbook(filesname)
        xlssheet = xls.add_worksheet()
        for s1 in range(0,4):
            for s2 in range(0,4):
                xlssheet.write(s1,s2,str(random.randint(1,600)))
        xls.close()
        xlspath=sys.path[0]+r'/'+filesname
        print('Upload files (may occasionally fail to create uploads)')
        with open(xlspath,'rb') as f:
            UploadFile(a,filesname,f)
        choosenum = random.sample(range(1, 5),2)
        if config['allstart'] == 1 or 1 in choosenum:
            print('Excel file operations')
            excelWrite(a,filesname,'QVQ'+str(random.randint(1,600)))
        if config['allstart'] == 1 or 2 in choosenum:
            print('team operation')
            teamWrite(a,'QVQ'+str(random.randint(1,600)))
        if config['allstart'] == 1 or 3 in choosenum:
            print('task operation')
            taskWrite(a,'QVQ'+str(random.randint(1,600)))
        if config['allstart'] == 1 or 4 in choosenum:
            print('onenote operation')
            onenoteWrite(a,'QVQ'+str(random.randint(1,600)))
        print('-')
