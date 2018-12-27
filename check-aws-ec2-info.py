#### provided by Jenny
import boto3
import xlwt
import sys
import json
import xlrd
from xlutils.copy import copy
from datetime import datetime

ProName = 'training'
AmiID = []
InsID = []
InsType = []
Name = []
AvaZone = []
PriIP = []
Status = []
SgLIST = []
SgGRPS = {}
AllSgs = []

def json_serial(obj):
    """JSON serializer for objects not serializable by default json code"""

    if isinstance(obj, datetime):
        serial = obj.isoformat()
        return serial
    raise TypeError ("Type not serializable")

def get_InstanceID():
    b = boto3.session.Session(profile_name=ProName)
    ec2 = b.resource('ec2')
    #ec2 = boto3.resource('ec2')
    global idlist
    idlist = []
    for status in ec2.meta.client.describe_instance_status()['InstanceStatuses']:
        idlist.append(status['InstanceId'])

def get_data():
    b = boto3.session.Session(profile_name=ProName)
    ec2 = b.client('ec2')
    #ec2 = boto3.client('ec2')
    i = 0
    Data = {}
    for id in idlist:
        i += 1
        Data = ec2.describe_instances(InstanceIds=[id])
        #f = open("test.jsonfile","w+")
        #print(Data.keys())
        RData = Data['Reservations']
        #print(RData)
        BaseData = RData[0]['Instances'][0]
        #print(len(BaseData))
        #print(BaseData)
        AmiID.append(BaseData['ImageId'])
        InsID.append(BaseData['InstanceId'])
        #print(len(InsID))
        InsType.append(BaseData['InstanceType'])
        AvaZone.append(BaseData['Placement']['AvailabilityZone'])
        PriIP.append(BaseData['PrivateIpAddress'])
        Status.append(BaseData['State']['Name'])
        #SgGRPS.append(BaseData['NetworkInterfaces']['Groups'])
        SgLIST.append(BaseData['SecurityGroups'])
        try:
            tags = BaseData['Tags']
        except KeyError:
            pass
            
        for i in range(1,len(tags)+1):
            tagname = tags[i-1]['Key']
            if tagname == 'Name':
                Name.append(tags[i-1]['Value'])
        print(len(Name))
        #tags = BaseData['Tags'][3]['Value']

        
        #print(Data.Reservaitions.Instances.ImageId)
        #Data = json.dumps(Data,default=json_serial)
        #print(Data)
        #f.write(Data)
        #f.close()

def write_execl():
    testexecl = xlwt.Workbook()
    sheet1 = testexecl.add_sheet('server_list',cell_overwrite_ok=True)
    sheet1.write(0,0,'seq')
    sheet1.write(0,1,'Name')
    sheet1.write(0,2,'InstanceId')
    sheet1.write(0,3,'InstanceType')
    sheet1.write(0,4,'AvailabilityZone')
    sheet1.write(0,5,'ImageId')
    sheet1.write(0,6,'PrivateIpAddress')
    sheet1.write(0,7,'State')
    sheet1.write(0,8,'SecurityGroups')
    testexecl.save('ServerList.xls')

def write_data(num,data):
    x = 0
    old_excel = xlrd.open_workbook('ServerList.xls', formatting_info=True)
    new_excel = copy(old_excel)
    sheet2 = new_excel.get_sheet(0)
    #print(len(idlist))
    for i in range(1,len(idlist)+1):
        sheet2.write(i,0,i)
    for j in range(1,len(data)+1):

        ls = data[j-1]
        x += 1

        #sheet1.write(j,0,j)
        sheet2.write(j,num,ls)
    new_excel.save('ServerList.xls')

get_InstanceID()
get_data()
write_execl()
#print(AmiID,InsID,InsType,AvaZone,PriIP)
write_data(1,Name)
write_data(2,InsID)
write_data(3,InsType)
write_data(4,AvaZone)
write_data(5,AmiID)
write_data(6,PriIP)
write_data(7,Status)

#print(type(len(SgLIST[1])))
for idnum in range(0,len(idlist)):
    #print(len(SgLIST[idnum]))
    SGPerIns = []
    for groupnum in range(0,len(SgLIST[idnum])):
        SGPerIns.append(SgLIST[idnum][groupnum]['GroupName'])
        SGPerIns.append(', ')
        #print(idnum, SGPerIns)
        #print(SgLIST[idnum][groupnum]['GroupName'])
    AllSgs.append(SGPerIns)


write_data(8,AllSgs)



