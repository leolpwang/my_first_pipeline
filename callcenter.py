import pandas
import datetime 
n = 0
cc_list = []
active_cc_list = []
date = input('input the date(yyyy/mm/dd): ')
year = int(date.split('/')[0])
month = int(date.split('/')[1])
day = int(date.split('/')[2])
d1 = datetime.datetime(year,month,day) 
data_df = pandas.read_excel('CSKS4.xlsx')
items = data_df.to_dict('records')

#for item in items:
#    cc_list.append(item['Cost Center'])
#cc_list = list(set(cc_list))
#print ("There are {} cost centers in total on the list.".format(len(cc_list))) 

for item in items:
    if d1 <= item['Valid To']:
        if d1 > item['Valid From']:
                if item['Actual primary costs'] != 'X':
                    n = n+1
                    active_cc_list.append(item)
df = pandas.DataFrame.from_dict(active_cc_list)
df.to_excel('active_cc_list_' + str(year) + str(month) + str(day) + '.xlsx')

print ("There are {} active call centers in total on {}/{}/{}.".format(n,year,month,day)) 
