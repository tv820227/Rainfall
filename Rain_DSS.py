#%%
import os
import xlwt
import copy

# os.chdir('D:\\TV\\Hou-Cun Zun\\model\\SWMM model\\rainfall')
os.chdir('C:\\TV\\model\\SWMM model\\rainfall')

# duration
d = [1,2,3,6,12,24]

# 抓出Horner的參數
temp = []
with open('HornerCoefficient.txt') as coeff:
    for line in coeff:
        temp_l = [i for i in line.split()]
        temp.append(temp_l)
            
# period        
p = [int(i) for i in temp[0][1:]]
a = [float(i) for i in temp[1][1:]]
b = [float(i) for i in temp[2][1:]]
c = [float(i) for i in temp[3][1:]]


# 抓出長延時時間序列之比例
per_l = []
with open('LongTermPercentage.txt') as series:
    cnt = 0
    for line in series:
        if cnt < 1:
            cnt += 1
            continue
        else:
            per_l.append(float(line.split()[1]))

# 抓出短延時時間序列之比例
per_s = []
with open('ShortTermPercentage.txt') as series:
    cnt = 0
    for line in series:
        if cnt < 1:
            cnt += 1
            continue
        else:
            per_s.append(float(line.split()[1]))


rainfall = {}
# 計算短延時時間序列
for j in range(6): # j = index of duration
    for i in range(6): # i = index of period
        I = a[i]/(d[j]*60.0 + b[i])**c[i] # 計算總降雨強度
        V = I*d[j] # 計算單位時間內的降雨量
        raintype = str(d[j])+"hr_"+str(p[i])+"year"
        if d[j] < 6:
            rainfall[raintype] = [round(V*k/100.0,2) for k in per_s]
        else:
            rainfall[raintype] = [round(V*k/100.0,2) for k in per_l]
        

#%%
# create time series
time_series = []
time_series_s = {}
for h in d:
    total_time = h * 60
    if h > 3:
        n_interval = 24
    else:
        n_interval = 12
    time_step = total_time/n_interval
    
    time_temp = []
    for i in range(n_interval):
        time = i*time_step
        hour = "0"+str(int(time//60))
        minute = "0"+str(int(time%60))
        date_time = "01Jan2000 "+hour[-2:]+minute[-2:]
        time_temp.append(date_time)
        if date_time not in time_series:
            time_series.append(date_time)
    time_series_s[h] = sorted(time_temp)
time_series = sorted(time_series)


#%%
# 輸出計算之時間序列
book = xlwt.Workbook()
sheet1 = book.add_sheet("sheet1")
first_col = ["A","B","C","E","F","Units","Type"]

# 先輸出整個excel固定的格式
cnt = 0
for i in first_col:
    sheet1.write(cnt,0,i)
    cnt += 1

# output time series into excel
cnt = 0
for i in time_series:
    sheet1.write(7+cnt,0,cnt+1)
    sheet1.write(7+cnt,1,i)
    cnt += 1

#%%
# 輸出各個降雨的時間序列
# y = year, h = hour
cnt = 0
for h in range(6):
    for y in range(6):
        col = 2 + cnt

        # set the type of this rainfall
        sheet1.write(0,col,str(d[h])+"HR")
        sheet1.write(1,col,str(p[y])+"YEAR")
        sheet1.write(5,col,"mm")
        sheet1.write(6,col,"INST-VAL")
        
        # use key to find time series of this rainfall
        rain_key = str(d[h])+"hr_"+str(p[y])+"year"
        rain = rainfall[rain_key]
        last_row = 6

        # set the rainfall to specific time
        for i in range(len(rain)):
            row = 7+time_series.index(time_series_s[d[h]][i])
            
            # fill the rainfall equal to last number into blank cell
            for r in range(last_row+1,row+1,1):
                sheet1.write(r,col,rain[i])
            last_row = row
        
        # fill the rest of blank cells with 0
        for r in range(last_row+1,7+len(time_series)):
            sheet1.write(r,col,0)
        cnt += 1

book.save("C:\\TV\\model\\Hou_Cun_Zun1\\HMS_rainfall.xls")