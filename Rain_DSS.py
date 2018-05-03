import xlwt
import os

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


# 計算短延時時間序列
for j in range(6):
    for i in range(6):
        I = a[i]/(d[j]*60.0 + b[i])**c[i]
        V = I*d[j]
        if d[j] < 6:
#             rainfall = [round(V*k/100.0,2) for k in per_l]
            rainfall = [round(V*k/100.0,2) for k in per_s]
        else:
            rainfall = [round(V*k/100.0,2) for k in per_l]
        
        # 輸出計算之時間序列
        out = open('.\\Dat\\%ihr_%iyear.dat'%(d[j],p[i]),'w')
        out.write(';Rainfall Data for Gage G1\n')
        
        time_step = int(d[j]*60/len(rainfall))
        h = 0
        m = 0
        for l in rainfall:
            # 判斷時間是否需要進位
            if m >= 60:
                h += 1
                m -= 60
            
            # 判斷小時是不是兩個數字，如果只有一個在前面補0
            if len(str(h)) < 2:
                out_h = '0' + str(h)
            else:
                out_h = str(h)
            
            # 判斷分鐘是不是兩個數字，如果只有一個在前面補0
            if len(str(m)) < 2:
                out_m = '0' + str(m)
            else:
                out_m = str(m)
            
            # 欲輸出之時間格式
            time_series = out_h+':'+out_m
            out.write(time_series+'\t'+str(l)+'\n')
            m += time_step
        out.close()