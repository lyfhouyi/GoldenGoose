import os
import time

import pandas as pd
import xlwt
from selenium import webdriver


# 获取数据
def getCsv():
    tmp_path=os.getcwd()
    prefs={'profile.default_content_settings.popups':0,'download.default_directory':tmp_path}
    option=webdriver.ChromeOptions()
    option.add_experimental_option('prefs',prefs)
    option.add_argument('disable-infobars')
    driver=webdriver.Chrome(options=option)
    driver.get('https://www.lixinger.com/analytics/index/dashboard/value#')

    # 登录
    driver.find_element_by_xpath('/html/body/div[7]/div[1]/div/div/div/div[2]/div/div/form/div[1]/input').send_keys('18910629881')
    driver.find_element_by_xpath('/html/body/div[7]/div[1]/div/div/div/div[2]/div/div/form/div[2]/input').send_keys('LYF@houyi')
    driver.find_element_by_xpath('/html/body/div[7]/div[1]/div/div/div/div[2]/div/div/form/div[3]/div[2]/button').click()
    time.sleep(5)

    # 导出 csv 数据
    driver.find_element_by_xpath('/html/body/div[1]/div[1]/div/div[2]/div[2]/div[1]/button').click()
    print('获取数据中...')
    time.sleep(5)
    driver.quit()
    print('已获取 csv 文件')



def processCsv():
    alignment=xlwt.Alignment()
    alignment.horz=0x01
    alignment.vert=0x01

    font_head=xlwt.Font()
    font_head.bold=True
    font_head.height=20*15

    font_partion=xlwt.Font()
    font_partion.bold=False
    font_partion.height=20*15
    pattern_partion=xlwt.Pattern()
    pattern_partion.pattern=xlwt.Pattern.SOLID_PATTERN
    pattern_partion.pattern_fore_colour=5 # 黄色

    font_value=xlwt.Font()
    font_value.bold=False
    font_value.height=20*15

    style_head=xlwt.XFStyle()
    style_head.font=font_head
    style_head.alignment=alignment

    style_partion=xlwt.XFStyle()
    style_partion.font=font_partion
    style_partion.pattern=pattern_partion
    style_partion.alignment=alignment

    style_value=xlwt.XFStyle()
    style_value.font=font_value
    style_value.alignment=alignment

    head=['指数代码','指数名称','发布时间','市盈率PE','市盈率温度','市净率PB','市净率温度','指数温度']
    indexsNo_1=['000016','000905','000300','399006'] # 宽基指数
    indexsNo_2=['000922','000015','000925'] # 策略指数
    indexsNo_3=['HSI','.INX','HSCEI'] # 境外指数

    fileExcel=xlwt.Workbook(encoding='ascii')
    ws=fileExcel.add_sheet('指数温度')
    ws.col(1).width=16*256
    # 写首行
    lineLoc=0
    for i in range(len(head)):
        ws.write(lineLoc,i,head[i],style_head)
        ws.col(i).width=22*256
    ws.col(1).width=30*256

    filesList=[]
    for root,dirs,files in os.walk(os.getcwd()):
        for file in files:
            if file.split('.')[-1]=='csv':
                filesList.append(file)
    csvFile=pd.read_csv(filesList[0],usecols=[0,1,2,5,6,12,13])
    
    # 宽基指数
    lineLoc+=1
    ws.write(lineLoc,0,'宽基指数',style_partion)

    lineLoc+=1
    for line in csvFile.iloc:
        lineList=list(line)
        lineList[0]=lineList[0][2:-1]
        lineList[4]=lineList[4]*100
        lineList[6]=lineList[6]*100
        lineList.append(float('%.2f'%((lineList[4]+lineList[6])/2)))
        if lineList[0] in indexsNo_1:
            for item,j in zip(lineList,range(len(lineList))):
                ws.write(lineLoc,j,item,style_value)
            lineLoc+=1
    
    # 策略指数
    ws.write(lineLoc,0,'策略指数',style_partion)

    lineLoc+=1
    for line in csvFile.iloc:
        lineList=list(line)
        lineList[0]=lineList[0][2:-1]
        lineList[4]=lineList[4]*100
        lineList[6]=lineList[6]*100
        lineList.append(float('%.2f'%((lineList[4]+lineList[6])/2)))
        if lineList[0] in indexsNo_2:
            for item,j in zip(lineList,range(len(lineList))):
                ws.write(lineLoc,j,item,style_value)
            lineLoc+=1

    # 境外指数
    ws.write(lineLoc,0,'境外指数',style_partion)

    lineLoc+=1
    for line in csvFile.iloc:
        lineList=list(line)
        lineList[0]=lineList[0][2:-1]
        lineList[4]=lineList[4]*100
        lineList[6]=lineList[6]*100
        lineList.append(float('%.2f'%((lineList[4]+lineList[6])/2)))
        if lineList[0] in indexsNo_3:
            for item,j in zip(lineList,range(len(lineList))):
                ws.write(lineLoc,j,item,style_value)
            lineLoc+=1

    fileExcel.save(os.getcwd() + r'\\指数温度\\%s.xls'%filesList[0].split('_')[-2])
    os.remove(os.path.join(os.getcwd(),filesList[0]))
    print('csv 文件分析完毕')



if __name__ == '__main__':
    getCsv()
    processCsv()
    