import unittest, time, requests, webbrowser, datetime, os, re, xlrd, openpyxl
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.alert import Alert


#spreadSheetPath = 'C:\\Users\\tomoa\\Workspace\\buyma\\spreadsheets\\'
spreadSheetPath = 'spreadsheets\\'
photoPath = ''

#エクセル指定
excelName = ("anthropologie")
xl_bk = xlrd.open_workbook(spreadSheetPath + excelName + '.xlsx')
xl_sh = xl_bk.sheet_by_name(excelName)
wb = openpyxl.load_workbook(spreadSheetPath + excelName + '.xlsx') 
ws = wb.active



#列指定
yoko = input('出品番号:')
yoko = int(yoko) - 1
folderNum = int(yoko) + 1

    

###　ログイン　###
#################
browser = webdriver.Chrome()
browser.get("https://www.buyma.com/my/itemedit/")
email = browser.find_element_by_id('txtLoginId')
email.send_keys('taira4420@gmail.com')
password = browser.find_element_by_id('txtLoginPass')
password.send_keys('taira442054')
browser.find_element_by_id('login_do').click()


while True:
    print(folderNum)
    browser.implicitly_wait(40)
    time.sleep(1)

    ###J列読み込み###
    ################
    action = xl_sh.cell_value(yoko,5)



    
    ###########################
    ### もし全サイズなかったら ###
    ###########################

    if action == 'allNo' or action == 'change':
        ITEMURL = xl_sh.cell_value(yoko,10)
        browser.get(ITEMURL)
        browser.implicitly_wait(40)
        time.sleep(1)
        browser.find_element_by_xpath('//*[@id="stop"]').click()
        browser.implicitly_wait(40)
        time.sleep(1)
        browser.find_element_by_xpath('//*[@id="stop"]').click()
        browser.implicitly_wait(40)
        time.sleep(1)
        browser.get("https://www.buyma.com/my/itemedit/")
        browser.implicitly_wait(40)
        time.sleep(1)
        yoko += 1
        folderNum += 1
        continue


    #################
    ### サイズ変更 ###
    #################
        
    elif action == 'yes':
        ITEMURL = xl_sh.cell_value(yoko,10)
        browser.get(ITEMURL)

        #設定するボタンをクリック
        browser.find_element_by_xpath('//*[@id="chkForm"]/div[2]/table[1]/tbody/tr[9]/td/div[1]/div[1]/a').click()
        browser.implicitly_wait(40)
        time.sleep(1)

        #サイズボックスをサイズボックスをすべて消す
        box = 0
        while True:
            if box == 10:
                break
            browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[2]/td[2]/div/span/i').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            box += 1

        
        #サイズをリストとしてスプリットさせる
        changeSize = xl_sh.cell_value(yoko,8)
        changeSize = changeSize.split(',')


        #サイズ分だけサイズボックスを追加
        len_newSize = len(changeSize)
        count = 1
        while True:
            if count == len_newSize:
                break
            browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/div[4]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            count += 1
        

        #サイズを反映
        sizeNumber = 2
        for i in changeSize:

            #サイズを反映
            browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[2]/input').send_keys(i)
            browser.implicitly_wait(40)
            time.sleep(1)
            
            #日本参考サイズ(洋服)
            category = xl_sh.cell_value(yoko,2)
            if category == 'メンズファッション 靴・ブーツ・サンダル 靴・ブーツ・サンダルその他' or category == 'メンズファッション 靴・ブーツ・サンダル スニーカー' or category == 'メンズファッション 靴・ブーツ・サンダル サンダル':
                if i == '4' or i == '4.5' or i == '5':
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '5.5' or i == '35 1/2':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("23.5cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '6' or i == '36':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("24cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '6.5' or i == '36 1/2':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("24.5cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '7' or i == '37' or i == 'M 7/W 8.5' or i == 'US 7/EU 40':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("25cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '7.5' or i == '37 1/2' or i == 'M 7.5/W 9':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("25.5cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '8' or i == '38' or i == 'M 8/W 9.5' or i == 'US 8/EU 41':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("26cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '8.5' or i == '38 1/2' or i == 'M 8.5/W 10':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("26.5cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '9' or i == '39' or i == 'M 9/W 10.5' or i == 'US 9/EU 42':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("27cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '9.5' or i == '39 1/2' or i == 'M 9.5/W 11' or i == 'US 9.5/EU 43':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("27.5cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '10' or i == '40' or i == 'M 10/W 11.5' or i == 'US 10/EU 44':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("28cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '10.5' or i == '40 1/2' or i == 'M 10.5/W 12':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("28.5cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '11' or i == '41' or i == 'M 11/W 12.5' or i == 'US 11/EU 45':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("29cm以上")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == 'M 11.5/W 13':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("29cm以上")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == 'M 12/W 13.5' or i == 'US 12/EU 46' or i == '12':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("29cm以上")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == 'M 12.5/W 14' or i == '12.5':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("29cm以上")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == 'M 13/W 14.5' or i == '13':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("29cm以上")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == 'M 13.5/W 15' or i == '13.5':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("29cm以上")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == 'M 14/W 15.5' or i == '14':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("29cm以上")
                    browser.implicitly_wait(40)
                    time.sleep(1)

            
            elif i == 'XXS' or i == '2' or i == '24':
                browser.implicitly_wait(40)
                time.sleep(1)
            elif i == 'XS' or i == '4' or i == '25' or i == 'XS/S' or i == '30/32':
                browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("S")
                browser.implicitly_wait(40)
                time.sleep(1)
            elif i == 'S' or i == '6' or i == '26' or i == 'S/M' or i == '32/32':
                browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("M")
                browser.implicitly_wait(40)
                time.sleep(1)
            elif i == 'M' or i == '8' or i == '27' or i == 'M/L' or i == '34/32':
                browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("L")
                browser.implicitly_wait(40)
                time.sleep(1)
            elif i == 'L' or i == 'XL' or i == 'XXL' or i == 'L/XL' or i == '10' or i == '12' or i == '14' or i == '28' or i == '29' or i == '30' or i == '31' or i == '32' or i == '36/32':
                browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("XL以上")
                browser.implicitly_wait(40)
                time.sleep(1)
            elif i == 'ONE SIZE':
                browser.find_element_by_xpath(' //*[@id="rdoSelectSize2"]').click()
                browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("ONE SIZE")
                break
            else:
                browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys(i)
                browser.implicitly_wait(40)
                time.sleep(1)
            
            sizeNumber += 1


        #全サイズ反映後、設定ボタンをクリック
        browser.find_element_by_xpath('//*[@id="components"]/div/div[2]/a[2]').click()
        browser.implicitly_wait(40)
        time.sleep(1)


        #購入期限を今日から90日延ばす
        twoweeks = browser.find_element_by_xpath('//*[@id="itemedit_yukodate"]')
        browser.implicitly_wait(40)
        time.sleep(1)
        today = datetime.date.today()
        changeDate = datetime.timedelta(days=90)
        changeDate = today + changeDate
        changeDate = str(changeDate)
        changeDate = changeDate.replace('-','/')
        twoweeks.clear()
        browser.implicitly_wait(40)
        time.sleep(1)
        twoweeks.send_keys(changeDate)
        browser.implicitly_wait(40)
        time.sleep(1)


        #入力内容を保存する
        browser.find_element_by_xpath('//*[@id="confirmButton"]').click()
        browser.implicitly_wait(40)
        time.sleep(1)
        browser.get("https://www.buyma.com/my/itemedit/")
        yoko += 1
        folderNum += 1
        continue



        
    #################
    ### 新商品出品 ###
    #################
    
    elif action == 'new':
    
    

    
        #####　カテゴリ　####
        ####################
        category = xl_sh.cell_value(yoko,2)
    
        #検索ボタンをクリック
        kesaku = browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/form/div[3]/div[2]/table/tbody/tr[1]/td/table/tbody/tr/td[1]/a')
        time.sleep(3)
        kesaku.click()
        browser.implicitly_wait(40)
        time.sleep(1)
    
        #カテゴリ分け
        if category == 'メンズファッション アウター・ジャケット ジャケットその他':
            browser.find_element_by_link_text('アウター・ジャケット').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('アウターその他').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション トップス シャツ':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[1]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('シャツ').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション ファッション雑貨・小物 ベルト':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[11]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ベルト').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            
        elif category == 'メンズファッション 水着・ビーチグッズ 水着':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[14]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('水着').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション バッグ・カバン バッグ・カバンその他':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[5]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('バッグ・カバンその他').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション バッグ・カバン ボストンバッグ':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[5]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ボストンバッグ').click()
            browser.implicitly_wait(40)
            time.sleep(1)

        elif category == 'メンズファッション トップス Tシャツ・カットソー':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[1]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('Tシャツ・カットソー').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション トップス ニット・セーター':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[1]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ニット・セーター').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション トップス パーカー・フーディ':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[1]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('パーカー・フーディ').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション ボトムス ボトムスその他':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[2]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ボトムスその他').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション ボトムス デニム・ジーパン':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[2]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('デニム・ジーパン').click()
            browser.implicitly_wait(40)
            time.sleep(1)

        elif category == 'メンズファッション 靴・ブーツ・サンダル スニーカー':
            browser.find_element_by_link_text('靴・ブーツ・サンダル').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('スニーカー').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション 靴・ブーツ・サンダル ブーツ':
            browser.find_element_by_link_text('靴・ブーツ・サンダル').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ブーツ').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション 靴・ブーツ・サンダル サンダル':
            browser.find_element_by_link_text('靴・ブーツ・サンダル').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('サンダル').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション 帽子 ニットキャップ・ビーニー':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[10]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ニットキャップ・ビーニー').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション 帽子 キャップ':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[10]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('キャップ').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション アイウェア サングラス':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[9]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('サングラス').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション バッグ・カバン バックパック・リュック':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[5]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('バックパック・リュック').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション バッグ・カバン ショルダーバッグ':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[5]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ショルダーバッグ').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション インナー・ルームウェア アンダーシャツ・インナー':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[13]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('アンダーシャツ・インナー').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション インナー・ルームウェア トランクス':
            browser.find_element_by_xpath('//*[@id="my"]/div[8]/div[2]/div/div[1]/dl[2]/dd/ul/li[13]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('トランクス').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif 'レディース' in category and 'T' in category:
            browser.find_element_by_link_text('トップス').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('Tシャツ・カットソー').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif 'レディース' in category and 'スカート' in category:
            browser.find_element_by_link_text('ボトムス').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('スカート').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'レディースファッション トップス パーカー・フーディ':
            browser.find_element_by_link_text('トップス').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('パーカー・フーディ').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'レディースファッション 靴・シューズ スニーカー':
            browser.find_element_by_link_text('靴・シューズ').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('スニーカー').click()
            browser.implicitly_wait(40)
            time.sleep(1)

        elif category == 'レディースファッション ファッション雑貨・小物 マフラー':
            browser.find_element_by_link_text('ファッション雑貨・小物').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('マフラー').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'レディースファッション バッグ・カバン ショルダーバッグ・ポシェット':
            browser.find_element_by_link_text('バッグ・カバン').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ショルダーバッグ・ポシェット').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'レディースファッション バッグ・カバン バックパック・リュック':
            browser.find_element_by_link_text('バッグ・カバン').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('バックパック・リュック').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'レディースファッション 帽子 ニットキャップ・ビーニー':
            browser.find_element_by_link_text('帽子').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ニットキャップ・ビーニー').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'レディースファッション ファッション雑貨・小物 手袋':
            browser.find_element_by_link_text('ファッション雑貨・小物').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('手袋').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'レディースファッション ボトムス ボトムスその他':
            browser.find_element_by_link_text('ボトムス').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ボトムスその他').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'レディースファッション ボトムス デニム・ジーパン':
            browser.find_element_by_link_text('ボトムス').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('デニム・ジーパン').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif 'レディース' in category and 'ドレス' in category:
            browser.find_element_by_link_text('ワンピース・オールインワン').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ワンピース').click()
            browser.implicitly_wait(40)
            time.sleep(1)

        elif category == 'レディースファッション インナー・ルームウェア ルームウェア・パジャマ':
            browser.find_element_by_link_text('インナー・ルームウェア').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ルームウェア・パジャマ').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'レディースファッション 帽子 キャップ':
            browser.find_element_by_link_text('帽子').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('キャップ').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'レディースファッション アウター コート':
            browser.find_element_by_link_text('アウター').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('コート').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif 'レディース' in category and 'ジャケット' in category:
            browser.find_element_by_link_text('アウター').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ダウンジャケット・コート').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif 'レディース' in category and 'スウェット' in category:
            browser.find_element_by_link_text('トップス').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('スウェット・トレーナー').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'メンズファッション アウター・ジャケット ダウンジャケット':
            browser.find_element_by_link_text('アウター・ジャケット').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ダウンジャケット').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif 'レディース' in category and 'シャツ' in category:
            browser.find_element_by_link_text('トップス').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ブラウス・シャツ').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'レディースファッション アウター アウターその他':
            browser.find_element_by_link_text('アウター').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('アウターその他').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'レディースファッション バッグ・カバン バッグ・カバンその他':
            browser.find_element_by_link_text('バッグ・カバン').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('バッグ・カバンその他').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif 'レディース' in category and 'セーター' in category:
            browser.find_element_by_link_text('トップス').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ニット・セーター').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'ベビー・キッズ 子供服・ファッション用品(85cm～) キッズアウター':
            browser.find_element_by_link_text('子供服・ファッション用品(85cm～)').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('キッズアウター').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'ベビー・キッズ 子供服・ファッション用品(85cm～) キッズ用トップス':
            browser.find_element_by_link_text('子供服・ファッション用品(85cm～)').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('キッズ用トップス').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'ベビー・キッズ 子供服・ファッション用品(85cm～) キッズ用ボトムス':
            browser.find_element_by_link_text('子供服・ファッション用品(85cm～)').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('キッズ用ボトムス').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'ベビー・キッズ 子供服・ファッション用品(85cm～) 子供用パジャマ・ルームウェア・スリーパー':
            browser.find_element_by_link_text('子供服・ファッション用品(85cm～)').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('子供用パジャマ・ルームウェア・スリーパー').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'ベビー・キッズ 子供服・ファッション用品(85cm～) 子供用帽子・手袋・ファッション小物':
            browser.find_element_by_link_text('子供服・ファッション用品(85cm～)').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('子供用帽子・手袋・ファッション小物').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif category == 'ベビー・キッズ 子供服・ファッション用品(85cm～) 子供用リュック・バックパック':
            browser.find_element_by_link_text('子供服・ファッション用品(85cm～)').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('子供用リュック・バックパック').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            




        ###ブランド名###
        ###############

        A = '//*[@id="alpha"]/table/tbody/tr[1]/td[1]/a/img'
        B = '//*[@id="alpha"]/table/tbody/tr[1]/td[2]/a/img'
        C = '//*[@id="alpha"]/table/tbody/tr[1]/td[3]/a/img'
        D = '//*[@id="alpha"]/table/tbody/tr[1]/td[4]/a/img'
        E = '//*[@id="alpha"]/table/tbody/tr[1]/td[5]/a/img'
        F = '//*[@id="alpha"]/table/tbody/tr[1]/td[7]/a/img'
        G = '//*[@id="alpha"]/table/tbody/tr[1]/td[8]/a/img'
        H = '//*[@id="alpha"]/table/tbody/tr[1]/td[9]/a/img'
        I = '//*[@id="alpha"]/table/tbody/tr[1]/td[10]/a/img'
        J = '//*[@id="alpha"]/table/tbody/tr[1]/td[11]/a/img'
        K = '//*[@id="alpha"]/table/tbody/tr[2]/td[1]/a/img'
        L = '//*[@id="alpha"]/table/tbody/tr[2]/td[2]/a/img'
        M = '//*[@id="alpha"]/table/tbody/tr[2]/td[3]/a/img'
        N = '//*[@id="alpha"]/table/tbody/tr[2]/td[4]/a/img'
        O = '//*[@id="alpha"]/table/tbody/tr[2]/td[5]/a/img'
        P = '//*[@id="alpha"]/table/tbody/tr[2]/td[7]/a/img'
        Q = '//*[@id="alpha"]/table/tbody/tr[2]/td[8]/a/img'
        R = '//*[@id="alpha"]/table/tbody/tr[2]/td[9]/a/img'
        S = '//*[@id="alpha"]/table/tbody/tr[2]/td[10]/a/img'
        T = '//*[@id="alpha"]/table/tbody/tr[2]/td[11]/a/img'
        U = '//*[@id="alpha"]/table/tbody/tr[3]/td[1]/a/img'
        V = '//*[@id="alpha"]/table/tbody/tr[3]/td[2]/a/img'
        W = '//*[@id="alpha"]/table/tbody/tr[3]/td[3]/a/img'
        X = '//*[@id="alpha"]/table/tbody/tr[3]/td[4]/a/img'
        Y = '//*[@id="alpha"]/table/tbody/tr[3]/td[5]/a/img'
        Z = '//*[@id="alpha"]/table/tbody/tr[3]/td[7]/a/img'

        brand = xl_sh.cell_value(yoko,38) 

        brandURL1 = browser.find_element_by_xpath('//*[@id="tbl_latest_syupin_brand"]/tbody/tr[2]/td').text
        brandURL2 = browser.find_element_by_xpath('//*[@id="tbl_latest_syupin_brand"]/tbody/tr[3]/td').text
        brandURL3 = browser.find_element_by_xpath('//*[@id="tbl_latest_syupin_brand"]/tbody/tr[4]/td').text
        
        browser.implicitly_wait(40)
        time.sleep(1)

        if brand in brandURL1:
            browser.find_element_by_xpath('//*[@id="tbl_latest_syupin_brand"]/tbody/tr[2]/td/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand in brandURL2:
            browser.find_element_by_xpath('//*[@id="tbl_latest_syupin_brand"]/tbody/tr[3]/td/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand in brandURL3:
            browser.find_element_by_xpath('//*[@id="tbl_latest_syupin_brand"]/tbody/tr[4]/td/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)


            

        elif brand == "Free People(フリーピープル)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath(F).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == "Levi's(リーバイス)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(L).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == "BIRKENSTOCK(ビルケンシュトック)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath(B).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'PUMA(プーマ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(P).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == 'VANS(バンズ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath(V).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == 'New Balance(ニューバランス)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath(N).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)                    
        elif brand == 'adidas(アディダス)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(A).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == 'Nike(ナイキ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(N).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == 'asics(アシックス)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(A).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)

        elif brand == 'CONVERSE(コンバース)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(C).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)

        elif brand == 'BIRKENSTOCK(ビルケンシュトック)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(B).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'CAMPER(カンペール)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(C).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Reebok(リーボック)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(R).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == 'New Balance(ニューバランス)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(N).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)

        elif brand == 'VANS(バンズ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(V).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)

        elif brand == 'SKECHERS(スケッチャーズ)': 
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(S).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Timberland(ティンバーランド)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(T).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'SUPERGA(スペルガ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(S).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Jeffrey Campbell(ジェフリーキャンベル)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(J).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'ROCKET DOG(ロケットドッグ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(R).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'KATIN(ケイティン)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(K).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Dickies(ディッキーズ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(D).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Guess(ゲス)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(G).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Herschel Supply(ハーシェルサプライ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(H).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'KAPTEN & SON(キャプテン＆サン)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(K).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'UMBRO(アンブロ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(U).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'THE NORTH FACE(ザノースフェイス)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(T).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")
            browser.implicitly_wait(40)
        elif brand == 'Patagonia(パタゴニア)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(P).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == 'PUMA(プーマ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(P).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)          
        elif brand == 'ALPHA INDUSTRIES(アルファ　インダストリーズ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(A).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text('ALPHA INDUSTRIES(アルファ　インダストリーズ)').click()
        elif brand == 'Urban Outfitters(アーバンアウトフィッターズ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(U).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'CHAMPION(チャンピオン)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(C).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'STUSSY(ステューシー)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(S).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == "Levi's(リーバイス)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(L).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == "FILA(フィラ)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(F).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == "LAUREN RALPH LAUREN(ローレンラルフローレン)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(L).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == "Calvin Klein(カルバンクライン)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(C).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)


        elif brand == "Tommy Hilfiger(トミーヒルフィガー)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(T).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == "HUF(ハフ)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(H).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == "Teva(テバ)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(T).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == "NAUTICA(ノーティカ)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(N).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == "Fossil(フォッシル)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(F).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'DIESEL(ディーゼル)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(D).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == 'UNDER ARMOUR(アンダーアーマー)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(U).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == "ALDO(アルド)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(A).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Cole Haan(コールハーン)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(C).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")
            browser.implicitly_wait(40)
        elif brand == "TOMS(トムス)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(T).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Tory Burch(トリーバーチ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(T).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)

        elif brand == "MANGO(マンゴ)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(M).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Alexander Wang(アレキサンダーワン)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(A).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            if 'DIEGO' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            elif 'EMILE' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[2]/a').click()
            elif 'ROCKIE' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[5]/a').click()
            else:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == "COMME des GARCONS(コムデギャルソン)":
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(C).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'kate spade new york(ケイトスペード)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(K).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            if '2 park' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            elif 'broome' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[2]/a').click()
            elif 'catherine' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[3]/a').click()
            elif 'cedar' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[4]/a').click()
            elif 'cobble' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[5]/a').click()
            elif 'magnolia' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[6]/a').click()
            elif 'madison' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[7]/a').click()
            elif 'newbury' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[8]/a').click()
            elif 'southport' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[9]/a').click()
            elif 'wellesley' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[10]/a').click()
            elif 'west valley' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[11]/a').click()
            else:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == 'Michael Kors(マイケルコース)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(M).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            if 'Adele' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            elif 'Anabelle' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[2]/a').click()
            elif 'Angelina' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[3]/a').click()
            elif 'Astor' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[4]/a').click()
            elif 'Ava' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[5]/a').click()
            elif 'Bancroft' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[6]/a').click()
            elif 'Barbara' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[7]/a').click()
            elif 'Bedford' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[8]/a').click()
            elif 'Billy' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[9]/a').click()
            elif 'Blakely' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[10]/a').click()
            elif 'Bridgette' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[11]/a').click()
            elif 'Bristol' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[12]/a').click()
            elif 'Blooklyn' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[13]/a').click()
            elif 'Bryant' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[14]/a').click()
            elif 'Cali Tie Dye' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[15]/a').click()
            elif 'Casey' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[16]/a').click()
            elif 'Cate' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[17]/a').click()
            elif 'Chatham' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[18]/a').click()
            elif 'Cindy' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[19]/a').click()
            elif 'Cori' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[20]/a').click()
            elif 'Cynthia' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[21]/a').click()
            elif 'Dalia' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[22]/a').click()
            elif 'Daniela' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[23]/a').click()
            elif 'Emry' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[24]/a').click()
            elif 'Evie' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[25]/a').click()
            elif 'Fulton' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[26]/a').click()
            elif 'Ginny' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[27]/a').click()
            elif 'Gramercy' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[28]/a').click()
            elif 'Greenwich' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[29]/a').click()
            elif 'Gwen' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[30]/a').click()
            elif 'Hamilton' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[31]/a').click()
            elif 'Harrison' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[32]/a').click()
            elif 'Hayley' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[33]/a').click()
            elif 'Hutton' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[34]/a').click()
            elif 'Izzy' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[35]/a').click()
            elif 'Jade' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[36]/a').click()
            elif 'James' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[37]/a').click()
            elif 'Jet Set Travel' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[38]/a').click()
            elif 'Kent' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[39]/a').click()
            elif 'Loren' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[40]/a').click()
            elif 'Mae' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[41]/a').click()
            elif 'Maldives' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[42]/a').click()
            elif 'Malibu' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[43]/a').click()
            elif 'Mercer' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[44]/a').click()
            elif 'Miranda' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[45]/a').click()
            elif 'Mott' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[46]/a').click()
            elif 'Naomi' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[47]/a').click()
            elif 'Nolita' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[48]/a').click()
            elif 'Odin' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[49]/a').click()
            elif 'Prescott' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[50]/a').click()
            elif 'Ravem' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[51]/a').click()
            elif 'Reagan' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[52]/a').click()
            elif 'Rhea' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[53]/a').click()
            elif 'Riey' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[54]/a').click()
            elif 'Runway' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[55]/a').click()
            elif 'Santorini' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[56]/a').click()
            elif 'Savannah' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[57]/a').click()
            elif 'Scout' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[58]/a').click()
            elif 'Selma' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[59]/a').click()
            elif 'Skoppios' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[60]/a').click()
            elif 'Sloan' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[61]/a').click()
            elif 'Stanwyck' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[62]/a').click()
            elif 'Venice' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[63]/a').click()
            elif 'Viv' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[64]/a').click()
            elif 'Vivian' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[65]/a').click()
            elif 'Voyager' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[66]/a').click()
            elif 'Whitney' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[67]/a').click()
            elif 'Wythe' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[68]/a').click()
            elif 'Yasmeen' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[69]/a').click()
            else:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == 'Coach(コーチ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(C).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            if 'baseball' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            elif 'bleecker' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[2]/a').click()
            elif 'classic' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[3]/a').click()
            elif 'crosby' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[4]/a').click()
            elif 'dinky' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[5]/a').click()
            elif 'edie' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[6]/a').click()
            elif 'hamptons' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[7]/a').click()
            elif 'harrison' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[8]/a').click()
            elif 'heritage' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[9]/a').click()
            elif 'hugo' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[10]/a').click()
            elif 'julia' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[11]/a').click()
            elif 'kristin' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[12]/a').click()
            elif 'legacy' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[13]/a').click()
            elif 'madison' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[14]/a').click()
            elif 'margot' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[15]/a').click()
            elif 'mercer' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[16]/a').click()
            elif 'nolita' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[17]/a').click()
            elif 'op art' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[18]/a').click()
            elif 'poppy' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[19]/a').click()
            elif 'rogue' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[20]/a').click()
            elif 'saddle' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[21]/a').click()
            elif 'signature' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[22]/a').click()
            elif 'stanton' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[23]/a').click()
            elif 'swagger' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[24]/a').click()
            elif 'thompson' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[25]/a').click()
            elif 'transatlantic' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[26]/a').click()
            elif 'turnlock' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[27]/a').click()
            elif 'washed canvas' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[28]/a').click()
            elif 'willis' in brandDetails:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[29]/a').click()
            else:
                browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == 'TRUE RELIGION(トゥルーレリジョン)': 
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(T).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'PRPS(ピーアールピーエス)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(P).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'John Varvatos(ジョンバルベイトス)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(J).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'RVCA(ルカ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(R).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Young & Reckless(ヤングアンドレックレス)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(Y).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'OPENING CEREMONY(オープニングセレモニー)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(O).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Alternative Apparel(オルタナティブアパレル)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(A).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'KINFOLK(キンフォーク)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(K).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Burton(バートン)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(B).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'R13(アール13)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(R).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'PX clothing(ピーエックスクロージング)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(P).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'TED BAKER(テッドベーカー)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(T).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Club Monaco(クラブモナコ)': 
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(C).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'GANT(ガント)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(G).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'REISS(リース)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(R).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Volcom(ボルコム)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(V).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'TRUE RELIGION(トゥルーレリジョン)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(T).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == '7 For All Mankind(セブンフォーオールマンカインド)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(7).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Scotch & Soda(スコッチアンドソーダ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(S).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'ANDREW MARC(アンデュリューマーク)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(A).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'FRENCH CONNECTION(フレンチコネクション)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(F).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'TED BAKER(テッドベーカー)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(T).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'WARBY PARKER(ワービーパーカー)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(W).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'DIANE von FURSTENBERG(ダイアンフォンファステンバーグ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(D).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Joe Fresh(ジョーフレッシュ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(J).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Rebecca Minkoff(レベッカミンコフ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(R).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Columbia(コロンビア)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(C).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'COS(コス)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(C).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'BB Dakota(ビービーダコタ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(B).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Salvatore Ferragamo(サルヴァトーレフェラガモ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(S).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
        elif brand == 'Anthropologie(アンソロポロジー)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath(A).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
    
        elif brand == 'Dr Martens(ドクターマーチン)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath(D).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == 'UGG(アグ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath(U).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == 'LACOSTE(ラコステ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath(L).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)
        elif brand == 'Kappa(カッパ)':
            browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath(K).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_link_text(brand).click()
            browser.implicitly_wait(5)
            time.sleep(1)
            browser.find_element_by_xpath('/html/body/div[8]/div[2]/div/div[1]/ul/li[1]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.find_element_by_xpath('//*[@id="lstItemModel"]').send_keys("指定なし")                    
            browser.implicitly_wait(40)
            time.sleep(1)
            
            

        browser.implicitly_wait(40)
        time.sleep(1)


        ###商品名###
        ############
        productName = xl_sh.cell_value(yoko,1)        
        browser.find_element_by_xpath('//*[@id="item_name"]').send_keys(productName)
        browser.implicitly_wait(40)
        time.sleep(1)



        ###　商品写真　###
        #################
        imageNum = 1
        dirs = photoPath + excelName + '\\' + str(folderNum)
        files = os.listdir(dirs)# ファイルのリストを取得

        for file in files:

            pictures = photoPath + excelName + '\\' + str(folderNum) + '\\' + file
            images = browser.find_element_by_xpath('//*[@id="js-async-upload-area"]/div/ul/li[' + str(imageNum) + ']/div/div/input')
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.execute_script("arguments[0].style.display = 'block';", images)
            browser.implicitly_wait(40)
            time.sleep(1)
            images.send_keys(pictures)
            browser.implicitly_wait(40)
            time.sleep(1)

            imageNum += 1

        folderNum += 1  
            
        


        ###色とサイズ###
        ###############
        white = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[1]/span[1]'
        black = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[2]/span[1]'
        grey = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[3]/span[1]'
        brown = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[4]/span[1]'
        beju = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[5]/span[1]'
        green = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[6]/span[1]'
        blue = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[7]/span[1]'
        navy = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[8]/span[1]'
        purple = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[9]/span[1]'
        yellow = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[10]/span[1]'
        pink = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[11]/span[1]'
        red = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[12]/span[1]'
        orange = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[13]/span[1]'
        silver = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[14]/span[1]'
        gold = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[15]/span[1]'
        clear = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[16]/span[1]'
        multiColor = '//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[1]/div/a[17]/span[1]'


        #設定するボタンをクリック
        browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/form/div[3]/div[2]/table/tbody/tr[6]/td/div[1]/div[1]/a[1]').click()
        browser.implicitly_wait(40)

        #複数色を選択、入力
        color = [8,11,14,17,20,23,26,29,32,35]          
        for i in color:
            color = xl_sh.cell_value(yoko,i)
            colorName = xl_sh.cell_value(yoko,i - 1)
            if color == '':
                break
            
            
            #色アイコンの指定
            if color == 'ホワイト':
                browser.find_element_by_xpath(white).click()
            elif color == 'ブラック':
                browser.find_element_by_xpath(black).click()
            elif color == 'グレー':
                browser.find_element_by_xpath(grey).click()
            elif color == 'ブラウン':
                browser.find_element_by_xpath(brown).click()
            elif color == 'ベージュ':
                browser.find_element_by_xpath(beju).click()
            elif color == 'グリーン':
                browser.find_element_by_xpath(green).click()
            elif color == 'ブルー':
                browser.find_element_by_xpath(blue).click()
            elif color == 'ネイビー':
                browser.find_element_by_xpath(navy).click()
            elif color == 'パープル':
                browser.find_element_by_xpath(purple).click()
            elif color == 'イエロー':
                browser.find_element_by_xpath(yellow).click()
            elif color == 'ピンク':
                browser.find_element_by_xpath(pink).click()
            elif color == 'レッド':
                browser.find_element_by_xpath(red).click()
            elif color == 'オレンジ':
                browser.find_element_by_xpath(orange).click()
            elif color == 'シルバー':
                browser.find_element_by_xpath(silver).click()
            elif color == 'ゴールド':
                browser.find_element_by_xpath(gold).click()
            elif color == 'クリア':
                browser.find_element_by_xpath(clear).click()
            elif color == 'マルチカラー':
                browser.find_element_by_xpath(multiColor).click()
            #色アイコンの名称を記入
            browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[2]/div/input').send_keys(colorName)
            browser.implicitly_wait(40)
            time.sleep(2)
            #追加するボタンをクリック
            browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[3]/div/div/div/div[2]/div/a').click()




        #全色の重複しない全てのサイズをリストにまとめる
        newSizeLIST = []
        newSizecolum = [9,12,15,18,21,24,27,30,33,36]
        for i in newSizecolum:
            newSize = xl_sh.cell_value(yoko,i)
            if newSize == '':
                lennewSize = len(newSizeLIST)
                break
            
            #サイズをリストとしてスプリットさせる
            newSize = newSize.split(',')
            #もしリストにサイズがなければ追加
            for i in newSize:

                #sortでサイズ順に並ぶように一時変換
                if i == 'XS':
                    i = 'AS'
                elif i == 'XXS':
                    i = 'AAS'
                elif i == 'S':
                    i = 'B'
                elif i == 'M':
                    i = 'C'
                elif i == 'L':
                    i = 'D'
                elif i == '5/6':
                    i = '1'
                elif i == '7/8':
                    i = '11'
                elif i == '9/10':
                    i = '11/'

                if not i in newSizeLIST:
                    newSizeLIST.append(i)


        #一番多いサイズ数分ボックスを追加する
        count = 1
        while True:
            if count == lennewSize:
                break
            browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/div[4]/a').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            count += 1


        #サイズを反映
        sizecheckBox = []
        sizeNumber = 2
        newSizeLIST.sort()
        for i in newSizeLIST:
            #サイズを元に変換
            if i == 'AS':
                i = 'XS'
            elif i == 'AAS':
                i = 'XXS'
            elif i == 'B':
                i = 'S'
            elif i == 'C':
                i = 'M'
            elif i == 'D':
                i = 'L'
            elif i == '1':
                i = '5/6'
            elif i == '11':
                i = '7/8'
            elif i == '11/':
                i = '9/10'

            if not i in sizecheckBox:
                sizecheckBox.append(i)
                

            #サイズを反映
            browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[2]/input').send_keys(i)
            browser.implicitly_wait(40)
            time.sleep(1)
            
            #日本参考サイズ(洋服)
            if category == 'メンズファッション 靴・ブーツ・サンダル 靴・ブーツ・サンダルその他' or category == 'メンズファッション 靴・ブーツ・サンダル スニーカー' or category == 'メンズファッション 靴・ブーツ・サンダル サンダル':
                if i == '4' or i == '4.5' or i == '5':
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '5.5' or i == '35 1/2':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("23.5cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '6' or i == '36':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("24cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '6.5' or i == '36 1/2':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("24.5cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '7' or i == '37' or i == 'M 7/W 8.5' or i == 'US 7/EU 40':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("25cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '7.5' or i == '37 1/2' or i == 'M 7.5/W 9':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("25.5cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '8' or i == '38' or i == 'M 8/W 9.5' or i == 'US 8/EU 41':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("26cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '8.5' or i == '38 1/2' or i == 'M 8.5/W 10':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("26.5cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '9' or i == '39' or i == 'M 9/W 10.5' or i == 'US 9/EU 42':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("27cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '9.5' or i == '39 1/2' or i == 'M 9.5/W 11' or i == 'US 9.5/EU 43':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("27.5cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '10' or i == '40' or i == 'M 10/W 11.5' or i == 'US 10/EU 44':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("28cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '10.5' or i == '40 1/2' or i == 'M 10.5/W 12':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("28.5cm")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == '11' or i == '41' or i == 'M 11/W 12.5' or i == 'US 11/EU 45':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("29cm以上")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == 'M 11.5/W 13':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("29cm以上")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == 'M 12/W 13.5' or i == 'US 12/EU 46' or i == '12':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("29cm以上")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == 'M 12.5/W 14' or i == '12.5':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("29cm以上")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == 'M 13/W 14.5' or i == '13':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("29cm以上")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == 'M 13.5/W 15' or i == '13.5':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("29cm以上")
                    browser.implicitly_wait(40)
                    time.sleep(1)
                elif i == 'M 14/W 15.5' or i == '14':
                    browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("29cm以上")
                    browser.implicitly_wait(40)
                    time.sleep(1)


                
            elif i == 'XXS' or i == '2' or i == '24':
                browser.implicitly_wait(40)
                time.sleep(1)
            elif i == 'XS' or i == '4' or i == '25' or i == 'XS/S' or i == '30/32':
                browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("S")
                browser.implicitly_wait(40)
                time.sleep(1)
            elif i == 'S' or i == '6' or i == '26' or i == 'S/M' or i == '32/32':
                browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("M")
                browser.implicitly_wait(40)
                time.sleep(1)
            elif i == 'M' or i == '8' or i == '27' or i == 'M/L' or i == '34/32':
                browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("L")
                browser.implicitly_wait(40)
                time.sleep(1)
            elif i == 'L' or i == 'XL' or i == 'XXL' or i == 'L/XL' or i == '10' or i == '12' or i == '14' or i == '28' or i == '29' or i == '30' or i == '31' or i == '32' or i == '36/32':
                browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("XL以上")
                browser.implicitly_wait(40)
                time.sleep(1)
            elif i == 'ONE SIZE':
                browser.find_element_by_xpath(' //*[@id="rdoSelectSize2"]').click()
                browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys("ONE SIZE")
                break
            else:
                browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(sizeNumber) + ']/td[3]/select').send_keys(i)
                browser.implicitly_wait(40)
                time.sleep(1)

            sizeNumber += 1



        #全色の色のチェックを外す
        for i in newSizecolum:
            newSize = xl_sh.cell_value(yoko,i)
            if newSize == '':
                break

            newSize = newSize.split(',')

            if i == 9:
                yokocheck = '4'
            elif i == 12:
                yokocheck = '5'
            elif i == 15:
                yokocheck = '6'
            elif i == 18:
                yokocheck = '7'
            elif i == 21:
                yokocheck = '8'
            elif i == 24:
                yokocheck = '9'
            elif i == 27:
                yokocheck = '10'
            elif i == 30:
                yokocheck = '11'
            elif i == 33:
                yokocheck = '12'
            elif i == 36:
                yokocheck = '13'
                
            tatecheck = 2
            for i in sizecheckBox:
                browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(tatecheck) + ']/td[' + yokocheck + ']/input').click()
                tatecheck += 1


        #色ごとにサイズボックスの色を入れる
        for i in newSizecolum:
            newSize = xl_sh.cell_value(yoko,i)
            if newSize == '':
                break
            newSize = newSize.split(',')

            if i == 9:
                yokocheck = '4'
            elif i == 12:
                yokocheck = '5'
            elif i == 15:
                yokocheck = '6'
            elif i == 18:
                yokocheck = '7'
            elif i == 21:
                yokocheck = '8'
            elif i == 24:
                yokocheck = '9'
            elif i == 27:
                yokocheck = '10'
            elif i == 30:
                yokocheck = '11'
            elif i == 33:
                yokocheck = '12'
            elif i == 36:
                yokocheck = '13'



            for ii in newSize:
                tatecheck = 2

                for a in sizecheckBox:
                    if a == ii:
                        browser.find_element_by_xpath('//*[@id="components"]/div/div[1]/div[4]/table/tbody/tr[' + str(tatecheck) + ']/td[' + yokocheck + ']/input').click()
                    tatecheck += 1
                
                   
                
                

        #全サイズ反映後、設定ボタンをクリック
        browser.find_element_by_xpath('//*[@id="components"]/div/div[2]/a[2]').click()



        #サイズ補足説明欄###
        ##################
        category = xl_sh.cell_value(yoko,2)
        if brand == 'Hollister Co.(ホリスター)':
            if 'メンズ' in category:
                sizeComment = "★メンズ参考サイズ★\n\n基本注文確認後、買い付けしています。手元に在庫がありませんので下記の参考サイズを参考にして頂ければ幸いです。\n\nサイズに不安な場合は、注文後、買い付けが完了しだい実寸の平置きをお知らせすることはできます。確認後のサイズ変更も可能です。\n\n（胸囲cm）\nS　91 - 96\nM　97 - 101\nL　102 - 106\nXL 107 - 111\n\n（袖長さcm）\nS　82 - 85\nM　86 - 87\nL　89 - 90\nXL 91 - 93\n\n\n（ウェスト）\nXS 28 (71cm)\nS　 29 - 30 (74-76cm)\nM　 31 - 32 (79-81cm)\nL　 33 - 34 (84-86cm)\nXL 36 (89cm)\n\n(足のサイズ）\nS　26.5ｃｍ\nM　27.5ｃｍ\nL　28.3ｃｍ\nXL　29.1ｃｍ\n\nその他アバクロ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816/\nその他ホリスター商品はこちら→https://www.buyma.com/r/_HOLLISTER-%E3%83%9B%E3%83%AA%E3%82%B9%E3%82%BF%E3%83%BC/-B4256816/\nその他アバクロキッズ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816F1/%E5%A4%A7%E4%BA%BA%E3%82%82OK/"
            elif 'レディース' in category:
                sizeComment = "★レディース参考サイズ★\n\n基本注文確認後、買い付けしています。手元に在庫がありませんので下記の参考サイズを参考にして頂ければ幸いです。\n\nサイズに不安な場合は、注文後、買い付けが完了しだい実寸の平置きをお知らせすることはできます。確認後のサイズ変更も可能です。\n\n（胸囲cm）\nXS　80 - 84　（5 - 7号）\nS　86 - 89　　（９号）\nM　91 - 94　　（１１号）\nL　96 - 97　　（１３号）\n\n\n（ウェスト INCHES）\nXS　23 - 25　（5 - 7号）\nS　26 - 27　 （7 - 9号）\nM　28 - 29　 （9 - 11号）\nL　30 - 31　 （11 - 13号）\n\n(足のサイズ）\nXS　23.2ｃｍ\nS　23.8ｃｍ\nM　24.8ｃｍ\nL　25.4ｃｍ\nXL　29.1ｃｍ\n\nその他アバクロ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816/\nその他ホリスター商品はこちら→https://www.buyma.com/r/_HOLLISTER-%E3%83%9B%E3%83%AA%E3%82%B9%E3%82%BF%E3%83%BC/-B4256816/\nその他アバクロキッズ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816F1/%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD%E3%82%AD%E3%83%83%E3%82%BA/"
        elif brand == 'Abercrombie & Fitch(アバクロ)':
            if 'メンズ' in category:
                sizeComment = "★メンズ参考サイズ★\n\n基本注文確認後、買い付けしています。手元に在庫がありませんので下記の参考サイズを参考にして頂ければ幸いです。\n\nサイズに不安な場合は、注文後、買い付けが完了しだい実寸の平置きをお知らせすることはできます。確認後のサイズ変更も可能です。\n\n（胸囲cm）\nS　91 - 96\nM　97 - 101\nL　102 - 106\nXL 107 - 111\n\n（袖長さcm）\nS　82 - 85\nM　86 - 87\nL　89 - 90\nXL 91 - 93\n\n\n（ウェスト）\nXS 28 (71cm)\nS　 29 - 30 (74-76cm)\nM　 31 - 32 (79-81cm)\nL　 33 - 34 (84-86cm)\nXL 36 (89cm)\n\n(足のサイズ）\nS　26.5ｃｍ\nM　27.5ｃｍ\nL　28.3ｃｍ\nXL　29.1ｃｍ\n\nその他アバクロ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816/\nその他ホリスター商品はこちら→https://www.buyma.com/r/_HOLLISTER-%E3%83%9B%E3%83%AA%E3%82%B9%E3%82%BF%E3%83%BC/-B4256816/\nその他アバクロキッズ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816F1/%E5%A4%A7%E4%BA%BA%E3%82%82OK/"
            elif 'レディース' in category:
                sizeComment = "★レディース参考サイズ★\n\n基本注文確認後、買い付けしています。手元に在庫がありませんので下記の参考サイズを参考にして頂ければ幸いです。\n\nサイズに不安な場合は、注文後、買い付けが完了しだい実寸の平置きをお知らせすることはできます。確認後のサイズ変更も可能です。\n\n（胸囲cm）\nXS　80 - 84　（5 - 7号）\nS　86 - 89　　（９号）\nM　91 - 94　　（１１号）\nL　96 - 97　　（１３号）\n\n\n（ウェスト INCHES）\nXS　23 - 25　（5 - 7号）\nS　26 - 27　 （7 - 9号）\nM　28 - 29　 （9 - 11号）\nL　30 - 31　 （11 - 13号）\n\n(足のサイズ）\nXS　23.2ｃｍ\nS　23.8ｃｍ\nM　24.8ｃｍ\nL　25.4ｃｍ\nXL　29.1ｃｍ\n\nその他アバクロ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816/\nその他ホリスター商品はこちら→https://www.buyma.com/r/_HOLLISTER-%E3%83%9B%E3%83%AA%E3%82%B9%E3%82%BF%E3%83%BC/-B4256816/\nその他アバクロキッズ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816F1/%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD%E3%82%AD%E3%83%83%E3%82%BA/"
            elif 'ベビー・キッズ' in category:
                sizeComment = "★アバクロキッズ参考サイズ★\n\n基本注文確認後、買い付けしています。手元に在庫がありませんので下記の参考サイズを参考にして頂ければ幸いです。\n\n（身長cm）\n5/6　110 - 122\n7/8　122 - 135\n9/10　135 - 145\n11/12　145 - 152\n13/14　152 - 160\n15/16　160 - 165\n\n（胸囲ｃｍ）\n5/6　58 - 64\n7/8　64 - 69\n9/10　69 - 72\n11/12　72 - 76\n13/14　76 - 80\n15/16　80 - 84\n\n(足のサイズ）\n12/13　18.4ｃｍ\n1/2　20.3ｃｍ\n3/4　21.9ｃｍ\n5/6　23.5ｃｍ\n7/8　24.8ｃｍ\n\n※在庫の変動が激しいので、購入前に在庫確認をよろしくお願いします。\n\nその他アバクロ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816/\nその他ホリスター商品はこちら→https://www.buyma.com/r/_HOLLISTER-%E3%83%9B%E3%83%AA%E3%82%B9%E3%82%BF%E3%83%BC/-B4256816/\nその他アバクロキッズ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816F1/%E5%A4%A7%E4%BA%BA%E3%82%82OK/"
        else:
            sizeComment = "※在庫の変動が激しいので在庫確認をよろしくお願いします。\n\n" + '基本注文確認後、買い付けしています。手元に在庫がありませんので下記の参考サイズを参考にして頂ければ幸いです。\n\n※ご注文後、お取り寄せ後に正確なサイズをお伝えすることも可能です。お気軽にお問合せください。\nサイズ (バスト / ナチュラルウエスト / ヒップ)\nXS　0-2（約81-86cm、61-66cm、89-91cm）\nＳ　4-6（約89-91cm、68-71cm、94-97cm）\nＭ　8-10（約94-97cm、73-76cm、99-102cm）\nＬ　12-14（約99cm、79cm、104cm）\nXL　16（約102cm、81cm、107cm）\n\n' + xl_sh.cell_value(yoko,37)
            browser.find_element_by_xpath('//*[@id="item_color_size"]').send_keys(sizeComment)
            browser.implicitly_wait(40)
            time.sleep(1)
            
        browser.implicitly_wait(40)
        time.sleep(1)

         

        ###商品コメント###
        #################
        itemcomment = xl_sh.cell_value(yoko,0)
        comment = '大人気の「' + itemcomment + '」をお届けします。\n\n（店舗にて売れ切れの場合はオンラインで発注します）\n●発送方法は、基本アメリカからファーストクラス便で発送します。\n発送後、到着までに早ければ１週間、税関や空輸が混雑していますと２週間-３週間掛かることもあります。\n●直接店舗で買い付けた場合は商品に、店舗で使われている香水の匂い、多少のヨレ感がありますこと予めご了承ください。\n●商品発送前に入念に検品をして発送することを徹底して心がけております。\n●商品の在庫数が極限られていますので、受注時に既に売れ切れている場合がございます。その場合にはキャンセルという形で対応させていただきますのでご理解ください。\n（バイマよりご返金)'
        browser.find_element_by_xpath('//*[@id="item_comment"]').send_keys(comment)
        browser.implicitly_wait(40)
        time.sleep(1)

        

        ###シーズン###
        #############
        season = "2018-19AW"
        browser.find_element_by_xpath('//*[@id="season"]').send_keys(season)
        browser.implicitly_wait(40)
        time.sleep(1)

        


        ### タグ　###
        #############

        if "子供服" in category:
            ninki = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[5]/a'
            trend = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[9]/a'
            otona = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[2]/a'        
        elif "メンズ" in category and "トップス" in category:
            ninki = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[8]/a'
            trend = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[12]/a'
        elif "レディース" in category and "トップス" in category or "靴" in category or "ワンピース" in category or "ブーツ" in category:
            ninki = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[10]/a'
            trend = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[14]/a'
        elif "レディース" in category and "財布" in category or "水着" in category or "アクセサリー" in category or "パジャマ" in category:
            ninki = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[6]/a'
            trend = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[10]/a'
        elif "レディース" in category and "バッグ" in category or "腕時計" in category or "雑貨" in category or "カバン" in category:
            ninki = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[7]/a'
            trend = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[11]/a'
        elif "メンズ" in category and "ボトムス" in category or "アクセサリー" in category or "財布" in category or "雑貨" in category or "バッグ" in category:
            ninki = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[6]/a'
            trend = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[10]/a'
        elif "メンズ" in category and "アウター" in category or "靴" in category:
            ninki = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[7]/a'
            trend = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[11]/a'
        elif "レディース" in category and "ボトムス" in category or "アウター" in category:
            ninki = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[8]/a'
            trend = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[12]/a'
        elif "帽子" in category:
            ninki = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[4]/a'
            trend = '//*[@id="r_tag_select_box"]/div/div[3]/ul/li[8]/a'


        time.sleep(3)            
        browser.find_element_by_xpath('//*[@id="chkForm"]/div[3]/div[2]/table/tbody/tr[11]/td/div[2]/a').click()
        browser.implicitly_wait(40)
        time.sleep(1)
        
        tag = xl_sh.cell_value(yoko,41)


        browser.implicitly_wait(40)
        time.sleep(3)


        if "日本未入荷" in tag:
            time.sleep(3)
            browser.find_element_by_xpath('//*[@id="r_tag_select_box"]/div/div[3]/div/div[1]/div[1]/div[2]/div/div/label').click()
            browser.implicitly_wait(40)
            time.sleep(3)
        if "ストリート" in tag:
            time.sleep(3)
            browser.find_element_by_xpath(ninki).click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath('//*[@id="_r_tag_check_430_8"]').click()
            browser.implicitly_wait(40)
            time.sleep(3)
        if "ベイクドカラー" in tag:
            time.sleep(3)
            browser.find_element_by_xpath(trend).click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath('//*[@id="_r_tag_check_1017_95"]').click()
            browser.implicitly_wait(40)
            time.sleep(3)
        if "もこもこ" in tag:
            time.sleep(3)
            browser.find_element_by_xpath(trend).click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath('//*[@id="_r_tag_check_872_95"]').click()
            browser.implicitly_wait(40)
            time.sleep(3)
        if "大人もok" in tag:
            time.sleep(3)
            browser.find_element_by_xpath(otona).click()
            browser.implicitly_wait(40)
            time.sleep(3)
            browser.find_element_by_xpath('//*[@id="_r_tag_check_413_24"]').click()
            browser.implicitly_wait(40)
            time.sleep(3)
                
        time.sleep(3)       
        browser.find_element_by_xpath('//*[@id="r_tag_select_box"]/div/div[7]/ul/li/button').click()
        browser.implicitly_wait(40)
        time.sleep(1)



        ###値段###
        ##########
        if 'シャツ' in category or '帽子' in category  or '水着' in category or 'キッズ用トップス' in category  or '子供用帽子' in category:
            sellingPrice = (((float(xl_sh.cell_value(yoko,3)))*100*1.076*1.2)+2000)*1.08
            sellingPrice = round(sellingPrice, -2)
            browser.find_element_by_xpath('//*[@id="price"]').send_keys(int(sellingPrice))
            browser.implicitly_wait(40)
            time.sleep(1)
        elif 'アウター' in category:
            sellingPrice = (((float(xl_sh.cell_value(yoko,3)))*100*1.076*1.2)+3000)*1.08
            sellingPrice = round(sellingPrice, -2)
            browser.find_element_by_xpath('//*[@id="price"]').send_keys(int(sellingPrice))
            browser.implicitly_wait(40)
            time.sleep(1)
        elif 'ブーツ' in category:
            sellingPrice = (((float(xl_sh.cell_value(yoko,3)))*100*1.076*1.2)+7000)*1.08
            sellingPrice = round(sellingPrice, -2)
            browser.find_element_by_xpath('//*[@id="price"]').send_keys(int(sellingPrice))
            browser.implicitly_wait(40)
            time.sleep(1)
        else:
            sellingPrice = (((float(xl_sh.cell_value(yoko,3)))*100*1.076*1.2)+4000)*1.08
            sellingPrice = round(sellingPrice, -2)
            browser.find_element_by_xpath('//*[@id="price"]').send_keys(int(sellingPrice))
            browser.implicitly_wait(40)
            time.sleep(1)
            

        

        ###配送方法###
        #############
        shipping = "360216"
        tracking = '360215'

        if sellingPrice <= 50000:
            sippingId = '//*[@id="shipping-method-checkbox' + str(shipping) + '"]'
            browser.find_element_by_xpath(sippingId).click()
            browser.implicitly_wait(40)
            time.sleep(1)
            trackingID = '//*[@id="shipping-method-checkbox' + str(tracking) + '"]'
            browser.find_element_by_xpath(trackingID).click()
            browser.implicitly_wait(40)
            time.sleep(1)
        elif sellingPrice >= 50000:
            trackingID = '//*[@id="shipping-method-checkbox' + str(tracking) + '"]'
            browser.find_element_by_xpath(trackingID).click()
            browser.implicitly_wait(40)
            time.sleep(1)

            
        '''
        #参考価格###
        ###########
        regularPrice = "190000"
        browser.find_element_by_xpath('//*[@id="itemedit[reference_price_kbn]2"]').click()
        browser.find_element_by_xpath('//*[@id="reference_price"]').send_keys(regularPrice)
        browser.implicitly_wait(40)
        time.sleep(1)
        '''

        ###数量###
        ##########
        productNum = "2"
        browser.find_element_by_xpath('//*[@id="pieces"]').send_keys(productNum)
        browser.implicitly_wait(40)
        time.sleep(1)



        ###買い付け発送地###
        ##################
        browser.find_element_by_xpath('//*[@id="rdoMyActArea2"]').click()
        browser.implicitly_wait(40)
        time.sleep(1)
        browser.find_element_by_xpath('//*[@id="itemedit_purchase_area"]').send_keys('北米')
        browser.implicitly_wait(40)
        time.sleep(1)
        browser.find_element_by_xpath('//*[@id="rdoMyHassoArea2"]').click()
        browser.implicitly_wait(40)
        time.sleep(1)
        browser.find_element_by_xpath('//*[@id="hasso_foreign"]').send_keys('北米')
        browser.implicitly_wait(40)
        time.sleep(1)
        

        '''
        #ショップ名###
        #############
        shopName = xl_sh.cell_value(yoko,6)
        browser.find_element_by_xpath('//*[@id="itemedit_konyuchi"]').send_keys(shopName)
        browser.implicitly_wait(40)
        time.sleep(1)
        '''

        ###入力内容を確認ボタン###
        ########################
        browser.find_element_by_xpath('//*[@id="confirmButton"]').click()
        browser.implicitly_wait(40)
        time.sleep(1)


        ###出品完了###
        #############
        browser.find_element_by_xpath('//*[@id="done"]').click()
        browser.implicitly_wait(40)
        time.sleep(1)


        ###出品URLを保存###
        ##################
        browser.find_element_by_link_text('出品リストへ戻る').click()
        browser.implicitly_wait(40)
        time.sleep(1)
        ItemURL = browser.find_element_by_xpath('//*[@id="inputform"]/table/tbody/tr[2]/td[4]/p[2]/a[1]').text
        ItemURL = 'https://www.buyma.com/my/itemedit/?iid=' + ItemURL
        browser.implicitly_wait(40)
        time.sleep(1)
        ws['AN' + str(int(yoko) + 1)].value = ItemURL
        wb.save(spreadSheetPath + excelName + '.xlsx')


        ###出品画面に移動###
        ###################

        browser.get('https://www.buyma.com/my/itemedit/')
        browser.implicitly_wait(40)
        time.sleep(1)


        ###次の行に行く###
        #################
        yoko += 1





    ####################
    ### サイズ変更なし ###
    ####################

    else:
        yoko += 1
        folderNum += 1




