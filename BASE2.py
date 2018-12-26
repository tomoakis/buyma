import unittest, time, requests, webbrowser, bs4, datetime
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import csv, os, re, xlrd, openpyxl
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.alert import Alert


###############
### 基本設定 ###
###############

#エクセル指定
excelName = input('エクセル名:')
xl_bk = xlrd.open_workbook(excelName + '.xlsx')
xl_sh = xl_bk.sheet_by_name(excelName)
wb = openpyxl.load_workbook('C:\\Users\\tomoa\\Desktop\\' + excelName + '.xlsx') 
ws = wb.active

#列指定
yoko = input('出品番号(途中から出品したい場合):')
if not yoko == '':
    yoko = int(yoko) - 1
    folderNum = int(yoko) + 1
elif yoko == '':
    yoko = 1
    folderNum = int(yoko) + 1

###　ログイン　###
#################
browser = webdriver.Chrome()
browser.get("https://admin.thebase.in/items/add")
email = browser.find_element_by_id('loginUserMailAddress')
email.send_keys('namitaketomi123@gmail.com')
password = browser.find_element_by_id('UserPassword')
password.send_keys('Tomoaki0093')
browser.find_element_by_xpath('//*[@id="userLoginForm"]/button').click()


while True:
    print(yoko)
    try:
        browser.implicitly_wait(40)
        time.sleep(3)

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
            browser.find_element_by_xpath('//*[@id="13307596"]/div[6]/label/div/ul').click()
            browser.implicitly_wait(40)
            time.sleep(1)
            browser.get("https://admin.thebase.in/items")
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
            browser.implicitly_wait(40)
            time.sleep(1)






            
        #################
        ### 新商品出品 ###
        #################
        
        elif action == 'new':
        
        
            ###商品名###
            ############
            productName = xl_sh.cell_value(yoko,1)        
            browser.find_element_by_xpath('//*[@id="mail"]').send_keys(productName)
            browser.implicitly_wait(40)
            time.sleep(1)

            ###値段###
            ##########
            sellingPrice = (((float(xl_sh.cell_value(yoko,3)))*100*1.0685*1.2)+6000)*1.08
            sellingPrice = round(sellingPrice, -2)
            browser.find_element_by_xpath('//*[@id="price"]').send_keys(int(sellingPrice))
            browser.implicitly_wait(40)
            time.sleep(1)

            
            ###商品コメント + 色と在庫###
            ##########################
            itemcomment = xl_sh.cell_value(yoko,0)
            sizecomment = '大人気商品の「' + itemcomment + '」をお届けします。\n\n購入の際は、ご希望の色とサイズをお知らせください。\n\n============================'
            browser.find_element_by_xpath('//*[@id="ItemDetail"]').send_keys(sizecomment)

            colorNAME = [7,10,13,16,19,22,25,28,31,34]
            for i in colorNAME:
                colorName = xl_sh.cell_value(yoko,i)
                sizeName = xl_sh.cell_value(yoko,i + 2)
                if colorName == '':
                    break
                sizecomment2 = '\n' + colorName + '\n' + sizeName + '\n\n'
                browser.find_element_by_xpath('//*[@id="ItemDetail"]').send_keys(sizecomment2)

            comment = '============================\n送料、関税込みのお値段です！\n\nこちらの商品は、注文後の買い付けです。\n\n在庫に限りがあり、店舗の出品回転も速いためオンライン・店舗完売の時がよくあります。\n\n●サイズなどについては、商品が手元にない場合そのため正確な数字をお知らせできないことがありますが、ご希望であれば発送前にお伝えいたします。\n\n●注文後早ければ翌日、最大1週間ほどお時間かかることもあります。\n（店舗にて売れ切れの場合はオンラインで発注します）\n\n●発送方法は、基本アメリカからファーストクラス便で発送します。\n発送後、到着までに早ければ１週間、税関や空輸が混雑していますと２週間-３週間掛かることもあります。\n\n●直接店舗で買い付けた場合は商品に、店舗で使われている香水の匂い、多少のヨレ感がありますこと予めご了承ください。\n\n●商品発送前に入念に検品をして発送することを徹底して心がけております。\n\n●商品の在庫数が極限られていますので、受注時に既に売れ切れている場合がございます。その場合にはキャンセルという形で対応させていただきますのでご理解ください。'
            browser.find_element_by_xpath('//*[@id="ItemDetail"]').send_keys(comment)
            browser.implicitly_wait(40)
            time.sleep(1)


            ###　商品写真　###
            #################
            imageNum = 1


            #一番最初の写真をアップロードする
            dir = "C:\\Users\\tomoa\\Desktop\\" + excelName + '\\' + str(folderNum)
            files = os.listdir(dir)# ファイルのリストを取得

            for file in files:
                pictures = "C:\\Users\\tomoa\\Desktop\\" + excelName + '\\' + str(folderNum) + '\\' + file
                images = browser.find_element_by_xpath('//*[@id="ddItems"]/li[' + str(imageNum) + ']/div/input')
                browser.implicitly_wait(40)
                time.sleep(1)
                images.send_keys(pictures)
                browser.implicitly_wait(40)
                time.sleep(1)

                imageNum += 1




##            #####　数量　####
##            ################
##            browser.find_element_by_xpath('//*[@id="ItemStock"]').send_keys('0')
##


        
            #####　カテゴリ　####
            ####################
            category = xl_sh.cell_value(yoko,2)
        
            #検索ボタンをクリック
            browser.find_element_by_xpath('//*[@id="x_openCatMordal"]').click()
            browser.implicitly_wait(40)
            time.sleep(1)
        
            #カテゴリ分け
            if 'メンズ' in category and 'Tシャツ' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341805"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341886"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341899"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
              

            elif 'メンズ' in category and 'パーカー' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341805"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341886"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341900"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
              
            elif 'メンズ' in category and 'セーター' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341805"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341886"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341901"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'メンズ' in category and 'アウター' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341805"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341903"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                
            elif 'メンズ' in category and 'デニム' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341805"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341887"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341905"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
        
            elif 'メンズ' in category and 'ボトムスその他' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341805"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341887"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341906"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
            
            elif 'メンズ' in category and 'スニーカー' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341805"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341888"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341908"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'メンズ' in category and 'ブーツ' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341805"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341888"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341909"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'メンズ' in category and 'バッグ' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341805"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341888"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341910"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'メンズ' in category and '帽子' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341805"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341888"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341911"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
             
            elif 'メンズ' in category and '小物' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341805"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341888"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341912"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'レディース' in category and 'Tシャツ' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341807"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341889"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341914"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'レディース' in category and 'パーカー' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341807"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341889"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341916"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'レディース' in category and 'セーター' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341807"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341889"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341917"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'レディース' in category and 'アウター' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341807"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341913"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'レディース' in category and 'デニム' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341807"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341890"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341918"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'レディース' in category and 'ボトムスその他' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341807"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341890"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341919"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'レディース' in category and 'スニーカー' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341807"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341891"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341920"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'レディース' in category and 'ブーツ' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341807"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341891"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341921"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'レディース' in category and 'バッグ' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341807"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341891"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341923"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'レディース' in category and '帽子' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341807"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341891"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341924"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'レディース' in category and '小物' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341807"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                browser.find_element_by_xpath('//*[@id="cat_1341891"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341925"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)


            browser.find_element_by_xpath('//*[@id="x_catSelectFix"]').click()
            browser.implicitly_wait(40)
            time.sleep(1)


            #ブランド選択
            brandName = xl_sh.cell_value(yoko,38)
            if 'adidas' in brandName and 'トップス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341828"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341856"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'adidas' in brandName and 'ボトムス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341828"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341857"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'adidas' in brandName:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341828"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341858"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'tommy' in brandName and 'トップス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341831"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341860"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'tommy' in brandName and 'ボトムス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341831"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341862"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'tommy' in brandName:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341831"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341865"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'stussy' in brandName and 'トップス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341833"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341863"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'stussy' in brandName and 'ボトムス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341833"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341864"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'stussy' in brandName:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341833"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341866"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'チャンピオン' in brandName and 'トップス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341835"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341867"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'チャンピオン' in brandName and 'ボトムス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341835"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341868"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'チャンピオン' in brandName:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341835"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341869"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'スケッチャーズ' in brandName:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341840"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'ラルフ' in brandName and 'トップス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341844"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341871"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'ラルフ' in brandName and 'ボトムス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341844"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341872"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'ラルフ' in brandName:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341844"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341873"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'North' in brandName and 'トップス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341847"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341874"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                

            elif 'North' in brandName and 'ボトムス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341847"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341875"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'North' in brandName:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341847"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341876"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'ナイキ' in brandName and 'トップス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341848"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341877"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'ナイキ' in brandName and 'ボトムス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341848"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341878"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)
                
            elif 'ナイキ' in brandName:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341848"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341879"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'ラコステ' in brandName and 'トップス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341851"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341880"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'ラコステ' in brandName and 'ボトムス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341851"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341881"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'ラコステ' in brandName:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341851"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341882"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'アーバン' in brandName and 'トップス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341852"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341883"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'アーバン' in brandName and 'ボトムス' in category:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341852"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341884"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)

            elif 'アーバン' in brandName:
                browser.find_element_by_xpath('//*[@id="cat_1341822"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)    
                browser.find_element_by_xpath('//*[@id="cat_1341852"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)                
                browser.find_element_by_xpath('//*[@id="cat_1341885"]').click()
                browser.implicitly_wait(40)
                time.sleep(1)


             
            browser.find_element_by_xpath('//*[@id="x_catSelectFix"]').click()
            browser.implicitly_wait(40)
            time.sleep(1)                





            ###出品完了###
            #############
            browser.find_element_by_xpath('//*[@id="x_submitItemForm"]').click()
            browser.implicitly_wait(40)
            time.sleep(1)


            ###出品URLを保存###
            ##################
            browser.find_element_by_xpath('/html/body/div[3]/div/div[4]/div[3]/ol/li[1]/div[3]/a/div[2]').click()
            cur_url = browser.current_url
            browser.implicitly_wait(40)
            time.sleep(1)
            ws['AO' + str(int(yoko) + 1)].value = cur_url
            wb.save('C:\\Users\\tomoa\\Desktop\\' + excelName + '.xlsx') 


            ###出品画面に移動###
            ###################

            browser.get('https://admin.thebase.in/items/add')
            browser.implicitly_wait(40)
            time.sleep(1)


            ###次の行に行く###
            #################
            yoko += 1
            folderNum += 1 


        else:
            yoko += 1
            folderNum += 1
            continue

        
    except IndexError:
        break


    except:
        try:
            ws['F' + str(int(yoko) + 1)].value = 'ERROR'
            wb.save('C:\\Users\\tomoa\\Desktop\\' + excelName + '.xlsx') 
            browser.get('https://admin.thebase.in/items/add')
            browser.implicitly_wait(40)
            time.sleep(1)
            yoko += 1
            folderNum = yoko + 1
        except:
            time.sleep(1)
            Alert(browser).accept()
            ws['F' + str(int(yoko) + 1)].value = 'ERROR'
            wb.save('C:\\Users\\tomoa\\Desktop\\' + excelName + '.xlsx') 
            browser.get('https://admin.thebase.in/items/add')
            browser.implicitly_wait(40)
            time.sleep(1)
            yoko += 1
            folderNum = yoko + 1
            

