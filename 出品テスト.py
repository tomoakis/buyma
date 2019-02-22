import unittest, time, requests, webbrowser, datetime, os, re, xlrd, openpyxl, yaml, actions, lib
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.alert import Alert
from selenium.webdriver import ActionChains
from enum import Enum

### アカウント設定
account = input('出品アカウント名 [ruri, yui, gen, sakihara]:')
if account == 'ruri':
    excelName = 'anthropologie'
    LogEmail = 'taira4420@gmail.com'
    LogPassword = 'taira442054'
elif account == 'yui':
    excelName = 'pacsun'
    LogEmail = 'j746ptut2@yahoo.ne.jp'
    LogPassword = 'Taketomi1219'
elif account == 'gen':
    excelName = 'zumiez'
    LogEmail = 'buyma.gen@gmail.com'
    LogPassword = 'Buyma0671'
elif account == 'sakihara':
    excelName = input('excel Name:[abercrombie, hollister]')
    LogEmail = 'namitaketomi123@gmail.com'
    LogPassword = 'seasider0093'

### エクセル指定
spreadSheetPath = 'spreadsheets\\'
photoPath = 'C:\\Users\\tomoa\\Workspace\\buyma\\画像\\'
xl_bk = xlrd.open_workbook(spreadSheetPath + excelName + '.xlsx')
xl_sh = xl_bk.sheet_by_name(excelName)
wb = openpyxl.load_workbook(spreadSheetPath + excelName + '.xlsx')
ws = wb.active

### 出品開始番号指定
yoko = input('出品開始番号:')
yoko = int(yoko) - 1
folderNum = int(yoko) + 1

### BOXの位置
class Select(Enum):
    CATEGORY = 2#カテゴリ1つ目のBOX
    SUB_CATEGORY = 12#カテゴリ2つ目のBOX
    ITEM = 13#カテゴリ3つ目のBOX
    SEASON = 3
    THEME = 4
    COLOR1 = 5
    COLOR2 = 15


###　ログイン
browser = webdriver.Chrome()
browser.get("https://www.buyma.com/my/sell/new/")
email = browser.find_element_by_id('txtLoginId')
email.send_keys(LogEmail)
password = browser.find_element_by_id('txtLoginPass')
password.send_keys(LogPassword)
browser.find_element_by_id('login_do').click()
browser.implicitly_wait(40)
time.sleep(1)

while True:
    print(folderNum)
    #try:

    ### J列読み込み
    status = xl_sh.cell_value(yoko,5)

    ### 新商品出品
    if status == 'new':

        # ###商品写真
        # dir = photoPath + excelName + '\\' + str(folderNum)
        # files = os.listdir(dir) #ファイルのリストを取得
        # for file in files:
        #     picture = photoPath + excelName + '\\' + str(folderNum) + '\\' + str(file)
        #     images = browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div/div/div/div/div[2]/form/div[1]/div/div/div[2]/div/div/div[1]/div/div/div/input')
        #     browser.implicitly_wait(40)
        #     time.sleep(1)
        #     browser.execute_script("arguments[0].style.display = 'block';", images)
        #     browser.implicitly_wait(40)
        #     time.sleep(1)
        #     images.send_keys(picture)
        #     browser.implicitly_wait(40)
        #     time.sleep(1)


        # ###ブランド名　#ブランド名は登録されているカタカナのブランド名を入力
        # brand = xl_sh.cell_value(yoko,38)
        # browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[3]/div[2]/div/div[2]/div/div/div/div/div/div[1]/div/div/div/div/input').send_keys(brand)
        # browser.implicitly_wait(40)
        # time.sleep(1)


        # ###商品名
        # productName = xl_sh.cell_value(yoko,1)
        # browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div/div/div/div/div[2]/form/div[2]/div[1]/div/div[2]/div/div/div[1]/input').send_keys(productName)
        # time.sleep(5)


        # ###商品コメント
        # productName = xl_sh.cell_value(yoko,0)
        # if 'ホリスター' in brand:
        #     comment = '大人気のホリスターから「' + productName + '」をお届けします。ホリスター・カンパニーとは、アメリカのカジュアルファッションブランドです。2000年にアバクロンビーアンドフィッチ社によって設立されたブランドであり、世界中に６００以上の店舗を構えるトップブランドの１つです。\n\n「カモメ」をモチーフにしたブランドロゴが特徴的で、アメリカ西海岸のサーファースタイルをイメージしたデザインとなっています。アメリカでは姉妹ブランドの「アバクロンビーアンドフィッチ」「ルール No ９２５」や「アメリカンイーグル」等と並んで、若い世代に絶大な人気を誇るブランドです。\n\n略してホリスターと呼ばれることが多く、アメリカの調査によると１０代の若者に２番目に人気のあるファッションブランドであり、世界的にも絶大な人気を博しているカジュアルブランドです。\n\n基本、注文後の買い付けです。\n\n在庫に限りがあり、店舗の出品回転も速いためオンライン・店舗完売の時がよくあります。\n\n●サイズなどについては、商品が手元にない場合そのため正確な数字をお知らせできないことがあります。公式サイトに記載されているサイズをそのまま記載しておりますので、そちらを参考にして頂けると幸いです。\n\n●注文後早ければ翌日、最大1週間ほどお時間かかることもあります。\n（店舗にて売れ切れの場合はオンラインで発注します）\n\n●発送方法は、基本アメリカからファーストクラス便で発送します。\n発送後、到着までに早ければ１週間、税関や空輸が混雑していますと２週間-３週間掛かることもあります。\n\n●直接店舗で買い付けた場合は商品に、店舗で使われている香水の匂い、多少のヨレ感がありますこと予めご了承ください。\n\n●商品発送前に入念に検品をして発送することを徹底して心がけております。\n\n●商品の在庫数が極限られていますので、受注時に既に売れ切れている場合がございます。その場合にはキャンセルという形で対応させていただきますのでご理解ください。\n（バイマよりご返金）'
        # elif 'アバクロ' in brand:
        #     comment = '大人気のアバクロから「' + productName + '」をお届けします。\n\n基本、注文後の買い付けです。\n\n在庫に限りがあり、店舗の出品回転も速いためオンライン・店舗完売の時がよくあります。\n\n●サイズなどについては、商品が手元にない場合そのため正確な数字をお知らせできないことがあります。公式サイトに記載されているサイズをそのまま記載しておりますので、そちらを参考にして頂けると幸いです。\n\n●注文後早ければ翌日、最大1週間ほどお時間かかることもあります。\n（店舗にて売れ切れの場合はオンラインで発注します）\n\n●発送方法は、基本アメリカからファーストクラス便で発送します。\n発送後、到着までに早ければ１週間、税関や空輸が混雑していますと２週間-３週間掛かることもあります。\n\n●直接店舗で買い付けた場合は商品に、店舗で使われている香水の匂い、多少のヨレ感がありますこと予めご了承ください。\n\n●商品発送前に入念に検品をして発送することを徹底して心がけております。\n\n●商品の在庫数が極限られていますので、受注時に既に売れ切れている場合がございます。その場合にはキャンセルという形で対応させていただきますのでご理解ください。\n（バイマよりご返金）'
        # elif 'アンソロ' in brand:
        #     comment = '大人気の「' + productName + '」をお届けします。\n\n●発送方法は、基本アメリカからファーストクラス便で発送します。\n●こちらはお取り寄せ商品となります。お取り寄せ、発送までに最大10日ほどかかります。\n●発送後、到着までに約2週間、税関や空輸が混雑していますと約３週間かかることもあります。\n●直接店舗で買い付けた場合は商品に、店舗で使われている香水の匂い、多少のヨレ感がありますこと予めご了承ください。\n●商品発送前に入念に検品をして発送することを徹底して心がけております。\n●商品の在庫数が極限られていますので、受注時に既に売れ切れている場合がございます。その場合にはキャンセルという形で対応させていただきますのでご理解ください。\n（バイマよりご返金)'
        # else:
        #     comment = '大人気の「' + productName + '」をお届けします。\n\n（店舗にて売れ切れの場合はオンラインで発注します）\n●発送方法は、基本アメリカからファーストクラス便で発送します。\n発送後、到着までに早ければ１週間、税関や空輸が混雑していますと２週間-３週間掛かることもあります。\n●直接店舗で買い付けた場合は商品に、店舗で使われている香水の匂い、多少のヨレ感がありますこと予めご了承ください。\n●商品発送前に入念に検品をして発送することを徹底して心がけております。\n●商品の在庫数が極限られていますので、受注時に既に売れ切れている場合がございます。その場合にはキャンセルという形で対応させていただきますのでご理解ください。\n（バイマよりご返金)'
        # browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div/div/div/div/div[2]/form/div[2]/div[2]/div/div[2]/div/div/div[1]/textarea').send_keys(comment)
        # time.sleep(1)


        # #####カテゴリ
        # class actions(Enum):
        #     def select(browser, id, option):
        #         # 1st click on the Select-control
        #         action = ActionChains(browser)
        #         element = browser.find_element_by_id(f'react-select-{id}--value')
        #         browser.execute_script("arguments[0].scrollIntoView();", element)
        #         action.move_to_element(element).click().perform()
        #         time.sleep(1)
        #         # 2nd click on the Select-menu-outer option
        #         element = browser.find_element_by_id(f'react-select-{id}--option-{option}')
        #         action = ActionChains(browser)
        #         action.move_to_element(element).click().perform()
        #         time.sleep(1)
        # # import category infomation
        # categoryType = xl_sh.cell_value(yoko,2)
        # categoryType = categoryType.split(' ')
        # # set category, subCategory and item
        # category = categoryType[0]
        # subCategory = categoryType[1]
        # item = categoryType[2]
        # # open yml file
        # with open('./constants/カテゴリ.yml', 'r') as file:#category, subcategory情報が入っている
        #     categoriesDoc = yaml.load(file)
        # with open(f'./constants/{category}.yml', 'r') as file:#item情報が入っている
        #     subCategoryDoc = yaml.load(file)
        # # set elements
        # id = categoriesDoc["categories"][category]['id']
        # subId = categoriesDoc["categories"][category]['subcategories'][subCategory]
        # itemId = subCategoryDoc[subCategory][item]
        # # category
        # actions.select(browser, Select.CATEGORY.value, id)
        # actions.select(browser, Select.SUB_CATEGORY.value, subId)
        # actions.select(browser, Select.ITEM.value, itemId)


        # ###シーズン
        # class actions(Enum):
        #     def select(browser, id, option):
        #         # 1st click on the Select-control
        #         action = ActionChains(browser)
        #         element = browser.find_element_by_id(f'react-select-{id}--value')
        #         browser.execute_script("arguments[0].scrollIntoView();", element)
        #         action.move_to_element(element).click().perform()
        #         time.sleep(1)
        #         # 2nd click on the Select-menu-outer option
        #         element = browser.find_element_by_id(f'react-select-{id}--option-{option}')
        #         action = ActionChains(browser)
        #         action.move_to_element(element).click().perform()
        #         time.sleep(1)
        # id = '1'#2019SS
        # actions.select(browser, Select.SEASON.value, id)


        # ###テーマ
        # class actions(Enum):
        #     def select(browser, id, option):
        #         # 1st click on the Select-control
        #         action = ActionChains(browser)
        #         element = browser.find_element_by_id(f'react-select-{id}--value')
        #         browser.execute_script("arguments[0].scrollIntoView();", element)
        #         action.move_to_element(element).click().perform()
        #         time.sleep(1)
        #         # 2nd click on the Select-menu-outer option
        #         element = browser.find_element_by_id(f'react-select-{id}--option-{option}')
        #         action = ActionChains(browser)
        #         action.move_to_element(element).click().perform()
        #         time.sleep(1)
        # # # import theme infomation
        # # categoryType = xl_sh.cell_value(yoko,6)
        # # print(categoryType)
        # # # open yml file
        # # with open('./constants/テーマ.yml', 'r') as file:
        # #     themeType = yaml.load(file)
        # # id = themeType["theme"][xl_sh.cell_value(yoko,6)]['id']
        # id = 4#日本未入荷
        # actions.select(browser, Select.THEME.value, id)


        # ### 色とサイズ
        # class actions(Enum):
        #     def select(browser, id, option):
        #         # 1st click on the Select-control
        #         action = ActionChains(browser)
        #         element = browser.find_element_by_id(f'react-select-{id}--value')
        #         browser.execute_script("arguments[0].scrollIntoView();", element)
        #         action.move_to_element(element).click().perform()
        #         time.sleep(1)
        #         # 2nd click on the Select-menu-outer option
        #         element = browser.find_element_by_id(f'react-select-{id}--option-{option}')
        #         action = ActionChains(browser)
        #         action.move_to_element(element).click().perform()
        #         time.sleep(1)
        #     def colorplus(browser):
        #         action = ActionChains(browser)
        #         element = browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[4]/div[1]/div/div[2]/div/div/div[1]/div/div[2]/div/div[3]/a')
        #         browser.execute_script("arguments[0].scrollIntoView();", element)
        #         action.move_to_element(element).click().perform()
        #         time.sleep(1)
        #     def sizetab(browser):
        #         action = ActionChains(browser)
        #         element = browser.find_element_by_xpath('//*[@id="react-tabs-2"]')
        #         browser.execute_script("arguments[0].scrollIntoView();", element)
        #         action.move_to_element(element).click().perform()
        #         time.sleep(1)
        #     def sizeplus(browser):
        #         action = ActionChains(browser)
        #         element = browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[4]/div[1]/div/div[2]/div/div/div[1]/div/div[3]/div/div/div[2]/div/div/div[3]/a')
        #         browser.execute_script("arguments[0].scrollIntoView();", element)
        #         action.move_to_element(element).click().perform()
        #         time.sleep(1)
        # # open file
        # with open('./constants/色.yml', 'r') as file:
        #     iro = yaml.load(file)
        # # 最初のカラーを入力
        # id = iro["color"][xl_sh.cell_value(yoko,8)]['id']
        # # select color type
        # actions.select(browser, Select.COLOR1.value, id)
        # # Translated color name
        # browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[4]/div[1]/div/div[2]/div/div/div[1]/div/div[2]/div/div[2]/table/tbody/tr[1]/td[3]/div/div/input').send_keys(xl_sh.cell_value(yoko,8))
        # time.sleep(1)
        # # 2色目以降ののカラーを入力
        # colorcell = [11,14,17,20,23,26,29,32,35]   
        # colornum = 2 
        # add = 0
        # for i in colorcell:
        #     colorcell = xl_sh.cell_value(yoko,i)
        #     colorName = xl_sh.cell_value(yoko,i - 1)
        #     if colorcell == '':
        #         break
        #     # Color Plus
        #     actions.colorplus(browser)
        #     #色アイコンの指定
        #     id = iro["color"][xl_sh.cell_value(yoko,i)]['id']
        #     # select color type
        #     actions.select(browser, Select.COLOR2.value + add, id)
        #     # Translated color name
        #     browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[4]/div[1]/div/div[2]/div/div/div[1]/div/div[2]/div/div[2]/table/tbody/tr['+ str(colornum) + ']/td[3]/div/div/input').send_keys(colorName)
        #     time.sleep(1)
        #     colornum += 1
        #     add += 1
        # # change to size section
        # actions.sizetab(browser)
        # # バリエーションあり
        # actions.select(browser, Select.COLOR2.value + add, 1)
        # # size plus
        # actions.sizeplus(browser)
        # actions.sizeplus(browser)
        # actions.sizeplus(browser) 
        # actions.sizeplus(browser)       
        # # clothes size
        # actions.select(browser, Select.COLOR2.value + add + 3, 0)
        # browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[4]/div[1]/div/div[2]/div/div/div[1]/div/div[3]/div/div/div[2]/div/div/div[2]/table/tbody/tr[1]/td[2]/div/div/div/input').send_keys('XS')
        # time.sleep(1)
        # actions.select(browser, Select.COLOR2.value + add + 4, 1)
        # browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[4]/div[1]/div/div[2]/div/div/div[1]/div/div[3]/div/div/div[2]/div/div/div[2]/table/tbody/tr[2]/td[2]/div/div/div/input').send_keys('S')
        # time.sleep(1)
        # actions.select(browser, Select.COLOR2.value + add + 5, 2)
        # browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[4]/div[1]/div/div[2]/div/div/div[1]/div/div[3]/div/div/div[2]/div/div/div[2]/table/tbody/tr[3]/td[2]/div/div/div/input').send_keys('M')
        # time.sleep(1)
        # actions.select(browser, Select.COLOR2.value + add + 6, 3)
        # browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[4]/div[1]/div/div[2]/div/div/div[1]/div/div[3]/div/div/div[2]/div/div/div[2]/table/tbody/tr[4]/td[2]/div/div/div/input').send_keys('L')
        # time.sleep(1)
        # actions.select(browser, Select.COLOR2.value + add + 7, 4)
        # browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[4]/div[1]/div/div[2]/div/div/div[1]/div/div[3]/div/div/div[2]/div/div/div[2]/table/tbody/tr[5]/td[2]/div/div/div/input').send_keys('XL')
        # time.sleep(3)


        # ### タグ
        # class actions(Enum):
        #     def tagclick(browser):
        #         action = ActionChains(browser)
        #         taglist_extra = []
        #         tagList = xl_sh.cell_value(yoko,41)
        #         tagList = tagList.split('、')
        #         for i in tagList:
        #             try:
        #                 tag = browser.find_element_by_xpath("//*[text()=" + "'" + i + "']")
        #                 browser.execute_script("arguments[0].scrollIntoView();", tag)
        #                 time.sleep(1)
        #                 action.move_to_element(tag).click().perform()
        #             except:
        #                 taglist_extra.append(i)
        #                 print(taglist_extra)
        #                 time.sleep(1)
        #         if not taglist_extra == []:
        #             actions.ListTagClick(browser,taglist_extra)
        #     def ListTagClick(browser,taglist_extra):
        #         action = ActionChains(browser)
        #         # open tag list
        #         element = browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[3]/div[5]/div/div[2]/div/div/p/a')
        #         browser.execute_script("arguments[0].scrollIntoView();", element)
        #         action.move_to_element(element).click().perform()
        #         time.sleep(3)
        #         for i in taglist_extra:
        #             browser.find_element_by_xpath("//*[text()=" + "'" + i + "']").click()
        #             time.sleep(3)
        #         actions.tagDone(browser)
        #     def tagDone(browser):
        #         action = ActionChains(browser)
        #         tagdone = browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[2]/div[7]/div/div/div[3]/button[2]')
        #         browser.execute_script("arguments[0].scrollIntoView();", tagdone)
        #         time.sleep(1)
        #         action.move_to_element(tagdone).click().perform()
        # actions.tagclick(browser)
        


        # #サイズ補足説明欄
        # category = xl_sh.cell_value(yoko,2)
        # if 'ホリスター' in brand:
        #     if 'メンズ' in category:
        #         sizeComment = "★メンズ参考サイズ★\n\n基本注文確認後、買い付けしています。手元に在庫がありませんので下記の参考サイズを参考にして頂ければ幸いです。\n\nサイズに不安な場合は、注文後、買い付けが完了しだい実寸の平置きをお知らせすることはできます。確認後のサイズ変更も可能です。\n\n（胸囲cm）\nS　91 - 96\nM　97 - 101\nL　102 - 106\nXL 107 - 111\n\n（袖長さcm）\nS　82 - 85\nM　86 - 87\nL　89 - 90\nXL 91 - 93\n\n\n（ウェスト）\nXS 28 (71cm)\nS　 29 - 30 (74-76cm)\nM　 31 - 32 (79-81cm)\nL　 33 - 34 (84-86cm)\nXL 36 (89cm)\n\n(足のサイズ）\nS　26.5ｃｍ\nM　27.5ｃｍ\nL　28.3ｃｍ\nXL　29.1ｃｍ\n\nその他アバクロ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816/\nその他ホリスター商品はこちら→https://www.buyma.com/r/_HOLLISTER-%E3%83%9B%E3%83%AA%E3%82%B9%E3%82%BF%E3%83%BC/-B4256816/\nその他アバクロキッズ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816F1/%E5%A4%A7%E4%BA%BA%E3%82%82OK/"
        #     elif 'レディース' in category:
        #         sizeComment = "★レディース参考サイズ★\n\n基本注文確認後、買い付けしています。手元に在庫がありませんので下記の参考サイズを参考にして頂ければ幸いです。\n\nサイズに不安な場合は、注文後、買い付けが完了しだい実寸の平置きをお知らせすることはできます。確認後のサイズ変更も可能です。\n\n（胸囲cm）\nXS　80 - 84　（5 - 7号）\nS　86 - 89　　（９号）\nM　91 - 94　　（１１号）\nL　96 - 97　　（１３号）\n\n\n（ウェスト INCHES）\nXS　23 - 25　（5 - 7号）\nS　26 - 27　 （7 - 9号）\nM　28 - 29　 （9 - 11号）\nL　30 - 31　 （11 - 13号）\n\n(足のサイズ）\nXS　23.2ｃｍ\nS　23.8ｃｍ\nM　24.8ｃｍ\nL　25.4ｃｍ\nXL　29.1ｃｍ\n\nその他アバクロ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816/\nその他ホリスター商品はこちら→https://www.buyma.com/r/_HOLLISTER-%E3%83%9B%E3%83%AA%E3%82%B9%E3%82%BF%E3%83%BC/-B4256816/\nその他アバクロキッズ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816F1/%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD%E3%82%AD%E3%83%83%E3%82%BA/"
        # elif 'アバクロ' in brand:
        #     if 'メンズ' in category:
        #         sizeComment = "★メンズ参考サイズ★\n\n基本注文確認後、買い付けしています。手元に在庫がありませんので下記の参考サイズを参考にして頂ければ幸いです。\n\nサイズに不安な場合は、注文後、買い付けが完了しだい実寸の平置きをお知らせすることはできます。確認後のサイズ変更も可能です。\n\n（胸囲cm）\nS　91 - 96\nM　97 - 101\nL　102 - 106\nXL 107 - 111\n\n（袖長さcm）\nS　82 - 85\nM　86 - 87\nL　89 - 90\nXL 91 - 93\n\n\n（ウェスト）\nXS 28 (71cm)\nS　 29 - 30 (74-76cm)\nM　 31 - 32 (79-81cm)\nL　 33 - 34 (84-86cm)\nXL 36 (89cm)\n\n(足のサイズ）\nS　26.5ｃｍ\nM　27.5ｃｍ\nL　28.3ｃｍ\nXL　29.1ｃｍ\n\nその他アバクロ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816/\nその他ホリスター商品はこちら→https://www.buyma.com/r/_HOLLISTER-%E3%83%9B%E3%83%AA%E3%82%B9%E3%82%BF%E3%83%BC/-B4256816/\nその他アバクロキッズ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816F1/%E5%A4%A7%E4%BA%BA%E3%82%82OK/"
        #     elif 'レディース' in category:
        #         sizeComment = "★レディース参考サイズ★\n\n基本注文確認後、買い付けしています。手元に在庫がありませんので下記の参考サイズを参考にして頂ければ幸いです。\n\nサイズに不安な場合は、注文後、買い付けが完了しだい実寸の平置きをお知らせすることはできます。確認後のサイズ変更も可能です。\n\n（胸囲cm）\nXS　80 - 84　（5 - 7号）\nS　86 - 89　　（９号）\nM　91 - 94　　（１１号）\nL　96 - 97　　（１３号）\n\n\n（ウェスト INCHES）\nXS　23 - 25　（5 - 7号）\nS　26 - 27　 （7 - 9号）\nM　28 - 29　 （9 - 11号）\nL　30 - 31　 （11 - 13号）\n\n(足のサイズ）\nXS　23.2ｃｍ\nS　23.8ｃｍ\nM　24.8ｃｍ\nL　25.4ｃｍ\nXL　29.1ｃｍ\n\nその他アバクロ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816/\nその他ホリスター商品はこちら→https://www.buyma.com/r/_HOLLISTER-%E3%83%9B%E3%83%AA%E3%82%B9%E3%82%BF%E3%83%BC/-B4256816/\nその他アバクロキッズ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816F1/%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD%E3%82%AD%E3%83%83%E3%82%BA/"
        #     elif 'ベビー・キッズ' in category:
        #         sizeComment = "★アバクロキッズ参考サイズ★\n\n基本注文確認後、買い付けしています。手元に在庫がありませんので下記の参考サイズを参考にして頂ければ幸いです。\n\n（身長cm）\n5/6　110 - 122\n7/8　122 - 135\n9/10　135 - 145\n11/12　145 - 152\n13/14　152 - 160\n15/16　160 - 165\n\n（胸囲ｃｍ）\n5/6　58 - 64\n7/8　64 - 69\n9/10　69 - 72\n11/12　72 - 76\n13/14　76 - 80\n15/16　80 - 84\n\n(足のサイズ）\n12/13　18.4ｃｍ\n1/2　20.3ｃｍ\n3/4　21.9ｃｍ\n5/6　23.5ｃｍ\n7/8　24.8ｃｍ\n\n※在庫の変動が激しいので、購入前に在庫確認をよろしくお願いします。\n\nその他アバクロ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816/\nその他ホリスター商品はこちら→https://www.buyma.com/r/_HOLLISTER-%E3%83%9B%E3%83%AA%E3%82%B9%E3%82%BF%E3%83%BC/-B4256816/\nその他アバクロキッズ商品はこちら→https://www.buyma.com/r/_%E3%82%A2%E3%83%90%E3%82%AF%E3%83%AD-ABERCROMBIE&FITCH/-B4256816F1/%E5%A4%A7%E4%BA%BA%E3%82%82OK/"
        # else:
        #         sizeComment = "※在庫の変動が激しいので在庫確認をよろしくお願いします。"
        # browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[4]/div[1]/div/div[2]/div/div/div[2]/div[1]/textarea').send_keys(sizeComment)
        # browser.implicitly_wait(40)
        # time.sleep(1)


        # ###数量
        # productNum = "2"
        # browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[4]/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div/div/div/input').send_keys(productNum)
        # browser.implicitly_wait(40)
        # time.sleep(1)


        # ###配送方法
        # class actions(Enum):
        #     def shippipng1(browser):
        #         action = ActionChains(browser)
        #         element = browser.find_element_by_xpath("//*[text()='USPS - First-Class Mail International（FCPIS含む）']")
        #         browser.execute_script("arguments[0].scrollIntoView();", element)
        #         time.sleep(1)
        #         action.move_to_element(element).click().perform()
        #         time.sleep(1)
        #     def shippipng2(browser):
        #         action = ActionChains(browser)
        #         element = browser.find_element_by_xpath("//*[text()='USPS - Priority Mail International']")
        #         browser.execute_script("arguments[0].scrollIntoView();", element)
        #         time.sleep(1)
        #         action.move_to_element(element).click().perform()
        #         time.sleep(1)
        # actions.shippipng1(browser)
        # actions.shippipng2(browser)


        '''
        ###購入期限
        today = datetime.date.today()
        changeDate = datetime.timedelta(days=60)
        changeDate = today + changeDate
        changeDate = str(changeDate)
        changeDate = changeDate.replace('-','/')
        twoweeks = browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[7]/div[1]/div/div[2]/div/div/div/div/div/div/div/div[1]/div/input')
        twoweeks.send_keys(changeDate)
        time.sleep(5)
        twoweeks.clear()
        time.sleep(5)
        twoweeks.send_keys(changeDate)
        browser.implicitly_wait(40)
        time.sleep(15)
        '''


        ###ショップ名
        shopName = xl_sh.cell_value(yoko,38)
        browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[7]/div[3]/div/div[2]/div/div/div/div/div[1]/input').send_keys(shopName)
        browser.implicitly_wait(40)
        time.sleep(1)


        ###値段
        if 'シャツ' in category or '帽子' in category  or '水着' in category or 'キッズ用トップス' in category  or '子供用帽子' in category:
            sellingPrice = (((float(xl_sh.cell_value(yoko,3)))*100*1.0685*1.2)+3200)*1.08
            sellingPrice = round(sellingPrice, -2)
            browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[8]/div[1]/div/div[2]/div/div/div[1]/div/div[1]/div/div/input').send_keys(int(sellingPrice))
            browser.implicitly_wait(40)
            time.sleep(1)
        elif 'デニム' in category:
            sellingPrice = (((float(xl_sh.cell_value(yoko,3)))*100*1.0685*1.2)+4500)*1.08
            sellingPrice = round(sellingPrice, -2)
            browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[8]/div[1]/div/div[2]/div/div/div[1]/div/div[1]/div/div/input').send_keys(int(sellingPrice))
            browser.implicitly_wait(40)
            time.sleep(1)
        else:
            sellingPrice = (((float(xl_sh.cell_value(yoko,3)))*100*1.0685*1.2)+4000)*1.08
            sellingPrice = round(sellingPrice, -2)
            browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[8]/div[1]/div/div[2]/div/div/div[1]/div/div[1]/div/div/input').send_keys(int(sellingPrice))
            browser.implicitly_wait(40)
            time.sleep(1)

        ###関税
        class actions(Enum):
            def kanzei(browser):
                action = ActionChains(browser)
                element = browser.find_element_by_xpath("//*[text()='関税込み (購入者の関税負担なし)']")
                browser.execute_script("arguments[0].scrollIntoView();", element)
                time.sleep(1)
                action.move_to_element(element).click().perform()
                time.sleep(1)
                element = browser.find_element_by_xpath("//*[text()='購入者申請時、全額負担']").click()
                time.sleep(1)
        actions.kanzei(browser)
        time.sleep(10)


        ###入力内容を確認ボタン
        class actions(Enum):
            def done(browser):
                action = ActionChains(browser)
                element = browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[1]/div/div/div/div[2]/form/div[10]/div/button[2]')
                browser.execute_script("arguments[0].scrollIntoView();", element)
                time.sleep(1)
                action.move_to_element(element).click().perform()
                browser.implicitly_wait(40)
                time.sleep(1)
        actions.done(browser)
        browser.implicitly_wait(40)
        time.sleep(3)

        ###出品完了
        class actions(Enum):
            def kanryo(browser):
                action = ActionChains(browser)
                element = browser.find_element_by_xpath('/html/body/div[3]/div[2]/div[1]/div/div[2]/div[4]/div/div/div[3]/button[2]')
                browser.execute_script("arguments[0].scrollIntoView();", element)
                time.sleep(1)
                action.move_to_element(element).click().perform()
                browser.implicitly_wait(40)
                time.sleep(1)
        actions.kanryo(browser)
        browser.implicitly_wait(40)
        time.sleep(3)

        # ###出品URLを保存###
        # ##################
        # browser.find_element_by_link_text('出品リストへ戻る').click()
        # browser.implicitly_wait(40)
        # time.sleep(1)
        # ItemURL = browser.find_element_by_xpath('//*[@id="inputform"]/table/tbody/tr[2]/td[4]/p[2]/a[1]').text
        # ItemURL = 'https://www.buyma.com/my/sell/new/?iid=' + ItemURL
        # browser.implicitly_wait(40)
        # time.sleep(1)
        # ws['AN' + str(int(yoko) + 1)].value = ItemURL
        # wb.save('C:\\Users\\tomoa\\Workspace\\buyma' + excelName + '.xlsx')

        ###出品画面に移動###
        ###################
        browser.get('https://www.buyma.com/my/sell/new/')
        browser.implicitly_wait(40)
        time.sleep(1)

        ###次の行に行く###
        #################
        yoko += 1
        folderNum = int(yoko) + 1

    ### 変更なし
    else:
        yoko += 1
        folderNum = int(yoko) + 1
        continue




    # except IndexError:
    #     break


    # except:
    #     try:
    #         ws['F' + str(int(yoko) + 1)].value = 'ERROR'
    #         wb.save('C:\\Users\\tomoa\\Workspace\\buyma' + excelName + '.xlsx')
    #         browser.get('https://www.buyma.com/my/sell/new/')
    #         browser.implicitly_wait(40)
    #         time.sleep(1)
    #         yoko += 1
    #         folderNum = yoko + 1
    #     except:
    #         time.sleep(1)
    #         Alert(browser).accept()
    #         ws['F' + str(int(yoko) + 1)].value = 'ERROR'
    #         wb.save('C:\\Users\\tomoa\\Workspace\\buyma' + excelName + '.xlsx')
    #         browser.get('https://www.buyma.com/my/sell/new/')
    #         browser.implicitly_wait(40)
    #         time.sleep(1)
    #         yoko += 1
    #         folderNum = yoko + 1
