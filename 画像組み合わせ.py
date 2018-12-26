import cv2, os, re, xlrd

#一枚目と二枚目の画像組み合わせ
excelName = input('Input Excel Name：')
folderNum = input('開始フォルダ番号')
yoko = str(int(folderNum) - 1)
endfolderNum = input('終了フォルダ番号')
endfolderNum = int(endfolderNum) + 1
firstNum = input('組み合わせ画像番号')
name = excelName
number = folderNum
xl_bk = xlrd.open_workbook(excelName + '.xlsx')
xl_sh = xl_bk.sheet_by_name(excelName)

while True:
    if int(folderNum) == int(endfolderNum):
        break
    try:
    
        brandName = xl_sh.cell_value(int(yoko),38)
        print(brandName)

        if brandName == 'adidas(アディダス)':
            brand = 'adidas'
        elif brandName == 'Hollister Co.(ホリスター)':
            brand = 'hollister'
        elif brandName == 'Alpha Industries(アルファインダストリー)':
            brand = 'alpha industries'
        elif brandName == 'Calvin Klein(カルバンクライン)':
            brand = 'calviin klein'
        elif brandName == 'KAPTEN & SON(キャプテン＆サン)':
            brand = 'capten&son'
        elif brandName == 'CHAMPION(チャンピオン)':
            brand = 'champion'
        elif brandName == 'CONVERSE(コンバース)':
            brand = 'converse'
        elif brandName == 'Dickies(ディッキーズ)':
            brand = 'dickies'
        elif brandName == 'Dr Martens(ドクターマーチン)':
            brand = 'dr.martens'
        elif brandName == 'FILA(フィラ)':
            brand = 'fila'
        elif brandName == 'Guess(ゲス)':
            brand = 'guess'
        elif brandName == 'Herschel Supply(ハーシェルサプライ)':
            brand = 'harchel'
        elif brandName == 'Kappa(カッパ)':
            brand = 'kappa'
        elif brandName == 'LACOSTE(ラコステ)':
            brand = 'lacoste'
        elif brandName == "Levi's(リーバイス)":
            brand = 'levi_s'
        elif brandName == "New Balance(ニューバランス)":
            brand = 'new balance'
        elif brandName == "Nike(ナイキ)":
            brand = 'nike'
        elif brandName == "Patagonia(パタゴニア)":
            brand = 'patagonia'
        elif brandName == "PUMA(プーマ)":
            brand = 'puma'
        elif brandName == "LAUREN RALPH LAUREN(ローレンラルフローレン)":
            brand = 'ralph lauren'
        elif brandName == "Reebok(リーボック)":
            brand = 'reebok'
        elif brandName == "ROCKET DOG(ロケットドッグ)":
            brand = 'rocket dog'
        elif brandName == "STUSSY(ステューシー)":
            brand = 'stussy'
        elif brandName == "Timberland(ティンバーランド)":
            brand = 'timberland'
        elif brandName == "Tommy Hilfiger(トミーヒルフィガー)":
            brand = 'tommy hilfiger'
        elif brandName == "UMBRO(アンブロ)":
            brand = 'umbro'
        elif brandName == "VANS(バンズ)":
            brand = 'vans'
        elif brandName == "SKECHERS(スケッチャーズ)":
            brand = 'skechers'
        elif brandName == "Urban Outfitters(アーバンアウトフィッターズ)":
            brand = 'urban'
        elif brandName == "Anthropologie(アンソロポロジー)":
            brand = 'anthropologie'
        elif brandName == "Free People(フリーピープル)":
            brand = 'freepeople'          
        elif brandName == "Abercrombie & Fitch(アバクロ)":
            brand = 'abercrombie'          

        else:
            brand = 'urban'
        print(number)
        print(brand)
        

        dir = "C:\\Users\\tomoa\\Desktop\\" + name + '\\' + str(folderNum)
        files = os.listdir(dir)# ファイルのリストを取得
        count = 0# カウンタの初期化
        for file in files:# ファイルの数だけループ
            index = re.search('.jpg', file)# 拡張子がjpgのものを検出
            if index:# jpgの時だけ（今回の場合は）カウンタをカウントアップ
                count = count + 1
        if count >= 2:
            folder = 'C:\\Users\\tomoa\\Desktop\\' + name + '\\' + str(folderNum) + '\\'
            img1 = cv2.imread(folder + str(firstNum) + ".jpg")
            img2 = cv2.imread("C:\\Users\\tomoa\\Desktop\\logos\\" + brand + '.jpg')
            img3 = cv2.hconcat([img1, img2])
            cv2.imwrite(folder + '0.jpg', img3)
        elif count == 1:
            folder = 'C:\\Users\\tomoa\\Desktop\\' + name + '\\' + str(folderNum) + '\\'
            img1 = cv2.imread(folder + str(firstNum) + ".jpg")
            img2 = cv2.imread("C:\\Users\\tomoa\\Desktop\\logos\\" + brand + '.jpg')
            img3 = cv2.hconcat([img1, img2])
            cv2.imwrite(folder + '0.jpg', img3)

        elif count == 0:
            folderNum = int(folderNum) + 1
            yoko = int(yoko) + 1
            number = int(number) + 1
            continue

        folderNum = int(folderNum) + 1
        yoko = int(yoko) + 1
        number = int(number) + 1


    except IndexError:
        break
