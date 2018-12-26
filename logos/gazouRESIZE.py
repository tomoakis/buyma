#  画像のリサイズ

import cv2
import numpy as np
imagename = input('画像名を入力：')

img = cv2.imread(imagename + '.jpg', cv2.IMREAD_COLOR)

# cv2.resize(画像, (幅, 高さ))
twiceImg = cv2.resize(img, (300, 1000))

cv2.imwrite(imagename + '.jpg', twiceImg)
