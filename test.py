import cv2
import numpy as np

# 读取大图和小图
img_big = cv2.imread('C:\\Users\\mail\\Documents\\GitHub\\pdfmanage\\images\\003-230103-0009-1.png')
img_small = cv2.imread('C:\\Users\\mail\\Documents\\GitHub\\pdfmanage\\search.jpg')

# 使用 TM_CCOEFF_NORMED 匹配算法进行模板匹配
result = cv2.matchTemplate(img_big, img_small, cv2.TM_CCOEFF_NORMED)

# 找到最佳匹配结果的坐标
min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
top_left = max_loc
bottom_right = (top_left[0] + img_small.shape[1], top_left[1] + img_small.shape[0])

# 在大图中标出匹配区域
cv2.rectangle(img_big, top_left, bottom_right, (0, 0, 255), 2)

# 显示结果图像
cv2.imshow('result', img_big)
cv2.waitKey(0)
cv2.destroyAllWindows()