import cv2
import json
import numpy as np
from tencentcloud.common import credential
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.ocr.v20181119 import ocr_client, models
import base64

def shibie (filePath):
    try:
        # 用你的ID和KEY替换掉SecretId、SecretKey
        cred = credential.Credential("AKIDKzOWO0EnmUzEBqoRp54okQwwhHadIYhZ", "lHXbCZtw83ScV6RSCfrnY3xaeSzFFcjZ")
        httpProfile = HttpProfile()
        httpProfile.endpoint = "ocr.tencentcloudapi.com"
        # 使用TC3-HMAC-SHA256加密方法，不使用可能会报错
        clientProfile = ClientProfile("TC3-HMAC-SHA256")
        clientProfile.httpProfile = httpProfile
        # 按就近的使用，所以我用的是ap-shanghai
        client = ocr_client.OcrClient(cred, "ap-shanghai", clientProfile)

        req = models.GeneralAccurateOCRRequest()
        # 将本地文件转换为ImageBase64
        with open(filePath, 'rb') as f:
            base64_data = base64.b64encode(f.read())
            s = base64_data.decode()
        params = '{"ImageBase64":"%s"}' % s
        req.from_json_string(params)

        resp = client.GeneralAccurateOCR(req)
        resp = resp.to_json_string()
        # 将官网文档里输出字符串格式的转换为字典，如果不需要可以直接print(resp)
        resp = json.loads(resp)
        # 下面都是从字典中取出识别出的文本内容，不需要其他的参数内容
        resp_list = resp['TextDetections']
        for resp in resp_list:
            result = resp['DetectedText']
            if ('|A|' in result):
                return result
    except TencentCloudSDKException as err:
        print(err)


# 读取大图和小图
img_big = cv2.imread('C:\\Users\\mail\\Documents\\GitHub\\pdfmanage\\images\\003-230103-0009-1.png')
img_small = cv2.imread('C:\\Users\\mail\\Documents\\GitHub\\pdfmanage\\search.jpg')

height, width, channels = img_big.shape
print(width)

def changeImage(img, angle):
    # 图片中心点坐标
    cx, cy = img.shape[1] / 2, img.shape[0] / 2

    # 获取旋转矩阵
    M = cv2.getRotationMatrix2D((cx, cy), angle, 1)
    # 进行旋转
    rotated_img = cv2.warpAffine(img, M, (img.shape[1], img.shape[0]))
    return rotated_img
if (width < height):
    img_big = changeImage(img_big, 90)
# img_big = changeImage(img_big, 180)
# 裁切图片
widthCut = width- 1400
img_bigCop = img_big[height-480:height, widthCut:width]

cv2.imwrite("./cv_cut_thor.jpg", img_bigCop)
resImg = shibie("./cv_cut_thor.jpg")
# print(resImg)
if (not resImg):
    img_big = changeImage(img_big, 180)
    img_bigCop = img_big[height-480:height, widthCut:width]
    cv2.imwrite("./cv_cut_thor.jpg", img_bigCop)
    resImg = shibie("./cv_cut_thor.jpg")

print(resImg)
# maxTop = 0
# maxLeft = 0

# # 使用 TM_CCOEFF_NORMED 匹配算法进行模板匹配
# result = cv2.matchTemplate(img_big, img_small, cv2.TM_CCOEFF_NORMED)

# # 找到最佳匹配结果的坐标
# min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
# top_left = max_loc
# bottom_right = (top_left[0] + img_small.shape[1], top_left[1] + img_small.shape[0])

# if (top_left[0] > maxTop):
#     maxTop = top_left[0]
# if (top_left[1] > maxLeft):
#     maxLeft = top_left[1]

# print(maxTop, maxLeft)

# # 在大图中标出匹配区域
# cv2.rectangle(img_big, top_left, bottom_right, (0, 0, 255), 2)

# # 显示结果图像
# cv2.imshow('result', img_big)
# cv2.waitKey(0)
# cv2.destroyAllWindows()