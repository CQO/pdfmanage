import fitz
import os
import re
import cv2
import json
from PIL import Image
import numpy as np
from tencentcloud.common import credential
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.ocr.v20181119 import ocr_client, models
import base64
import shutil
import time

if os.path.exists('temp'):
    shutil.rmtree('temp')

os.mkdir("temp")


def image_to_pdf(input_path, output_path):
    img = Image.open(input_path)
    pdf_path = output_path
    img.save(pdf_path, "PDF", resolution=72.0)

# if not os.path.exists('output'):
#     os.mkdir("output")

def saveToPdf(image_file, savePath, width, height):

    # 新建PDF文件
    pdf_file = fitz.open()

    # 新建一页
    page = pdf_file.new_page(width=width, height=height)

    # 将图像文件插入到页面中
    img_rect = fitz.Rect(0, 0, page.mediabox[2], page.mediabox[3])
    page.insert_image(img_rect, filename=image_file)

    # 保存PDF文件
    pdf_file.save(savePath)
    pdf_file.close()

def shibie (filePath):
    try:
        # 用你的ID和KEY替换掉SecretId、SecretKey
        cred = credential.Credential("AKID62ub6KoNnDWkz50ymMq58mQxTp0161mO", "*")
        httpProfile = HttpProfile()
        httpProfile.endpoint = "ocr.tencentcloudapi.com"
        # 使用TC3-HMAC-SHA256加密方法，不使用可能会报错
        clientProfile = ClientProfile("TC3-HMAC-SHA256")
        clientProfile.httpProfile = httpProfile
        # 按就近的使用，所以我用的是ap-shanghai
        client = ocr_client.OcrClient(cred, "ap-shanghai", clientProfile)

        req = models.GeneralBasicOCRRequest()
        # 将本地文件转换为ImageBase64
        with open(filePath, 'rb') as f:
            base64_data = base64.b64encode(f.read())
            s = base64_data.decode()
        params = '{"ImageBase64":"%s"}' % s
        req.from_json_string(params)

        resp = client.GeneralBasicOCR(req)
        resp = resp.to_json_string()
        # 将官网文档里输出字符串格式的转换为字典，如果不需要可以直接print(resp)
        resp = json.loads(resp)
        # 下面都是从字典中取出识别出的文本内容，不需要其他的参数内容
        resp_list = resp['TextDetections']
        
        ind = 0
        textTemp = ''
        str0 = ''
        str2 = ''
        str1 = 'A'
        for resp in resp_list:
            match = re.findall('[1-9]/[1-9]', resp['DetectedText'])
            if len(match) >= 1:
                str2 = match[0].replace('/', '-')
                str2 = '-' + str2.split('-')[0]
                if (int(str2) == 1):
                    str2 = ''
            if 'A' in resp['DetectedText']:
                str1 = '-A'
            if 'B' in resp['DetectedText']:
                str1 = '-B'
        for resp in resp_list:
            result = resp['DetectedText']
            textTemp += result
            
            # print(result)
            # if ((len(result.split('.')) > 3 or '-' in result) and ':' not in result and len(result) > 12 and '/' not in result and '\\' not in result):
            if (('RM' in result and '-' in result) or ('ME' in result and '-' in result) or ('TF' in result and '.' in result) or (result.count('.') == 4)):
                result = result.replace(')', '1')
                result = result.replace('图', '')
                result = result.replace('号', '')
                result = result.replace(' ', '|')
                
                
                result = result.split('|')
                for item in result:
                    if ('ME' in item and '-' in item):
                        str0 = item
                    if ('RM' in item and '-' in item):
                        str0 = item
                    if ('TF' in item and '.' in item):
                        str0 = item
                    if (item.count('.') == 4):
                        str0 = item
                        # print([str0, str1, str2])
            ind += 1
        # print('无识别结果')
        # textTemp.replace('\r', '')
        # textTemp.replace('\n', '')
        # textTemp.replace(' ', '')
        # match = re.search(r'.............ME...-.|..............ME...-..', textTemp)
        # print(textTemp)
        # if (match):
        #     temp = match.group(0)
        #     temp = temp.replace('/', '-')
        #     temp = temp.replace(')', '1')
        #     return [temp, '', '']
        if ('A' in str0):
            str0 = str0.split('A')[0]
        if ('B' in str0):
            str0 = str0.split('B')[0]
        print([str0, str2, str1])
        return [str0, str2, str1]
    except TencentCloudSDKException as err:
        # print('无识别结果')
        return [None]

def changeImage(img, angle):
    # 图片中心点坐标
    cx, cy = img.shape[1] / 2, img.shape[0] / 2
    if (angle == 90):
        img = cv2.transpose(img)
        img = cv2.flip(img, flipCode=1)

        # 显示旋转后的图片
        # cv2.imshow('Rotated Image', img)
        # cv2.waitKey(0)
        # cv2.destroyAllWindows()
        return img
    # 获取旋转矩阵
    M = cv2.getRotationMatrix2D((cx, cy), angle, 1)
    
    # 进行旋转
    rotated_img = cv2.warpAffine(img, M, (img.shape[1], img.shape[0]))
    return rotated_img

def mkdir(path):
    folder = os.path.exists(path)
    if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)  # makedirs 创建文件时如果路径不存在会创建这个路径
    else:
        pass


def pdf_image(pdfPath, imgPath, zoom_x, zoom_y, rotation_angle):
    """
    :param pdfPath: pdf文件的路径
    :param imgPath: 图像要保存的文件夹
    :param zoom_x: x方向的缩放系数
    :param zoom_y: y方向的缩放系数
    :param rotation_angle: 旋转角度
    :return: None
    """
    # 打开PDF文件
    pdf = fitz.open(pdfPath)
    # print(pdf)
    name = pdf.name
    name = name.replace('图纸/', '').replace('.pdf', '')
    
    # 逐页读取PDF
    for pg in range(0, len(pdf)):
        page = pdf[pg]
        # 设置缩放和旋转系数
        trans = fitz.Matrix(zoom_x, zoom_y)
        pm = page.get_pixmap(matrix=trans, alpha=False)
        # 开始写图像
        # imgPath = imgPath + name + '-' + str(pg + 1) + ".png"
        imgPath = imgPath + name + ".png"
        pm.save(imgPath)
        
    pdf.close()

    

# pdf_image(r"图纸/01.pdf", r"images/", 10, 10, 0)

file_dir = r'图纸/'
file_list = []
for filename in os.listdir(file_dir):
    # 对文件名进行解码，避免中文乱码
    filename = filename.encode('utf-8').decode('utf-8')
    file_list.append(filename)
print(file_list)

indTemp = 0
for file in file_list:
    indTemp += 1
    print('PDF文件处理进度: %s/%s' % (indTemp, len(file_list)))
    head = '图纸/'
    if ('.pdf' in file):
        pdf_image(head + file, r"temp/", 5, 5, 0)

images_dir = r'temp/'
images_list = []
for filename in os.listdir(images_dir):
    # 对文件名进行解码，避免中文乱码
    filename = filename.encode('utf-8').decode('utf-8')
    images_list.append(filename)

#读取图像，解决imread不能读取中文路径路径的问题
def cv_imread(file_path):
    #imdedcode读取的是RGB图像
    cv_img = cv2.imdecode(np.fromfile(file_path,dtype=np.uint8),-1)
    return cv_img

print(images_list)
indTemp = 0
# 先把图片旋转90°
for file in images_list:
    indTemp += 1
    print('调整图片方向: %s/%s' % (indTemp, len(images_list)))
    imgPath = '.\\temp\\' + file
    imgPath = imgPath.encode('utf-8').decode('utf-8')
    img_big = cv_imread(imgPath)
    height, width, channels = img_big.shape
    if (width < height):
        # print('旋转图片:' + imgPath)
        img_big = changeImage(img_big, 90)
        cv2.imwrite(imgPath, img_big)
# 修改图片名称
# 读取大图和小图
# 逐页读取PDF
indTemp = 0
for file in images_list:
    indTemp += 1
    print('图片处理进度: %s/%s' % (indTemp, len(images_list)))
    imgPath = '.\\temp\\' + file
    imgPathTemp = '.\\temp\\temp_' + file
    # print(imgPath)
    # 判断文件是否存在
    if os.path.exists(imgPath):
        img_big = cv_imread(imgPath)
        height, width, channels = img_big.shape
        img_bigCop = img_big[int(height * 0.7):height, int(width * 0.45):width]
        cv2.imencode('.png', img_bigCop)[1].tofile(imgPathTemp)
        
        # 判断文件是否存在
        if os.path.exists(imgPathTemp):
            resImg = shibie(imgPathTemp)
            # print(resImg)
            if (not resImg[0]):
                img_big = changeImage(img_big, 180)
                height, width, channels = img_big.shape
                img_bigCop = img_big[int(height * 0.7):height, int(width * 0.45):width]
                cv2.imencode('.png', img_bigCop)[1].tofile(imgPathTemp)
                resImg = shibie(imgPathTemp)
            if (not resImg[0]):
                img_big = changeImage(img_big, 90)
                height, width, channels = img_big.shape
                img_bigCop = img_big[int(height * 0.7):height, int(width * 0.45):width]
                cv2.imencode('.png', img_bigCop)[1].tofile(imgPathTemp)
                resImg = shibie(imgPathTemp)
            if (not resImg[0]):
                img_big = changeImage(img_big, 180)
                height, width, channels = img_big.shape
                img_bigCop = img_big[int(height * 0.7):height, int(width * 0.45):width]
                cv2.imencode('.png', img_bigCop)[1].tofile(imgPathTemp)
                resImg = shibie(imgPathTemp)
            if (resImg[0] and 'sworkbel' not in resImg):
                if (resImg[1] == '-1'):
                    resImg[1] = ''
                # outFileName = './output/' + resImg[0] +  resImg[1] + resImg[2] + '.pdf'
                # if os.path.exists(outFileName):
                #     print('文件已存在:' + outFileName)
                # else:
                #     cv2.imwrite('./temp/' + file, img_big)
                #     image_to_pdf('./temp/' + file, outFileName, width, height)
                
                outFileName = './图纸/' + resImg[0] +  resImg[1] + resImg[2] + '.pdf'
                if os.path.exists(outFileName):
                    print('文件已存在:' + outFileName)
                else:
                    cv2.imwrite('./图纸/' + resImg[0] +  resImg[1] + resImg[2] + '.png', img_big)
                    os.rename('./图纸/' + file.replace('png', 'pdf'), outFileName)
        else:
            print('文件不存在:' + imgPathTemp)
    else:
        print('文件不存在:' + imgPath)
shutil.rmtree('temp')
input("按任意键退出")