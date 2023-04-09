import fitz
import os


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
    print(pdf)
    name = pdf.name
    name = name.replace('pdfs/', '').replace('.pdf', '')
    # 逐页读取PDF
    for pg in range(0, len(pdf)):
        page = pdf[pg]
        # 设置缩放和旋转系数
        trans = fitz.Matrix(zoom_x, zoom_y)
        pm = page.get_pixmap(matrix=trans, alpha=False)
        # 开始写图像
        pm.save(imgPath + name + '-' + str(pg + 1) + ".png")
    pdf.close()


# pdf_image(r"pdfs/01.pdf", r"images/", 10, 10, 0)

file_dir = r'pdfs/'
file_list = []
for items in os.walk(file_dir, topdown=False):
    file_list = items[2]
print(file_list)

for file in file_list:
    head = 'pdfs/'
    pdf_image(head + file, r"images/", 5, 5, 0)