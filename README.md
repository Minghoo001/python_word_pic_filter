使用spire.doc解析word文件，按节读取识别  ps 如果使用FixedLayoutDocument按页进行读取容易读到重复的两行图片不建议使用
tkinter界面
相似度匹配
   vgg16  最优
   sift特征  最优
   单通道直方图和三直方图检索  优
   BOW算法检索  不稳定
   哈希算法检索  不稳定
   orb特征检索   最差
参考https://zhuanlan.zhihu.com/p/545851284
本代码基于 ：单通道直方图和三直方图

def calculate_single(img1, img2):
    hist1 = cv2.calcHist([img1], [0], None, [256], [0.0, 255.0])
    hist1 = cv2.normalize(hist1, hist1, 0, 1, cv2.NORM_MINMAX, -1)
    hist2 = cv2.calcHist([img2], [0], None, [256], [0.0, 255.0])
    hist2 = cv2.normalize(hist2, hist2, 0, 1, cv2.NORM_MINMAX, -1)

degree = 0
for i in range(len(hist1)):
if hist1[i] != hist2[i]:
            degree = degree + (1 - abs(hist1[i] - hist2[i]) / max(hist1[i], hist2[i]))
else:
            degree = degree + 1
degree = degree / len(hist1)
return degree
def classify_hist_of_three(img1, img2, size=(256, 256)):
image1 = cv2.resize(img1, size)
    image2 = cv2.resize(img2, size)
    sub_image1 = cv2.split(img1)
    sub_image2 = cv2.split(img2)
    sub_data = 0
for im1, im2 in zip(sub_img1, sub_img2):
        sub_data += calculate_single(im1, im2)
    sub_data = sub_data / 3
return sub_data
