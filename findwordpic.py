from tkinter.scrolledtext import ScrolledText

import cv2
import numpy
from spire.doc import *
import os
import tkinter as tk

def classify_hist_with_split(image1, image2, size=(256, 256)):
    image1 = cv2.resize(image1, size)
    image2 = cv2.resize(image2, size)
    sub_image1 = cv2.split(image1)
    sub_image2 = cv2.split(image2)
    sub_data = 0
    for im1, im2 in zip(sub_image1, sub_image2):
        sub_data += caculate(im1, im2)
    sub_data = sub_data / 3
    return sub_data


def caculate(image1, image2):
    hist1 = cv2.calcHist([image1], [0], None, [256], [0.0, 255.0])
    hist2 = cv2.calcHist([image2], [0], None, [256], [0.0, 255.0])
    degree = 0
    for i in range(len(hist1)):
        if hist1[i] != hist2[i]:
            degree = degree + (1 - abs(hist1[i] - hist2[i]) / max(hist1[i], hist2[i]))
        else:
            degree = degree + 1
    degree = degree / len(hist1)
    return degree


def process():
    global input_path
    filenames = [os.path.join(input_path, i) for i in os.listdir(input_path)]
    for i, filename in enumerate(filenames):
        if not filename.endswith('.docx'):  # 只处理.docx文件
            continue
        document = Document()
        document.LoadFromFile(filename)
        images = []
        file_names = []
        cv_images = []
        for section in range(document.Sections.Count):#第几节
            sectObj = document.Sections.get_Item(section)
            if sectObj.Body.ChildObjects is None:
                continue
            for i in range(sectObj.Body.ChildObjects.Count):
                docobj = sectObj.Body.ChildObjects[i]
                if isinstance(docobj, Paragraph):
                    for j in range(docobj.ChildObjects.Count):
                        obj = docobj.ChildObjects[j]
                        if isinstance(obj, DocPicture):
                            print(f"Image-{section}-{i}-{j}.png")
                            picture = obj
                            # 将图片数据添加到列表中
                            data_bytes = picture.ImageBytes
                            images.append(data_bytes)
                            lastfilename = f"Image-sec{section}-{i}-j{j}.png"
                            file_names.append(lastfilename)
                            # 读取目标图片和模板图片
                            image_array = numpy.frombuffer(data_bytes, numpy.uint8)
                            template = cv2.imdecode(image_array, cv2.IMREAD_COLOR)

                            h, w = template.shape[:2]
                            if i == 51:
                                print()
                            for i, img_rgb in enumerate(cv_images):

                                if img_rgb is None:
                                    print("")
                                if template is None:
                                    print("")
                                # min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
                                max_val = classify_hist_with_split(img_rgb, template)
                                if (max_val > 0.65):
                                    print(f'图片{lastfilename}与{file_names[i]}图片存在相似度为{max_val}')
                                    text = f'图片{lastfilename}与{file_names[i]}图片存在相似度为{max_val}\n'
                                    scrolled_text.insert(tk.END, text)
                            cv_images.append(template)
        name, file_extension = os.path.splitext(filename)
        # 将图片数据保存为图像文件
        output_path = f'{name}_image'
        os.makedirs(output_path, exist_ok=True)
        for i, image_data in enumerate(images):
            # file_name = f"Image-{i}.png"
            with open(os.path.join(output_path, file_names[i]), 'wb') as image_file:
                image_file.write(image_data)

        document.Close()


# 初始化输入路径和标志
input_path = "/path/to/your/directory"
show_text_box = False
img = numpy.zeros((200, 400, 3), dtype=numpy.uint8)
frame = numpy.zeros((400, 500, 3), numpy.uint8)
input_text = ""
input_state = False

def get_input():
    global input_path
    input_path = entry.get()
    print("你输入的内容是: " + entry.get())
    process()


# 创建主窗口
root = tk.Tk()
root.title("识别文档中相同图片，相似度70%以上：")

label = tk.Label(root, text="请输入文件所在文件夹:")
label.pack()

entry = tk.Entry(root)
entry.pack()

# 创建一个按钮，点击后获取输入框的内容
button = tk.Button(root, text="识别", command=get_input)
button.pack()
# text_widget = tk.Text(root, width=100, height=50)
# text_widget.pack()

scrolled_text = ScrolledText(root, width=80, height=30)
scrolled_text.pack(expand=True, fill="both")

# 启动事件循环
root.mainloop()
cv2.destroyAllWindows()

# def cv2_template(image1,image2):
# methods = ['cv2.TM_CCOEFF_NORMED', 'cv2.TM_CCOEFF_NORMED']
# for method in methods:
#     meth = eval(method)
#     if(img_rgb.size < template.size):
#         w, h = img_rgb.shape[:2]
#         w1,h1 = template. shape[:2]
#
#         res = cv2.matchTemplate(template, img_rgb, cv2.TM_CCOEFF_NORMED)
#     else:
#         w, h = img_rgb.shape[:2]
#         w1, h1 = template.shape[:2]
#         if(w < w1 or h < h1):
#             resize = min(w, h)
#             resize1 = max(w1, h1)
#             scale = resize / resize1
#             template = cv2.resize(template,(0,0),fx=scale,fy=scale)
#         res = cv2.matchTemplate(img_rgb, template, cv2.TM_CCOEFF_NORMED)

# layoutDoc = FixedLayoutDocument(document)
# 访问文档的第一页
# for layoutPage in layoutDoc.Pages:
#     if layoutPage is None:
#         break
#     page+=1
#     print(f"page====is{page}")
#     # 遍历第一页第一列中的每一行
#     col = 0
#     column = layoutPage.Columns[0]
#     childObjs = None
#     print(f"col====is{col}")
#     for i in range(layoutPage.Columns[0].Lines.Count):
#         if(page < 7):
#             print(f"i====is{column.Lines.Count}")
#             print(f"i====is{column.Lines[0].Paragraph.ChildObjects.Count}")
#             print(f"i====is{column.Lines[1].Paragraph.ChildObjects.Count}")
#             print(f"i====is{column.Lines[2].Paragraph.ChildObjects.Count}")
#             print(f"i====is{column.Lines[3].Paragraph.ChildObjects.Count}")
#             line0 = column.Lines[0]
#             line1 = column.Lines[1]
#
#             line = layoutPage.Columns[0].Lines[i]
#             # 获取当前行的段落
#             line.Paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
#             childObjs = line.Paragraph.ChildObjects
#             # 遍历段落中的子对象
#             for j in range(line.Paragraph.ChildObjects.Count):
#                 if (childObjs[j] == line.Paragraph.ChildObjects[j]):
#                     continue
#                 print(f"j====is{j}")
#                 obj = line.Paragraph.ChildObjects[j]
#                 # if obj.
#                 # 查找图片
#                 if isinstance(obj, DocPicture):
#                     print(f"Image-{page}-{i}-{j}.png")
#                     picture = obj
#                     # 将图片数据添加到列表中
#                     data_bytes = picture.ImageBytes
#                     images.append(data_bytes)
#                     file_names.append(f"Image-p{page}-i{i}-j{j}.png")
