import os
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.style import WD_STYLE_TYPE
import argparse

# 获取工程目录
# 遍历目录获得指定后缀的文件
# str = 读取文件字符串
# document.add(str)

# 标题 ==== xxx.java ===
# 页码

# 去掉空行
# 去掉注释

parser = argparse.ArgumentParser()
parser.add_argument('-p', dest='project_path', help='the absolutely path of your project')
parser.add_argument('-t', dest='file_type', action='append', help='the file types your project contains')
args = parser.parse_args()

PATH = args.project_path
TYPES = args.file_type

if __name__ == '__main__':
    if PATH is None or not os.path.exists(PATH) or not os.path.isabs(PATH) or os.path.isfile(PATH):
        print("plz input an absolutely path of your project")
        exit()
    if TYPES is None or len(TYPES) == 0:
        print("plz input the file types your project contains")
        exit()
    list_dirs = os.walk(PATH)
    document = Document()
    document.add_heading('code', 0)
    styles = document.styles
    style = styles.add_style('CodeContent', WD_STYLE_TYPE.PARAGRAPH)
    style.base_style = styles['Normal']
    font = style.font
    font.name = '宋体'
    font.size = Pt(10.5)  # 10.5 五号
    paragraph_format = style.paragraph_format
    paragraph_format.line_spacing = Pt(18)  # 单倍行距 行间距18

    for root, dirs, files in list_dirs:
        for f in files:
            if os.path.splitext(f)[1].replace('.', '') in TYPES:
                abs_path = os.path.join(root, f)
                fp = open(abs_path)
                content = fp.read()
                paragraph = document.add_paragraph(content, 'CodeContent')
                document.add_page_break()

    document.save('demo1.docx')
