import fitz

import re

import openpyxl as xl


def reading(pageNum):
    page=f[pageNum]#[434]
    rect=page.rect
    clip=fitz.Rect(0,0, rect.width, rect.height)
    a_text=page.get_text(clip=clip)
    col_dict=a_text.split('\n')
    print(a_text)
    #print(col_dict)

    page_lst=[]


if __name__ == '__main__':
    f = fitz.open('import_pdf.pdf')

    for pageNum in range(518, 537):#(22, 23): #701
        page_lst=reading(pageNum=pageNum)

        print(pageNum+1,'!--------------------------------------------------------------!')

    f.close()