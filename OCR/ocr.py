try:
    from PIL import Image
except ImportError:
    import Image
import pytesseract
import os
from docx import Document
import tkinter



path1=os.path.dirname(os.path.realpath(__file__))
os.chdir(path1+"\\")
list_pic=os.listdir()
print("当前文件夹下共有.jpg图片文件%s个"%str(list_pic).count(".jpg"))

try:
    pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
except Exception as e:
    labinfo.config(text="请核对tesseract.exe路径，C:\Program Files\Tesseract-OCR\tesseract.exe")
    root.update()
    raise Exception("路径错误")

#pytesseract.pytesseract.tesseract_cmd = r'tesseract.exe'
def pic_to_word(pic):
    try:
        tessdata_dir_config = '--tessdata-dir "c://Program Files//Tesseract-OCR//tessdata"'
        a=pytesseract.image_to_string(Image.open(pic),lang = 'chi_sim',config=tessdata_dir_config)
    except Exception as e:
        labinfo.config(text="请核对chi_sim语言包是否存在,C:\Program Files\Tesseract-OCR\tessdata\chi_sim.traineddata")
        root.update()
        print(e)
        raise Exception("语言包错误")
    #print(a)
    if "丫" in a:
        a=a.replace("丫","")
    a=a.replace(" ","")
    a=a.replace("_","")
    a=a.replace("\\","")
    b=a.replace('“',"")
    #print(b)
    c=b.split("\n")
    #print("c",c)
    mystr=""
    for i in c:
        if "A" in i or "B" in i or "C" in i or "D" in i:
            mystr += "\n"+"    "+i
        else:
            mystr += i
    #print("mystr",mystr)
    return True,mystr
def run():
    mystr=""   
    document = Document()
    count=0
    for i in list_pic:
        if ".jpg" in i:
            count += 1
            labinfo.config(text="当前执行文件为%s,第%s个文件/共%s个文件"%(i,str(count),str(list_pic).count(".jpg")))
            root.update()
            r,mystr=pic_to_word(i) 
            paragraph = document.add_paragraph(str(count)+"."+mystr)
            if count == str(list_pic).count(".jpg"):
                labinfo.config(text="%s个文件执行完成！"%str(count))
                root.update()
    document.save('myword.docx')


if __name__=='__main__':
    
    root=tkinter.Tk()
    root.title('Picture to Word V1.0')
    root.geometry('500x550')
    button_bg = '#D5E0EE'  
    button_active_bg = '#E5E35B'
    labinfo1=tkinter.Label(root,text="欢迎使用本工具！请将.jpg文件同工具文件放置在同一文件夹下，执行完成后会在工具文件夹下生成myword.docx文件",bg='green',fg='yellow',wraplength=300)
    labinfo1.place(relx=0.5,y=100,anchor="center")
    labinfo=tkinter.Label(root,text="",wraplength=400)
    labinfo.place(relx=0.5,y=200,anchor="center")
    btnexe=tkinter.Button(root,text="开始处理图片",bg=button_bg,activebackground = button_active_bg,command=lambda:run())
    btnexe.place(relx=0.5,y=300,anchor="center")
   
    
    root.mainloop()

