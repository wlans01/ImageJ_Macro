import os
import pandas as pd
import openpyxl
from pywinauto.application import Application
import pyautogui as pg
import time
import pyperclip
from openpyxl.drawing.image import Image
import shutil

class Data:
    def __init__(self):
        self.abspath = os.path.dirname(os.path.abspath(__file__)).replace('ImageJ_Macro','')
        self.data_path = self.abspath+"data\\"
        self.temp_path = self.abspath+"temp\\"
        self.file_list_bmp = []
        self.file_list_txt = []
        self.folder_list =[]
        self.data = None

    def run(self):
        # 데이터 파일안에 폴더 갯수 확인
        self.folder_list= self.folder_listup()
        # 갯수만큼 temp,result에 파일생성
        for i in self.folder_list:
            self.createFolder(self.abspath+"temp\\"+i+"\\")
            self.createFolder(self.abspath+"result\\"+i+"\\")

        # Imagej 매크로 실행
        # temp안에 데이터가 있는경우 다음단계 진행
        folder=self.folder_list[-1]
        path = self.abspath+"temp\\"+folder+"\\"
        files_list = self.file_listup(path,".xls")
        if(len(files_list) <5): 
            self.imagrj_Macro()
             # 완료까지 대기
            self.macro_wait(self.folder_list)
       

        # data 별로 이미지 분석값확인해서 xlsx 만들기
        for i in self.folder_list:
            data_path = self.abspath +"data\\"+ i+"\\"
            temp_path = self.abspath +"temp\\"+ i+"\\"
            result_path = self.abspath +"result\\"+ i+"\\"

            file_list_bmp = self.file_listup(data_path,".bmp")
            data  = self.data_processor(file_list_bmp,i)
            files_list = self.file_listup(temp_path,".png")
            if(len(files_list) <5): 
                self.image_macro(file_list_bmp,data_path,temp_path)
            self.mkxl(data,data_path,result_path,temp_path,i)

        self.deleteAllFiles(self.temp_path)
        
     
    # bmp 이미지만 가져오기

    def file_listup(self,folder_path,end='*'):
        file_list = [file for file in os.listdir(
            folder_path) if file.endswith(end)]
        return file_list
    #폴더 가저오기
    def folder_listup(self):
        file_list = [file for file in os.listdir(
            self.data_path)]
        return file_list

    # 폴더 만들기

    def createFolder(self, folder_path):
        try:
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
        except OSError:
            print(OSError)

    # 폴더안에 파일 다지우기
    def deleteAllFiles(self,file_path):
        if(os.path.exists(file_path)):
            for file in os.scandir(file_path):
                shutil.rmtree(file.path)
    


    # 파일이름 _ 로 분리하기
    def filename_split(self, filename):
        _,  wl, lv = filename.split("_")
        wl = self.ledname_add(wl)
        lv = lv.strip('.bmp')
        full_name = filename.strip('.bmp')
        return wl, lv, full_name
        
    # LED 이름 넣기
    def ledname_add(self,name):
        return {'385nm': 'M385D1\n(385nm)',
        '470nm': 'XPEBBL-L1\n(470nm)',
        '565nm': 'M565D2\n(565nm)',
        '625nm': 'M625D3\n(625nm)',
        '660nm': 'M660D2\n(660nm)'}[name]

    # imageJ 매크로
    def imagrj_Macro(self):
        app = Application(backend='uia').connect(
            title="ImageJ", timeout=20)
        dig = app.ImageJ
        dig.set_focus()
        a=dig.child_window(title="응용 프로그램", auto_id="MenuBar", control_type="MenuBar").wrapper_object()
        a.click_input()
        time.sleep(0.4)
        pg.hotkey('down')
        pg.hotkey('right')
        pg.hotkey('down')
        pg.hotkey('enter')
        time.sleep(0.4)
        pyperclip.copy(
                self.abspath+"Macro.ijm")
        pg.hotkey('ctrl', 'v')
        pg.hotkey('enter')

    # imagej 매크로 부분
    def image_macro(self,file_list,data_path,temp_path):
        app = Application(backend='uia').connect(
            title="ImageJ", timeout=20)
        dig = app.ImageJ
        for img in file_list:
            _, _, full_name = self.filename_split(img)

            dig.set_focus()
            pg.hotkey('ctrl', 'o')
            time.sleep(0.4)

            pyperclip.copy(
                data_path+img)
            pg.hotkey('ctrl', 'v')
            pg.hotkey('enter')
            time.sleep(0.4)
            pg.hotkey('ctrl', 'a')
            pg.hotkey('ctrl', 'k')

            hwin = app.top_window()
            hwin.set_focus()
            window_title = hwin.window_text()

            img = app[window_title].capture_as_image()
            img.save(temp_path+full_name+'.png')
            dig.set_focus()
            time.sleep(0.4)
            pg.hotkey('ctrl', 'w')
            pg.hotkey('ctrl', 'w')

    def macro_wait(self,folder_list):
        t=0
        while t==0:
            folder=folder_list[-1]
            path = self.abspath+"temp\\"+folder+"\\"
            files_list = self.file_listup(path,".xls")
            if(len(files_list) >=5): 
                t=1
            time.sleep(1)
        

    # 데이터 처리
    def data_processor(self,file_path,folder_name):
        totdf = pd.DataFrame(
            columns=["Name", "wavelength", "Lv", "Average", "Min", "Max", "Uniformity"])
        for i in range(len(file_path)):
            df = pd.read_csv(self.temp_path +folder_name+"\\"+
                             file_path[i]+".xls", delimiter='\t')

            wl, lv, _ = self.filename_split(file_path[i])

            df_mean = round(df["Mean"][0],2)
            df_min = round(df["Min"][1],2)
            df_max = round(df["Max"][1],2)
            df_uniformity = round((1-((df["Max"][1]-df["Min"][1]) /
                               (df["Min"][1]+df["Max"][1])))*100,1)
            totdf.loc[i] = [file_path[i], wl, lv,
                            df_mean, df_min, df_max, df_uniformity]
        print(totdf)
        return totdf

    # 엑셀 데이터 넣기
    def mkxl(self,data,data_path,result_path,temp_path,folder_name):
        wb = openpyxl.load_workbook(self.abspath+'sample.xlsx')
        sheet = wb['Test Image2']
        
        for i in range(len(data)):
            temp = i*11
            #LED nm
            sheet[f'D{6+temp}'].value = data['wavelength'][i]
            #Lv
            sheet[f'G{6+temp}'].value = data['Lv'][i]
            #data image
            img = Image(data_path+data['Name'][i])
            img.height = 200
            img.width = 300
            sheet.add_image(img, f'G{7+temp}') 
            #intensity image
            img2 = Image(temp_path+data['Name'][i].replace('.bmp','.png'))
            img2.height = 200
            img2.width = 300
            sheet.add_image(img2, f'R{7+temp}') 
            #Uniformity Min
            sheet[f'AC{9+temp}'].value = data['Min'][i]
            #Uniformity Max
            sheet[f'AF{9+temp}'].value = data['Max'][i]
            #Uniformity Average
            sheet[f'AC{14+temp}'].value = data['Average'][i]
            
        
        

        wb.save(result_path+f"result{folder_name}.xlsx")
        
        
    
if __name__ == "__main__":

    

    data = Data()
    data.run()
