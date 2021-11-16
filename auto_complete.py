import datetime,sys,os.path,openpyxl,os,keyboard,time
from termcolor import colored, cprint
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
# 1. khai bao va khoi chay  web driver
options = webdriver.EdgeOptions()
options.page_load_strategy = 'normal'
options.add_argument("--headless")
service = EdgeService(executable_path='.\msedgedriver.exe')
# 2. thuc hien dang nhap  va dieu huong vao trang can lay du lieu
Usr = "kv2zoom@gmail.com"
pas = "27bnguyenthanhhan"
word = ''
driver = webdriver.Edge(service=service, options=options)
driver.get("https://eticket.danang.gov.vn/#/login")
driver.find_element(By.ID, 'username').send_keys(Usr)
driver.find_element(By.ID, 'password').send_keys(pas)
driver.find_element(
    By.XPATH, "/html/body/div/div/div/form/div[2]/button").click()
driver.implicitly_wait(1)
driver.find_element(By.PARTIAL_LINK_TEXT, 'check-ins').click()
# 3. lay ten cot ( header th) va luu vao mang
header = []
th_s = driver.find_elements(By.TAG_NAME, 'th')
for cot in th_s:
    header.append(cot.text)
header.append('ĐĂNG KÍ?')
# print(len(header))
# 4. check va mo file excel luu tru
# file_name = datetime.datetime.now().strftime('%d_%m_%Y')+'.xlsx'
# print(file_name)
# PATH = './log/'+file_name
# if os.path.isfile(PATH):
#     # print('da ton tai')
#     wb = openpyxl.load_workbook(filename=PATH)
#     ws = wb.active
# else:
#     # print('chau ton tai')
#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "NHAT KY NGAY"
#     ws.append(header[1:])
    # cells=ws['A1':'O1']
    #  cells.font = openpyxl.styles.Font(bold="true",size=16)
    # wb.save(PATH)
##lay thoi diem gan nhat checkin 

# print(last_time_checkin)
## 5. doc du lieu va dua vao excel
cls = lambda: os.system('cls')
dem_clear=0
word = input('Moi Ban Nhap C de thuc thi: ')
while(word.lower() != 'q'):
    
    if dem_clear==10:
        cls()
        dem_clear=0
    # 4. check va mo file excel luu tru
    file_name = datetime.now().strftime('%d_%m_%Y')+'.xlsx'
    ngay_hien_tai=int(datetime.now().strftime('%j'))
    # print(file_name)
    PATH = './log/'+file_name
    if os.path.isfile(PATH):
        # print('da ton tai')
        try:
            wb = openpyxl.load_workbook(filename=PATH)
            ws = wb.active
        except PermissionError :
            print('file log đang mở , vui lòng đóng file log để thực hiện tiếp')
            dem_clear=dem_clear+1
            time.sleep(3)
            continue
    else:
        # print('chau ton tai')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "NHAT KY NGAY"
        ws.append(header[1:])
    if word=='c':
        last_time_checkin = ws.cell(row=ws.max_row , column=12).value
        
        driver.refresh()
        time.sleep(1)
        row_data=driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div/main/div[2]/div/div[2]/div/table/tbody/tr[1]')
        td_s=row_data.find_elements(By.TAG_NAME,'td')
        cmnd_log_in=td_s[5].text
        HoTen=td_s[2].text
        DK=0
        if td_s[12].text != last_time_checkin:
            print('DANG THUC HIEN QUET DU LIEU TREN HE THONG ETICKET.DANANG.GOV.VN')
            print('HO va Ten nhan su dang checkin : '+HoTen)
            print('CMND/CCCD ĐANG THUC HIEN CHECK IN TAI CONG:'+cmnd_log_in)
            # driver.close()
            #mo file excel csdl nguoi dung de tim kiem
            # cur_dic=os.getcwd()
            wb2 = openpyxl.load_workbook('./data_usr.xlsx')
            sheet = wb2['data']
            
            max_row=sheet.max_row
            for i in range(2,max_row):
                DK=0
                ds_CMND=str(sheet.cell(row=i, column=9).value).strip().lower()
                if cmnd_log_in in ds_CMND:
                    ttr_string=str(sheet.cell(row=i, column=21).value).strip()
                    day_ttr_string=str(sheet.cell(row=i, column=22).value).strip()
                    xet_nghiem_string=str(sheet.cell(row=i, column=23).value).strip()
                    day_xet_nghiem_string=str(sheet.cell(row=i, column=24).value).strip()
                    DK=1
                    break
        
            if DK==1 :
                # text = colored('DA DANG KI TRONG DANH SACH', 'blue', attrs=['reverse', 'blink'])
                cprint("DA DANG KI TRONG DANH SACH", 'blue', attrs=['bold'], file=sys.stderr)
                register='ĐÃ ĐĂNG KÍ'
                lech_ngay_ttr = ngay_hien_tai-int((datetime.strptime(ttr_string, "%d-%m-%Y")).strftime("%j"))
                lech_ngay_vac_xin = ngay_hien_tai-int((datetime.strptime(xet_nghiem_string, "%d-%m-%Y")).strftime("%j"))
                if (lech_ngay_ttr>int(day_ttr_string)):
                    # print('ĐÃ QUÁ HẠN TỜ TRÌNH')
                    cprint("ĐÃ QUÁ HẠN TỜ TRÌNH", 'red', attrs=['bold'], file=sys.stderr)
                else :
                    # print('CHƯA QUÁ HẠN TỜ TRÌNH')
                    cprint("CHƯA QUÁ HẠN TỜ TRÌNH", 'green', attrs=['bold'], file=sys.stderr)
                if (lech_ngay_vac_xin>int(day_xet_nghiem_string)):
                    # print('ĐÃ QUÁ HẠN VACCIN')
                    cprint("ĐÃ QUÁ HẠN XÉT NGHIỆM", 'RED', attrs=['bold'], file=sys.stderr)
                    
                else :
                    # print('CHƯA QUÁ HẠN VACCIN')
                    cprint("CHƯA QUÁ HẠN XÉT NGHIỆM", 'green', attrs=['bold'], file=sys.stderr)
                # print(lech_ngay)

            else:
                cprint("CHUA DANG KI TRONG DANH SACH", 'red', attrs=['bold'], file=sys.stderr)
                register='CHƯA ĐĂNG KÍ'
            dem_clear=dem_clear+1
        #sau khi checkin thi ghi du lieu vao log excel
        # last_time_checkin = ws.cell(row=ws.max_row , column=12).value
        # if td_s[12].text != last_time_checkin:
            data_to_write=[]
            for cot in td_s:
                data_to_write.append(cot.text)
            data_to_write.append(register)
            ws.append(data_to_write[1:])
            # try:
            wb.save(PATH)
            # except Exception as e:
                # print(e.message)
            wb.close()
            wb2.close()
        # print(data_to_write)c
        time.sleep(1)
        if keyboard.is_pressed("q"):
            print("You pressed q")
            break
    # word=input('Nhan Q de thoat , nhan C de tiep tuc :')
    # print('Nhan Q de thoat , nhan phim bat ki  de tiep tuc :')
    # if keyboard.read_key() == "q":
    #     driver.quit()
    #     wb.save(PATH)
    #     break
