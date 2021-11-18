import datetime,sys,os.path,openpyxl,os,keyboard,time,thu_vien
from termcolor import colored, cprint
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
# 1. khai bao va khoi chay  web driver
options = webdriver.EdgeOptions()
options.page_load_strategy = 'normal'
# options.add_argument("--headless")
service = EdgeService(executable_path='.\msedgedriver.exe')
# 2. thuc hien dang nhap  va dieu huong vao trang can lay du lieu
Usr = "kv2zoom@gmail.com"
pas = "27bnguyenthanhhan"
word = ''
driver = webdriver.Edge(service=service, options=options)
driver.get("https://eticket.danang.gov.vn/#/login")
driver.find_element(By.ID, 'username').send_keys(Usr)
driver.find_element(By.ID, 'password').send_keys(pas)
driver.find_element(By.XPATH, "/html/body/div/div/div/form/div[2]/button").click()
driver.implicitly_wait(1)
driver.find_element(By.PARTIAL_LINK_TEXT, 'check-ins').click()
# 3. lay ten cot ( header th) va luu vao mang
header = []
th_s = driver.find_elements(By.TAG_NAME, 'th')
for cot in th_s:
    header.append(cot.text)
header.append('ĐĂNG KÍ?')
def cls(): return os.system('cls')
dem_clear = 0 # biến đếm dùng để lưu trữ số lần xuất log để mà clear terminals
word = input('Moi Ban Nhap C de thuc thi: ')
while(word.lower() != 'q'):
    if dem_clear == 10:
        cls()
        dem_clear = 0
    # 4. check va mo file excel luu tru
    file_name = datetime.now().strftime('%d_%m_%Y')+'.xlsx'
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
            dem_clear+=1
            continue
    else:
        # print('chau ton tai')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "NHAT KY NGAY"
        ws.append(header[1:])
    # file_name = datetime.now().strftime('%d_%m_%Y')+'.xlsx'
    ngay_hien_tai = datetime.now().strftime('%d-%m-%Y')

    if thu_vien.check_open_filelog('./data_usr.xlsx') :
        wb2 = openpyxl.load_workbook('./data_usr.xlsx')
        sheet = wb2['data']
        max_row = sheet.max_row
    else:
        dem_clear+=1
        continue
    if word.lower() == 'c':
        last_time_checkin = ws.cell(row=ws.max_row, column=12).value # ;ấy thời điểm cúi cùng của nhân sự checkin trong file log
        driver.refresh()
        time.sleep(1)
        row_data = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/main/div[2]/div/div[2]/div/table/tbody/tr[1]')
        td_s = row_data.find_elements(By.TAG_NAME, 'td')
        cmnd_log_in = td_s[5].text # lấy cmnd cua nhân sự đang checkin 
        HoTen = td_s[2].text
        DK = 0
        if td_s[12].text != last_time_checkin:# check xem thời gian nhân sự cúi cùng checkin trên eticket có trùng với nhân sự cúi cùng trong file log kh? nếu trùng thì bỏ qua kh phải xuất lại 
            print('ĐANG THỰC HIỆN QUÉT DỮ LIỆU TRÊN HỆ THỐNG ETICKET.DANANG.GOV.VN')
            print('HỌ VÀ TÊN NHÂN SỰ CHECKIN : ')
            cprint(HoTen, 'cyan',attrs=['bold'], file=sys.stderr)
            print('######################################')
            print('CMND/CCCD ĐANG THUC HIEN CHECK IN TAI CONG:')
            cprint(cmnd_log_in, 'cyan',attrs=['bold'], file=sys.stderr)
            # mở file dữ liệu và tra cứu cmnd nhân sự đang checkin + tra cứu thời hạn của tờ trình (ttr) và xét nghiệm
            
            for i in range(2, max_row):
                DK = 0
                ds_CMND = str(sheet.cell(row=i, column=9).value).strip().lower()
                if cmnd_log_in in ds_CMND:# nếu tìm thấy số cmnd trong danh sách thì lấy dữ liệu trong row đó
                    ttr_expire = str(sheet.cell(row=i, column=21).value).strip() # lấy thời gian hết hạn ttr
                    print('ttr:'+ttr_expire)
                    xet_nghiem_expire = str(sheet.cell(row=i, column=22).value).strip()# láy thời gian hết hạn xét nghiệm
                    print('xet nghiem:'+xet_nghiem_expire)
                    cprint("ĐÃ ĐĂNG KÍ TRONG DANH SÁCH", 'blue',attrs=['bold'], file=sys.stderr)
                    # kiểm tra coi 2 giá trị ngày hết hạn có hay kh ? nếu có thì kiểm tra đã hết hạn chưa?
                    if ttr_expire!='None'  and xet_nghiem_expire!='None':
                        if thu_vien.compare_ngay(ttr_expire,ngay_hien_tai):
                            cprint("TỜ TRÌNH CÒN HẠN", 'blue',attrs=['bold'], file=sys.stderr)
                        else:
                            cprint("TỜ TRÌNH ĐÃ HẾT HẠN", 'red',attrs=['bold'], file=sys.stderr)
                        if thu_vien.compare_ngay(xet_nghiem_expire,ngay_hien_tai):
                            cprint("TỜ XÉT NGHIỆM CÒN HẠN", 'blue',attrs=['bold'], file=sys.stderr)
                        else:
                            cprint("TỜ XÉT NGHIỆM ĐÃ HẾT HẠN", 'red',attrs=['bold'], file=sys.stderr)
                    else:
                        cprint("KHÔNG CÓ TỜ TRÌNH HOẶC GIẤY XÉT NGHIỆM", 'red',attrs=['bold'], file=sys.stderr)
                    DK = 1
                    register = 'ĐÃ ĐĂNG KÍ'
                    break
            if DK==0:
                cprint("CHƯA ĐĂNG KÍ TRONG DANH SÁCH", 'red',attrs=['bold'], file=sys.stderr)
                register = 'CHƯA ĐĂNG KÍ'
            dem_clear = dem_clear+1
            #GHI VÀO FILE LOG kết quả checkin mới nhất
            data_to_write = []#mảng chứa dữ liệu để ghi
            for cot in td_s:
                data_to_write.append(cot.text)
            data_to_write.append(register)
            ws.append(data_to_write[1:])
            try:
                wb.save(PATH)
            except PermissionError:
                print('file log đang mở , vui lòng đóng file log để thực hiện tiếp')
                continue
    time.sleep(1)
    if keyboard.is_pressed("q"):
        print("You pressed q")
        driver.close()
        break
  