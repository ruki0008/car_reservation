import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import gspread
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from pathlib import Path
import locale
import smtplib, sys
from email.mime.text import MIMEText


mail_from = ''
mail_to = ''
mail_subject = '予約結果'
messages = []

key_name = Path.cwd() / Path(__file__).parents[0] / 'car-scraping.json'
print(key_name)
sheet_id = ''
login_url = 'https://api.timesclub.jp/view/pc/tpLogin.jsp?siteKbn=TP&doa=ON&redirectPath=https%3A%2F%2Fshare.timescar.jp%2Fview%2Fmember%2Fmypage.jsp'
user_id1 = ''
user_id2 = ''
user_pass = ''

def main():
    gc = gspread.service_account(key_name).open_by_key(sheet_id)
    ws = gc.worksheets()[0]
    last_row = len(ws.col_values(1)) + 1
    print(last_row)
    ws_title = ws.title
    print(ws_title)
    for m in range(2, last_row):
        select_class = ws.cell(m, 1).value
        select_name = ws.cell(m, 2).value
        select_map = ws.cell(m, 3).value
        select_date = ws.cell(m, 4).value
        reserve_date_str = date_format(select_date)
        select_date = str(select_date).replace('/', '-')
        start_time = ws.cell(m, 5).value
        start_hour = str(start_time).split(':')[0]
        start_min = str(start_time).split(':')[1]
        end_time = ws.cell(m, 6).value
        end_hour = str(end_time).split(':')[0]
        end_min = str(end_time).split(':')[1]
        result = ws.cell(m, 7).value
        if result == '予約済':
            continue
        else:
            ws.update_cell(m, 7, 'その他')
        print(select_class, select_name, select_map, select_date, start_time, end_time)

        place = ''
        send_text = ''

        options = Options()
        options.add_argument('--headless')
        # 通常
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        # ヘッドレスモード
        # driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
        driver.maximize_window()
        driver.get(login_url)
        time.sleep(1)

        id_1 = driver.find_element(By.ID, 'cardNo1')
        id_1.send_keys(user_id1)
        id_2 = driver.find_element(By.ID, 'cardNo2')
        id_2.send_keys(user_id2)
        password = driver.find_element(By.ID, 'tpPassword')
        password.send_keys(user_pass)

        login_button = driver.find_element(By.ID,'doLoginForTp')
        login_button.click()
        time.sleep(5)

        driver.get('https://share.timescar.jp/car/')
        time.sleep(5)

        middle = driver.find_element(By.ID, 'cont02')
        premium = driver.find_element(By.ID, 'cont03')
        car_dl = driver.find_elements(By.CSS_SELECTOR, 'dl')
        car_list = []
        for dl in car_dl:
            try:
                car_name = dl.find_element(By.CSS_SELECTOR, 'img').get_attribute('alt')
            except:
                break
            if car_name != '':
                car_list.append(car_name)

        gc.values_clear(f"{ws_title}!H1:H100")
        car_len = len(car_list)
        cell_list = ws.range(f'H1:H{car_len}')
        for i, cell in enumerate(cell_list):
            cell.value = car_list[i]
        ws.update_cells(cell_list)

        if select_class == 'ミドルクラス':
            middle.click()
        elif select_class == 'プレミアムクラス':
            premium.click()
        time.sleep(3)

        for dl in car_dl:
            if dl.find_element(By.CSS_SELECTOR, 'img').get_attribute('alt') == select_name:
                dl.find_element(By.CLASS_NAME, 'station-btn').find_element(By.CLASS_NAME,'alpha').click()
                break
        time.sleep(3)
        search = driver.find_element(By.ID, 'narrowWord')
        search.send_keys(select_map)
        map_button = driver.find_element(By.ID, 'doNarrowSearch')
        map_button.click()
        time.sleep(3)

        while True:
            reserves_len = len(driver.find_elements(By.ID, 'isEnableToReserve'))
            for i in range(reserves_len):
                next_flg = False
                reserve_flg = False
                not_reserve_flg = False
                if not driver.find_elements(By.ID, 'isDispNext'):
                    next_flg = True

                reserves = driver.find_elements(By.ID, 'isEnableToReserve')
                reserves[i].click()
                time.sleep(3)

                place = driver.find_element(By.ID, 'stationNm').text

                reserve_date(driver, select_date, start_hour, 0)


                for l in range(2):
                    print(l)

                    div_car = driver.find_element(By.CLASS_NAME, 'tableon')
                    car_text = div_car.find_element(By.CSS_SELECTOR, 'p').text
                    print(car_text)
                    reserve_color1 = div_car.find_elements(By.CLASS_NAME, 'timelinedot')
                    reserve_color2 = div_car.find_elements(By.CLASS_NAME, 'timelinespace')

                    color_check = []

                    rental_time = time_count(start_time, end_time)
                    check_count = int(rental_time * 4)
                    print(check_count)


                    for j, color1_td in enumerate(reserve_color1):
                        k = j + 1
                        color_class = color1_td.get_attribute('class')
                        color = color_class.split(' ')[1]
                        color_check.append(color)
                        if k % 3 == 0:
                            k = j // 3
                            color_class = reserve_color2[k - 1].get_attribute('class')
                            color = color_class.split(' ')[1]
                            color_check.append(color)
                    print(color_check)

                    if l == 1:
                        check_count = check_count - 47
                        print(check_count)
                    for j in range(check_count):
                        if color_check[j] == 'vacant':
                            print('ok')
                        else:
                            print(f'{place}では予約ができませんでした。')
                            ws.update_cell(m, 7, '予約失敗')
                            not_reserve_flg = True
                            break
                        if check_count - 1 == j:
                            print(f'{place}で予約に空きがあります。')
                            ws.update_cell(m, 7, '予約済')

                            input_reserve_pack = driver.find_element(By.ID, 'pack')
                            select_reserve_pack = Select(input_reserve_pack)
                            select_reserve_pack.select_by_value('1')

                            input(driver, 'dateStart', f'{select_date} 00:00:00.0')
                            input(driver, 'hourStart', start_hour)
                            input(driver, 'minuteStart', start_min)
                            input(driver, 'dateEnd', f'{select_date} 00:00:00.0')
                            input(driver, 'hourEnd', end_hour)
                            input(driver, 'minuteEnd', end_min)

                            driver.find_element(By.ID, 'doCheck').click()
                            time.sleep(3)

                            driver.find_element(By.ID, 'doOnceRegist').click()

                            reserve_flg = True
                            break
                        if j == 47:
                            reserve_date(driver, select_date, start_hour, 12)
                            time.sleep(3)
                            print(j)
                            break
                        print(j)

                    if not_reserve_flg == True or reserve_flg == True:
                        break
                if reserve_flg == True:
                    break

                driver.get('https://share.timescar.jp/view/station/list.jsp')
                time.sleep(3)

            if reserve_flg == True:
                break

            if next_flg == True:
                break
            driver.find_element(By.ID, 'isDispNext').click()

        if reserve_flg == True:
            send_text = f'{place}で予約が完了しました。'
        elif reserve_flg == False:
            send_text = '予約が完了しませんでした。'
        messages.append(send_text)
    print(messages)
    if messages:
        messages_post = '\n'.join(messages)
        check_mail = send_mail(mail_from, mail_to, mail_subject, messages_post)
        print(check_mail)

    quit()

def time_count(start, end):
    start_hour = int(start.split(':')[0])
    start_min = int(start.split(':')[1])
    start_min = start_min / 60
    print(start_min)

    end_hour = int(end.split(':')[0])
    end_min = int(end.split(':')[1])
    end_min = end_min / 60
    print(end_min)

    time = end_hour + end_min - start_hour - start_min
    return time

def reserve_date(driver, select_date, start_hour, plus_num):
    reserve_date = driver.find_element(By.ID, 'dateSpace')
    date_select = Select(reserve_date)
    date_select.select_by_value(f'{select_date} 00:00:00.0')

    reserve_time = driver.find_element(By.ID, 'hourSpace')
    select_time = Select(reserve_time)
    select_time.select_by_value(str(int(start_hour) + int(plus_num)))

    date_button = driver.find_element(By.ID, 'doSearchTargetTimetable')
    date_button.click()

def date_format(date):
    locale.setlocale(locale.LC_TIME, 'ja_JP.UTF-8')
    s = date
    s_format = '%Y/%m/%d'
    dt = datetime.datetime.strptime(s, s_format)
    dt = dt.strftime('%Y年%m月%d日（%a）')
    print(dt)
    return dt

def input(driver, css, value):
    input_area = driver.find_element(By.ID, css)
    select_area = Select(input_area)
    select_area.select_by_value(value)

def send_mail(mail_from, mail_to, mail_subject, mail_body):

    msg = MIMEText(mail_body, "plain", "utf-8")
    msg['Subject'] = mail_subject
    msg['From'] = mail_from
    msg['To'] = mail_to


    try:
        smtpobj = smtplib.SMTP_SSL('sv14082.xserver.jp', 465)
        smtpobj.ehlo()
        id = ''
        passwd = ''
        smtpobj.login(id, passwd)

        smtpobj.sendmail(mail_from, mail_to, msg.as_string())

        smtpobj.quit()

    except Exception as e:
        print(e)

    return "メール送信完了"

if __name__ == '__main__':
    main()