import openpyxl
from moviepy.editor import VideoFileClip, concatenate_videoclips
import urllib
import requests
from bs4 import BeautifulSoup
import cv2
import pytesseract
import numpy as np
import pyscreenshot
import pyautogui
from selenium import webdriver
from PIL import Image
from io import BytesIO
from selenium.webdriver.common.by import By
import shutil
import os
import re
import time
import pandas as pd
from googleapiclient.discovery import build
import google_auth_oauthlib.flow
import googleapiclient.errors
import googleapiclient.discovery
from googleapiclient.http import MediaFileUpload
from threading import Thread
import datetime
import pygetwindow as gw
import random

start_time = time.time()

string_url = 1

df = pd.read_excel('C:/Users/Home/Desktop/xl.xlsx', sheet_name=2, header=None)


dimensions = df.shape
finish_video = 0
print("Всего видосов:", dimensions[0])
while string_url <= dimensions[0]:
    book = openpyxl.open("C:/Users/Home/Desktop/xl.xlsx", read_only=True) #открытие файла для чтения
    sheet = book.active #открытие активного листа

    url = sheet[string_url][0].value #взятие адреса из заданой ячейки
    time_upload = f'{str(df[1][string_url-1])[:4]}-{str(df[1][string_url-1])[4:6]}' \
                  f'-{str(df[1][string_url-1])[6:8]}T19:00:00+07:00'

    f = urllib.request.urlopen(url).read()
    soup = BeautifulSoup(f, 'html.parser')
    for link in soup.find_all('a', class_="btn_l-orange"):
        url_download = "http://wotreplays.ru" + link.get('href')
    # читаем сайт парсером, находим кнопку скачать

    r = requests.get(url_download, stream=True)
    with open('1.wotreplay', 'bw') as f:
        for chunk in r.iter_content(8192):
            f.write(chunk)
    # скачиваем файл в директорию скрипта
    print('Обрабатывается видос №', string_url)
    string_url += 1
    # снимаем скрины результатов игры

    confidence = 0.8

    fox = webdriver.Firefox()
    fox.set_window_size(1200, 1200)
    fox.get(url)
    time.sleep(5)

    # берем имя танка, количество фрагов и урон
    r = requests.get(url)
    s = BeautifulSoup(r.text, 'html.parser')
    u = s.find("div", class_="replay-stats__tank_title")
    tank_name = ' '.join(re.findall(r'>([^<>]+)<', str(u)))
    tank_name = tank_name.replace('/', ' ')  # без замены выдает ошибку

    u = s.find_all("span", class_="replay-stats__summary_effective")
    u = ' '.join(re.findall(r'>([^<>]+)<', str(u)))
    u = u.split()
    DMG = u[2]
    DMG = int(DMG)/1000
    DMG = round(DMG, 1)
    DMG = str(DMG)
    DMG = DMG.replace('.', ',')
    frags = u[0]
    if frags == '-':
        frags = '0'

    u = s.find_all("ul", class_="replay-details__table")
    u = ' '.join(re.findall(r'>([^<>]+)<', str(u[-1])))
    x = u.index('м')
    minute = u[x-3:x-1]
    sec = u[x+4:x+6]
    minute = int(minute)
    sec = int(sec)
    time_fight = minute*60 + sec
    # print(time_fight)
    time.sleep(10)
    try:
        path0 = 'buttons/0.png'
        button = pyautogui.locateOnScreen(path0, confidence=confidence)
        pyautogui.click(button)

        pyautogui.keyDown('ctrl')
        time.sleep(1)
        pyautogui.vscroll(-1000)
        time.sleep(1)
        pyautogui.vscroll(-1000)
        time.sleep(1)
        pyautogui.vscroll(-1000)
        time.sleep(1)
        pyautogui.keyUp('ctrl')
        time.sleep(1)

        element = fox.find_element(By.CSS_SELECTOR, 'div.replay__tab:nth-child(2)')
        location = element.location
        size = element.size
        # print(location, size)
        png = fox.get_screenshot_as_png()

        im = Image.open(BytesIO(png))

        left = location['x']
        top = location['y']
        right = location['x'] + size['width']
        bottom = location['y'] + size['height']
        # print(left,top,right,bottom)
        im = im.crop((left, top, right, bottom)) # defines crop points

        im.save('1.png')

        path1 = 'buttons/1.png'

        button = pyautogui.locateOnScreen(path1, confidence=confidence)

        time.sleep(1)
        pyautogui.click(button)
        time.sleep(1)

        element = fox.find_element(By.CSS_SELECTOR, 'div.replay__tab:nth-child(3)')
        location = element.location
        size = element.size
        # print(location, size)
        png = fox.get_screenshot_as_png()

        im = Image.open(BytesIO(png))

        left = location['x']
        top = location['y']
        right = location['x'] + size['width']
        bottom = location['y'] + size['height']
        # print(left,top,right,bottom)
        im = im.crop((left, top, right, bottom))
        im.save('2.png')

        path2 = 'buttons/2.png'

        button = pyautogui.locateOnScreen(path2, confidence=confidence)

        time.sleep(1)
        pyautogui.click(button)
        time.sleep(1)

        pyautogui.keyDown('ctrl')
        time.sleep(1)
        pyautogui.vscroll(1000)
        time.sleep(1)
        pyautogui.vscroll(1000)
        time.sleep(1)
        pyautogui.vscroll(1000)
        time.sleep(1)
        pyautogui.keyUp('ctrl')
        time.sleep(1)
        otstup = -340
        pyautogui.vscroll(-200)
        time.sleep(1)
        pyautogui.vscroll(-200)
        time.sleep(1)
        pyautogui.vscroll(-200)
        time.sleep(1)
        pyautogui.vscroll(-200)
        time.sleep(1)

        time.sleep(3)

        element = fox.find_element(By.CSS_SELECTOR, 'div.replay__tab:nth-child(4)') # найдите часть страницы, изображение
        location = element.location
        size = element.size
        # print(location, size)
        png = fox.get_screenshot_as_png()

        im = Image.open(BytesIO(png))

        left = location['x']
        top = location['y'] + otstup
        right = location['x'] + size['width']
        bottom = location['y'] + size['height'] + otstup
        # print(left,top,right,bottom)
        im = im.crop((left, top, right, bottom))
        im.save('3.png') # сохраняет новое обрезанное изображение

        fox.quit()
    except:
        fox.quit()
        print("Видос №", string_url - 1, "не получился")
        continue

    # свернуть все окна
    # os. или pyautogui.

    os.startfile('C:/Users/Home/PycharmProjects/pythonProject1/1.wotreplay')
    # запускаем файл

    pytesseract.pytesseract.tesseract_cmd = 'C:\Program Files\Tesseract-OCR/tesseract.exe'

    time.sleep(10)
    # НАЧАЛО ЗАПИСИ

    filename = '1time.jpg'

    while True:
        screen =  np.array(pyscreenshot.grab(bbox=(1857, 22, 1906, 45)))
        cv2.imwrite(filename, screen)
        img = cv2.imread('1time.jpg')
        config = r'--oem 3 --psm 12'
        text = pytesseract.image_to_string(img, config=config, lang='rus')
        # print(text)
        cv2.destroyAllWindows()
        if '00:10' in text:
            pyautogui.hotkey('shift', 'f9')
            break

    #КОНЕЦ ЗАПИСИ

    # while(True):
    #     screen = np.array(pyscreenshot.grab(bbox=(820, 110, 1095, 185)))
    #     cv2.imwrite(filename, screen)
    #     img = cv2.imread('1time.jpg')
    #     config = r'--oem 3 --psm 12'
    #     text = pytesseract.image_to_string(img, config=config, lang='rus')
    #     print(text)
    #     cv2.destroyAllWindows()
    #     if 'РАЖЕ' in text or 'ОБЕД' in text or 'НИЧЬ' in text:
    #         pyautogui.hotkey('shift', 'f9')
    #         break

    time.sleep(time_fight+10)
    pyautogui.hotkey('shift', 'f9')
    time.sleep(6)
    pyautogui.hotkey('alt', 'f4')
    # закрываем файл
    # os.closefile('C:/Users/Home/PycharmProjects/pythonProject1/1.wotreplay')

    # ищем видеофайл в папке
    time.sleep(5)
    fails = os.listdir('E:/bandicam/World of Tanks/')
    lst_video = []
    for i in fails:
        if i.endswith('mp4'):
            lst_video.append(i)
            R = lst_video[-1]

    # перекидываем скрины из папки со скриптом в папку захвата видео
    shutil.move('1.png', 'E:/bandicam/World of Tanks/1.png')
    shutil.move('2.png', 'E:/bandicam/World of Tanks/2.png')
    shutil.move('3.png', 'E:/bandicam/World of Tanks/3.png')

    # делаем видос из трех скринов
    path = 'E:/bandicam/World of Tanks/'
    out_path = 'E:/bandicam/World of Tanks/'
    out_video_name = 'T'
    out_video_full_path = out_path+out_video_name

    pre_imgs = os.listdir(path)

    img = []

    for i in pre_imgs:
        if i.endswith('png'):
            i = path+i
            img.append(i)

    cv2_fourcc = cv2.VideoWriter_fourcc(*'mp4v')

    size_list = []

    for i in img:
        frame = cv2. imread(i)
        size = list(frame.shape)
        del size[2]
        size.reverse()
        size_list.append(size)

    for i in range(len(img)):
        video = cv2.VideoWriter(out_video_full_path+str(i+1)+'.mp4', cv2_fourcc, 0.5, size_list[i])
        video.write(cv2.imread(img[i]))
        video.release()

    # даем название клипу
    if int(frags) == 0:
        clip_name = f"WoT {tank_name} - {DMG}K урона {frags} фрагов ({DMG}K DMG {frags} frag).mp4"
    elif int(frags) == 1:
        clip_name = f"WoT {tank_name} - {DMG}K урона {frags} фраг ({DMG}K DMG {frags} frag).mp4"
    elif int(frags) < 5:
        clip_name = f"WoT {tank_name} - {DMG}K урона {frags} фрага ({DMG}K DMG {frags} frags).mp4"
    else:
        clip_name = f"WoT {tank_name} - {DMG}K урона {frags} фрагов ({DMG}K DMG {frags} frags).mp4"

    # пишем список тегов
    tags = ['world of tanks', 'wot', 'wot replays', 'free to play', 'world of tanks gameplay', 'мир танков',
            'world of tanks', 'лучший бой', f'world of tanks {tank_name}', 'геймплей', 'f2p',
            f'world of tanks {tank_name} обзор', f'wot {tank_name}', 'world of tanks alex', 'world of tanks update',
            'world of tanks the best battle', f'world of tanks {tank_name} gameplay',
            f'world of tanks {tank_name}', f'world of tanks {tank_name} review', f'wot {tank_name}',
            f'wot {tank_name}', 'реплей', 'реплеи', 'алекс вот реплей', 'вот реплеи', 'replays', 'replay', 'wot replay']

    tags = ['world of tanks', 'wot', 'wot replays', 'free 2 play', 'world of tanks gameplay', 'мир танков',
            'world of tanks', 'лучший бой', f'world of tanks {tank_name}', 'геймплей', 'f2p',
            f'world of tanks {tank_name} обзор', f'wot {tank_name}', 'world of tanks alex', 'world of tanks update',
            'world of tanks the best battle', f'world of tanks {tank_name} gameplay',
            f'world of tanks {tank_name}', f'world of tanks {tank_name} review', f'wot {tank_name}',
            f'wot {tank_name}', 'реплей', 'алекс вот реплей']

    # склеиваем видос и скрины
    clip_finish = VideoFileClip("E:/bandicam/Desktop/Finish.mp4")
    clip_start = VideoFileClip("E:/bandicam/Desktop/Start.mp4")
    clip_4 = VideoFileClip("E:/bandicam/World of Tanks/T3.mp4")
    clip_3 = VideoFileClip("E:/bandicam/World of Tanks/T2.mp4")
    clip_2 = VideoFileClip("E:/bandicam/World of Tanks/T1.mp4")
    clip_1 = VideoFileClip("E:/bandicam/World of Tanks/" + R)
    # final_clip = concatenate_videoclips([clip_start, clip_1, clip_2, clip_3, clip_4, clip_finish], method='compose')
    final_clip = concatenate_videoclips([clip_1, clip_2, clip_3, clip_4], method='compose')
    final_clip.write_videofile(
                               "E:/bandicam/World of Tanks/" + clip_name,
                               fps=30,
                               verbose=False,
                               threads=8,
                               preset='ultrafast',
                               logger=None
                               )
    os.replace("E:/bandicam/World of Tanks/" + clip_name,
               'E:/bandicam/ready/' + clip_name)
    finish_video += 1
    file_source = 'E:/bandicam/World of Tanks/'
    get_files = os.listdir(file_source)
    try:
        for g in get_files:
            os.remove(file_source + g)
    except:
        continue
    # удалить все временные файлы

    # грузим ролик на канал
    CLIENT_SECRET_FILE = 'client.json'
    API_NAME = 'youtube'
    API_VERSION = 'v3'
    SCOPES = ['https://www.googleapis.com/auth/youtube.upload']

    # Авторизация и получение доступа к API
    flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(
        CLIENT_SECRET_FILE, SCOPES)


    def b():
        confidence = 0.6
        while True:

            try:
                path4 = f'buttons/4.png'
                button_location = pyautogui.locateOnScreen(path4, confidence=confidence)
                button_width = button_location.width
                button_height = button_location.height
                click_x = button_location.left + button_width // 2
                click_y = button_location.top + button_height // 2
                pyautogui.click(click_x, click_y)
                time.sleep(2)
            except:
                continue
            try:
                path5 = 'buttons/5.png'
                button_location = pyautogui.locateOnScreen(path5, confidence=confidence)
                button_width = button_location.width
                button_height = button_location.height
                click_x = button_location.left + button_width // 2
                click_y = button_location.top + button_height // 2
                pyautogui.click(click_x, click_y)
                # pyautogui.click(click_x, click_y, clicks=1, interval=1)
                # pyautogui.mouseDown(); time.sleep(0.1); pyautogui.mouseUp()
                time.sleep(2)
            except:
                continue
            try:
                path6 = 'buttons/6.png'
                button_location = pyautogui.locateOnScreen(path6, confidence=confidence)
                button_width = button_location.width
                button_height = button_location.height
                click_x = button_location.left + button_width // 2
                click_y = button_location.top + button_height // 2
                pyautogui.click(click_x, click_y)
                time.sleep(2)
            except:
                continue
            try:
                path7 = 'buttons/7.png'
                button_location = pyautogui.locateOnScreen(path7, confidence=confidence)
                button_width = button_location.width
                button_height = button_location.height
                click_x = button_location.left + button_width // 2
                click_y = button_location.top + button_height // 2
                pyautogui.click(click_x, click_y)
                time.sleep(2)

                # Свернуть все окна
                # Получить список всех открытых окон
                windows = gw.getAllWindows()
                # Свернуть каждое окно
                for window in windows:
                    window.minimize()

                # print("Кнопочки протыканы")
            except:
                # print('Пока ничего не видно')
                continue

    T = Thread(target=b, daemon=True)
    T.start()

    def u():
        start_time = time.time()
        thread_clip_name = clip_name
        thread_tank_name = tank_name
        thread_time_upload = time_upload
        tags = ['world of tanks', 'wot', 'wot replays', 'free 2 play', 'world of tanks gameplay', 'мир танков',
                'world of tanks', 'лучший бой', f'world of tanks {tank_name}', 'геймплей',
                f'world of tanks {tank_name} обзор', f'wot {tank_name}', 'world of tanks alex',
                'world of tanks update',
                'world of tanks the best battle', f'world of tanks {tank_name} gameplay',
                f'world of tanks {tank_name}', f'world of tanks {tank_name} review', f'wot {tank_name}']

        # Загрузка видео
        CLIENT_SECRET_FILE = 'client.json'
        API_NAME = 'youtube'
        API_VERSION = 'v3'
        SCOPES = ['https://www.googleapis.com/auth/youtube.upload']

        # Авторизация и получение доступа к API
        flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)

        creds = flow.run_local_server(port=0)

        youtube = googleapiclient.discovery.build(API_NAME, API_VERSION, credentials=creds)

        # Создание объекта загрузки видео
        request_body = {
            'snippet': {
                'title': f'{thread_clip_name[:-4]}',
                'description': 'Подпишитесь чтобы увидеть больше новых реплеев! Subscribe for more replays!',
                'tags': tags,
                'categoryId': '20'
            },
            'status': {
                # 'privacyStatus': 'public',
                'privacyStatus': 'private',
                'publishAt': f'{thread_time_upload}',  # укажите желаемое время публикации в формате ISO 8601
                'madeForKids': False,
                'selfDeclaredMadeForKids': False,
            },
            'notifySubscribers': False
        }
        media = googleapiclient.http.MediaFileUpload(f'E:/bandicam/ready/{thread_clip_name}',
                                                     chunksize=5 * 1024 * 1024,
                                                     resumable=True)
        try:
            response = youtube.videos().insert(
                part='snippet,status',
                body=request_body,
                media_body=media,
                notifySubscribers=request_body['notifySubscribers']
            ).execute()

            youtube.thumbnails().set(
                videoId=response.get('id'),
                media_body=MediaFileUpload(
                    f'E:/bandicam/previews/{thread_tank_name.upper()}/{random.randint(0, 7)}.jpg'
                )
            ).execute()

            print(
                f'Видео {thread_clip_name} успешно загружено на YouTube'
                  )
            work_time = time.time() - start_time

            print('Время работы потока:', str(datetime.timedelta(seconds=work_time)))
            size_in_bytes = os.path.getsize(f'E:/bandicam/ready/{thread_clip_name}')
            size_in_Mb = size_in_bytes / (1024 * 1024)
            print(f"UPLOAD SPEED THREAD {thread_tank_name}: {round(size_in_Mb / work_time, 2)} Mb/s")
            os.remove(f'E:/bandicam/ready/{thread_clip_name}')
        except googleapiclient.errors.HttpError as e:
            print(f'Возникла ошибка при загрузке видео: {e}')


    T = Thread(target=u)
    T.start()
    time.sleep(30)

work_time = time.time()-start_time
mean_video = work_time // (dimensions[0])
print(
        'Обработано видео:',
        finish_video)
print(
    'Потрачено времени:',
    'часов:', int(work_time // 3600),
    '\tминут:', int((work_time // 60) % 60),
    '\tсекунд:', int(work_time % 60)
    )
print(
    'Среднее время обработки видеоклипа:',
    'часов:', int(mean_video // 3600),
    'минут:', int((mean_video // 60) % 60),
    '\tсекунд:', int(mean_video % 60)
    )

file_source = 'E:/bandicam/World of Tanks/'
get_files = os.listdir(file_source)
try:
    for g in get_files:
        os.remove(file_source + g)
except:
    print('Не удалось очистить папку для временных файлов')


