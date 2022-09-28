import fps as fps
import openpyxl
from moviepy.editor import VideoFileClip, concatenate_videoclips
import urllib
import requests
from bs4 import BeautifulSoup
import cv2
import pytesseract
import numpy as np
import pyscreenshot as ImageGrab
import time
import pyautogui
from selenium import webdriver
from PIL import Image
from io import BytesIO
from selenium.webdriver.common.by import By
import shutil
import os
import re

book = openpyxl.open("C:/Users/Home/Desktop/xl.xlsx", read_only=True) #открытие файла для чтения
sheet = book.active #открытие активного листа
string_url = 1
url = sheet[string_url][2].value #взятие адреса из заданой ячейки
# time_v = # взятие времени видоса
# date_v = # взятие даты видоса

f = urllib.request.urlopen(url).read()
soup = BeautifulSoup (f,'html.parser')
for link in soup.find_all('a', class_="btn_l-orange"):
    url_download = "http://wotreplays.ru" + link.get('href')
# читаем сайт парсером, находим кнопку скачать

r = requests.get(url_download, stream=True)
with open('1.wotreplay', 'bw') as f:
    for chunk in r.iter_content(8192):
        f.write(chunk)
# скачиваем файл в директорию скрипта

# снимаем скрины результатов игры

confidence = 0.8

fox = webdriver.Firefox()
fox.get(url)
time.sleep(5)

# берем имя танка, количество фрагов и урон
r = requests.get(url)
s = BeautifulSoup(r.text, 'html.parser')
u = s.find("div", class_="replay-stats__tank_title")
tank_name = ' '.join(re.findall(r'>([^<>]+)<', str(u)))

u = s.find_all("span", class_="replay-stats__summary_effective")
u = ' '.join(re.findall(r'>([^<>]+)<', str(u)))
u = u.split()
DMG = u[2]
DMG = int(DMG)/1000
DMG = round(DMG, 1)
DMG = str(DMG)
DMG = DMG.replace('.',',')
frags = u[0]

u = s.find_all("ul", class_="replay-details__table")
u = ' '.join(re.findall(r'>([^<>]+)<', str(u[-1])))
x = u.index('м')
min = u[x-3:x-1]
sec = u[x+4:x+6]
min = int(min)
sec = int(sec)
time_fight = min*60+sec
# print(time_fight)
time.sleep(10)

path0 = 'buttons/0.png'
button = pyautogui.locateOnScreen(path0, confidence = confidence)
pyautogui.click(button)

pyautogui.keyDown('ctrl')
time.sleep(1)
pyautogui.vscroll(-1000)
pyautogui.vscroll(-1000)
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

button = pyautogui.locateOnScreen(path1, confidence = confidence)

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

button = pyautogui.locateOnScreen(path2, confidence = confidence)

time.sleep(1)
pyautogui.click(button)
time.sleep(1)

pyautogui.keyDown('ctrl')
time.sleep(1)
pyautogui.vscroll(1000)
pyautogui.vscroll(1000)
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


os.startfile('C:/Users/Home/PycharmProjects/pythonProject1/1.wotreplay')
# запускаем файл

pytesseract.pytesseract.tesseract_cmd = 'C:\Program Files\Tesseract-OCR/tesseract.exe'

time.sleep(10)
# НАЧАЛО ЗАПИСИ

filename = '1time.jpg'

while(True):
    screen =  np.array(ImageGrab.grab(bbox=(1857, 22, 1906, 45)))
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
#     screen = np.array(ImageGrab.grab(bbox=(820, 110, 1095, 185)))
#     cv2.imwrite(filename, screen)
#     img = cv2.imread('1time.jpg')
#     config = r'--oem 3 --psm 12'
#     text = pytesseract.image_to_string(img, config=config, lang='rus')
#     print(text)
#     cv2.destroyAllWindows()
#     if 'РАЖЕ' in text or 'ОБЕД' in text or 'НИЧЬ' in text:
#         pyautogui.hotkey('shift', 'f9')
#         break

time.sleep(time_fight+13)
pyautogui.hotkey('shift', 'f9')
time.sleep(5)
pyautogui.hotkey('alt', 'f4')

# перекидываем скрины из папки со скриптом в папку захвата видео
shutil.move('1.png', 'E:/bandicam/World_of_tanks_ru/1.png')
shutil.move('2.png', 'E:/bandicam/World_of_tanks_ru/2.png')
shutil.move('3.png', 'E:/bandicam/World_of_tanks_ru/3.png')

# ищем первый видеофайл в папке
fails = os.listdir('E:/bandicam/World_of_tanks_ru/')
for i in fails:
    if i.endswith('mp4'):
        R = i
        break

# делаем видос из трех скринов
path = 'E:/bandicam/World_of_tanks_ru/'
out_path = 'E:/bandicam/World_of_tanks_ru/'
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

# склеиваем видос и скрины

clip_4 = VideoFileClip("E:/bandicam/World_of_tanks_ru/T3.mp4")
clip_3 = VideoFileClip("E:/bandicam/World_of_tanks_ru/T2.mp4")
clip_2 = VideoFileClip("E:/bandicam/World_of_tanks_ru/T1.mp4")
clip_1 = VideoFileClip("E:/bandicam/World_of_tanks_ru/"+R)
final_clip = concatenate_videoclips([clip_1, clip_2, clip_3, clip_4], method='compose')
final_clip.write_videofile("E:/bandicam/World_of_tanks_ru/WoT "+tank_name+' - '+DMG+'K урона '+frags+' фрагов ('+DMG+'K DMG '+frags+' frags)'+".mp4")
os.replace("E:/bandicam/World_of_tanks_ru/WoT "+tank_name+' - '+DMG+'K урона '+frags+' фрагов ('+DMG+'K DMG '+frags+' frags)'+".mp4",
           'E:/bandicam/ready/WoT '+tank_name+' - '+DMG+'K урона '+frags+' фрагов ('+DMG+'K DMG '+frags+' frags)'+".mp4")

file_source = 'E:/bandicam/World_of_tanks_ru/'
get_files = os.listdir(file_source)
for g in get_files:
    os.remove(file_source + g)
# удалить все временные файлы



# string_url +=1
