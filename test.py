import psutil
import colorama
import os
os.environ['PYGAME_HIDE_SUPPORT_PROMPT'] = "hide"
import win32api
import win32file
import wmi
import pygame
import pyaudio
from pygame import mixer
import time
import sounddevice as sd
import subprocess
from termcolor import colored
import socket
import requests
import bluetooth
import speech_recognition as sr
from fuzzywuzzy import fuzz
import wave
import keyboard

os.system("cls")
colorama.init()
# Автор кода 
print("After: NoFranxx")
# Оперативная память
memory_usage = psutil.virtual_memory().total

ram_size = psutil.virtual_memory().total  # total RAM size in bytes

if ram_size > 12 * 10 ** 9: 
    print(colorama.Fore.GREEN + "Memory_GOOD")
else:
    print(colorama.Fore.RED + "Memory_Fail")
    
 #вывод дисков v.2
drive_types = {
                win32file.DRIVE_UNKNOWN : "Unknown\nDrive type can't be determined.",
                win32file.DRIVE_REMOVABLE : "Removable\nDrive has removable media. This includes all floppy drives and many other varieties of storage devices.",
                win32file.DRIVE_FIXED : "Fixed\nDrive has fixed (nonremovable) media. This includes all hard drives, including hard drives that are removable.",
                win32file.DRIVE_REMOTE : "Remote\nNetwork drives. This includes drives shared anywhere on a network.",
                win32file.DRIVE_CDROM : "CDROM\nDrive is a CD-ROM. No distinction is made between read-only and read/write CD-ROM drives.",
                win32file.DRIVE_RAMDISK : "RAMDisk\nDrive is a block of random access memory (RAM) on the local computer that behaves like a disk drive.",
                win32file.DRIVE_NO_ROOT_DIR : "The root directory does not exist."
              }

drives = win32api.GetLogicalDriveStrings().split('\x00')[:-1]

for device in drives:
    type = win32file.GetDriveType(device)
    
#    print(colorama.Fore.WHITE +"Drive: %s" % device)
    
# отображение кол-ва дисков
c = wmi.WMI()

disk_devices = c.Win32_LogicalDisk()
print("Количество накопителей: %s" % len(disk_devices))
if len(disk_devices) > 3 :
    print(colorama.Fore.GREEN + "Storage_GOOD")
else:
    print(colorama.Fore.RED + "Storage_Fail")
 
#Проверка  микрофона сток
def check_microphone(input_data, frames, time, status):
 if any(input_data > 0):
    if len(input_data) > 3 :
       print(colorama.Fore.GREEN + "Micro_GOOD")
 else:
  print(colorama.Fore.RED + "Micro_Fail")

# Задаем параметры для записи звука
duration = 10  # Длительность записи в секундах
sample_rate = 44100  # Частота дискретизации звука

# Запускаем запись звука с микрофона
with sd.InputStream(callback=check_microphone, channels=1, samplerate=sample_rate):
    sd.sleep(duration * 5)
    
    # количество сетей Wi-Fi
def get_wifi_networks():
    try:
        result = subprocess.check_output(["netsh", "wlan", "show", "network"], encoding='cp866')
        networks = result.split("\n")
        return len(networks) - 6  # Исключаем заголовки и пустые строки
    except subprocess.CalledProcessError:
        return 0

# Получаем количество сетей Wi-Fi
count = get_wifi_networks()

if count > 5:
    message = colored("GOOD", "green")
elif count == 0:
    message = colored("BAD", "red")
else:
    message = str(count)

print(f"Wi-Fi {message}")
 
# проверка Ethernet
url = "http://yandex.ru/"

try:
    response = requests.get(url)
    if response.status_code == 200:
        print(colorama.Fore.GREEN + "Ethernet_Good")
    else:
        print(colorama.Fore.RED + "Ethernet_Bad")
except requests.ConnectionError:
    print("Bad")
    
# Bluetooth
# Получаем список сетей Bluetooth
devices = bluetooth.discover_devices()

# Получаем количество сетей Bluetooth
num_devices = len(devices)

# Проверяем количество сетей и выводим соответствующее сообщение
if num_devices > 0:
    print(colorama.Fore.GREEN + "Bluetooth_Good")
elif num_devices == 0:
    print(colorama.Fore.RED + "Bluetooth_Bad")


##### Проверка звука (человеком)    
import pyaudio
import wave
import keyboard

# Функция воспроизведения .wav файла
def play_wave(file_path):
    # Открытие файла
    wf = wave.open(file_path, 'rb')

    # Создание объекта PyAudio
    p = pyaudio.PyAudio()

    # Открытие потока
    stream = p.open(format=p.get_format_from_width(wf.getsampwidth()),
                    channels=wf.getnchannels(),
                    rate=wf.getframerate(),
                    output=True)

    # Чтение данных
    data = wf.readframes(1024)

    # Воспроизведение потока
    while len(data) > 0:
        stream.write(data)
        data = wf.readframes(1024)

    # Остановка потока и очистка ресурсов
    stream.stop_stream()
    stream.close()
    p.terminate()

# Путь к .wav файлу
file_path = 'C:/Users/NoFranxx JOB/Desktop/JOB_Programm/444/test/LRC.WAV'

print(colorama.Fore.WHITE + "Нажмите 'Y' для воспроизведения или 'Enter' для выхода.")

# Ожидание событий клавиатуры
while True:
    if keyboard.is_pressed('Y') or keyboard.is_pressed('y'):  # Если нажата 'Y'
        play_wave(file_path)
    if keyboard.is_pressed('Enter'):  # Если нажат Enter
        print(colorama.Fore.GREEN + "Volume_GOOD")
        break

#Проверка  микрофона сток
def check_microphone(input_data, frames, time, status):
 if any(input_data > 0):
    if len(input_data) > 3 :
       print(colorama.Fore.GREEN + "Micro_GOOD")
 else:
  print(colorama.Fore.RED + "Micro_Fail")

# Задаем параметры для записи звука
duration = 10  # Длительность записи в секундах
sample_rate = 44100  # Частота дискретизации звука

# Запускаем запись звука с микрофона
with sd.InputStream(callback=check_microphone, channels=1, samplerate=sample_rate):
    sd.sleep(duration * 5)



# Ваш остальной код здесь



## Проверка звука + Микро (Автоматически)
# Создаем объект Recognizer
#r = sr.Recognizer()
#pygame.init()
#pygame.mixer.music.load("LRC.mp3")

# Записываем звук с микрофона
#with sr.Microphone() as source:aaaa
#   print(colorama.Fore.WHITE + "Test speakers")
#    time.sleep(1)
#    pygame.mixer.music.play()
#
#   audio = r.listen(source)
    

#try:
    # Распознаем голос с помощью Google Web Speech API
#    text = r.recognize_google(audio, language="ru-RU")
#    print("Вы сказали: " + text)

#     Сравниваем распознанную фразу с эталонной фразой "left right" с помощью fuzzywuzzy
#    ratio = fuzz.ratio(text.lower(), "left right")
#    if ratio >= 80:  # Устанавливаем порог точности сравнения
#        print(colorama.Fore.GREEN + "Speakers_Micro_GOOD")
#except sr.UnknownValueError:
#    print("Не удалось распознать голос")
#except sr.RequestError as e:
#    print("Ошибка при запросе к сервису распознавания голоса: {0}".format(e))
    
input("")    
input("")    

#стабильный вариант на 26.01.2023.aa