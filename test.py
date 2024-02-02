import psutil
import colorama
import os
import wmi
import pyaudio
import time
import sounddevice as sd
import subprocess
from termcolor import colored
import requests
import wave
import keyboard


#Начало записи времени работы
start = time.time()



# Автор кода 
print(colorama.Fore.MAGENTA + "After: NoFranxx")



#Проверка объёма оперативной памяти
memory_usage = psutil.virtual_memory().total

ram_size = psutil.virtual_memory().total  # total RAM size in bytes

if ram_size > 12 * 10 ** 9: 
    print(colorama.Fore.GREEN + "Memory_GOOD")
else:
    print(colorama.Fore.RED + "Memory_Fail")

# отображение кол-ва дисков
c = wmi.WMI()

disk_devices = c.Win32_LogicalDisk()
print("Количество накопителей: %s" % len(disk_devices))
if len(disk_devices) > 4 :
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
duration = 10
sample_rate = 44100

# Запускаем запись звука с микрофона
with sd.InputStream(callback=check_microphone, channels=1, samplerate=sample_rate):
    sd.sleep(duration * 5)
    
 
 
 
 
 # количество сетей Wi-Fi
def get_wifi_networks():
    try:
        result = subprocess.check_output(["netsh", "wlan", "show", "network"], encoding='cp866')
        networks = result.split("\n")
        return len(networks) - 6  
    except subprocess.CalledProcessError:
        return 0


# Получаем количество сетей Wi-Fi
count = get_wifi_networks()

if count > 1:
    message = colored("GOOD", "green")
elif count == 0:
    message = colored("BAD", "red")
else:
    message = str(count)

print(f"Wi-Fi {message}")
 



# проверка Ethernet AOI (yandex.ru)
#url = "https://10.255.90.10/login/"
url = "https://yandex.ru/"

try:
    response = requests.get(url)
    if response.status_code == 200:
        print(colorama.Fore.GREEN + "Ethernet_Good")
    else:
        print(colorama.Fore.RED + "Ethernet_Bad")
except requests.ConnectionError:
    print(colorama.Fore.RED + "Ethernet_Bad")

    


    # проверка Ethernet server
url = "https://10.255.90.11/"

try:
    response = requests.get(url)
    if response.status_code == 200:
        print(colorama.Fore.GREEN + "Ethernet_Good")
    else:
        print(colorama.Fore.RED + "Ethernet_Bad")
except requests.ConnectionError:
    print(colorama.Fore.RED + "Ethernet_Bad")




# Bluetooth
import subprocess
exe_file_path = 'C:/test/test_BT.exe'
result = subprocess.run(exe_file_path, capture_output=True, text=True)

# Проверяем, выдаёт ли файл сообщение "1,2,3" (Это количество сетей bluetooth, которые нашёл файл test_BT.exe)
output = result.stdout.strip()
if output == '1':
    print(colorama.Fore.GREEN + 'Bluetooth_Good')
elif output == '2' or output == '3':
    print(colorama.Fore.GREEN + 'Bluetooth_Good')
else:
   print(colorama.Fore.RED + 'Bluetooth_Bad')






# Проверка звука (человеком)    
def play_wave(file_path):
    wf = wave.open(file_path, 'rb')

    p = pyaudio.PyAudio()
    stream = p.open(format=p.get_format_from_width(wf.getsampwidth()),
                    channels=wf.getnchannels(),
                    rate=wf.getframerate(),
                    output=True)
    data = wf.readframes(1024)

    while len(data) > 0:
        stream.write(data)
        data = wf.readframes(1024)

    stream.stop_stream()
    stream.close()
    p.terminate()

file_path = 'C:/test/LRC.WAV'

print(colorama.Fore.WHITE + "Нажмите 'Y' для воспроизведения звука, нажмите 'Enter' для продолжения.")

# Ожидание событий клавиатуры
while True:
    if keyboard.is_pressed('Y') or keyboard.is_pressed('y'):  # Если нажата 'Y'
        play_wave(file_path)
    if keyboard.is_pressed('Enter'):  # Если нажат Enter
        print(colorama.Fore.GREEN + "Volume_GOOD")
        break



#Проверка  микрофона наушники
def check_microphone(input_data, frames, time, status):
 if any(input_data > 0):
    if len(input_data) > 3 :
       print(colorama.Fore.GREEN + "Micro_GOOD")
 else:
  print(colorama.Fore.RED + "Micro_Fail")

duration = 10
sample_rate = 44100

with sd.InputStream(callback=check_microphone, channels=1, samplerate=sample_rate):
    sd.sleep(duration * 5)

input(colorama.Fore.BLACK + "")



# Ввод серийного номера
user_input = input(colorama.Fore.WHITE + "Введите Serial Number: ")

with open("C:/test/test.txt", "a") as file:
    file.write("\n" + "Валерий_test: " + user_input)

print(colorama.Fore.GREEN + "Серийный номер сохранён в test.txt.")

## Проверка звука + Микро (Автоматически)                       BETA
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
#Распознаем голос с помощью Google Web Speech API
#    text = r.recognize_google(audio, language="ru-RU")
#    print("Вы сказали: " + text)

#Сравниваем распознанную фразу с эталонной фразой "left right" 
#    ratio = fuzz.ratio(text.lower(), "left right")
#    if ratio >= 80:  # Устанавливаем порог точности сравнения
#        print(colorama.Fore.GREEN + "Speakers_Micro_GOOD")
#except sr.UnknownValueError:
#    print("Не удалось распознать фразу")
#except sr.RequestError as e:
#    print("Ошибка при запросе к сервису распознавания голоса: {0}".format(e))




# Окончание работы программы/запись времени работы
end = time.time()
print(colorama.Fore.BLACK + "Execution time of the program is- ", end-start)

with open("C:/test/test_time.txt", "a") as file:
    file.write("\n" + "Валерий_test: " + str(end-start))

print(colorama.Fore.BLACK + "Время работы записано в test_time.txt.")

input("")    

#стабильный вариант на 29.01.2023. (Готовый) (не забудь добавить в "C:/test/" все необходимые файлы: test_time.txt; test.txt; LRC.WAV )



#Необходимо добавить 

#Автозагрузку в системе
#Автовыключение системы после теста, если он пройден успешно
#Добавить ошибку в систему (автоматически), отдельный excel файл
#Ускорить время работы
#Исправить bag ethernet порта (необходимо протестировать именно к адресу сервера и aoi)
#Изменить проверку микрофона, считаю, что можно использовать фразу, к примеру "test"
