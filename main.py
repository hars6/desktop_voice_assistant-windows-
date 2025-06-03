import speech_recognition as sr
import pyttsx3
import subprocess
from AppOpener import open, close
import webbrowser
import datetime
import socket
import speedtest
from vosk import Model, KaldiRecognizer
import pyaudio
import urllib.request
import re
import json
from urllib.request import urlopen
import os



def chk_cnction():
    try:
        urllib.request.urlopen("https://www.google.com")
        return True
    except:
        return False

def open_lnk(str):
    first, *middle, last = str.split()
    otpt_spkr(f"opening {last}")
    webbrowser.open(f'www.{last}.com')

def lction():
    url='http://ipinfo.io/json'
    response=urlopen(url)
    data=json.load(response)
    # print(data)
    print(data["city"],",",data["region"])
    return data["city"],",",data["region"]

def cal_addp(str):
    arr=re.findall(r'[\d]+[.,\d]+|[\d]*[.][\d]+|[\d]+',str)
    arr = list(map(float, arr))
    tot=0
    for i in range(0,len(arr)):
        tot+=arr[i]
    print("The added value is: ",tot)
    return tot

def cal_subp(str):
    arr=re.findall(r'[\d]+[.,\d]+|[\d]*[.][\d]+|[\d]+',str)
    arr = list(map(float, arr))
    tot=0
    for i in range(0,len(arr)-1):
        if i==0:
            tot=arr[i]-arr[i+1]
        else:
            tot-=arr[i+1]
    print("The subtracted value is: ",tot)
    return tot

def cal_mulp(str):
    arr=re.findall(r'[\d]+[.,\d]+|[\d]*[.][\d]+|[\d]+',str)
    arr = list(map(float, arr))
    tot=None
    for i in range(0,len(arr)-1):
        if i==0:
            tot=arr[i]*arr[i+1]
        else:
            tot*=arr[i+1]
    print("The multiplied value is: ",tot)
    return tot

def cal_dvd(str):
    arr=re.findall(r'[\d]+[.,\d]+|[\d]*[.][\d]+|[\d]+',str)
    arr = list(map(float, arr))
    tot=None
    for i in range(0,len(arr)-1):
        if i==0:
            tot=arr[i]/arr[i+1]
        else:
            tot/=arr[i+1]
    tot = round(tot, 2)
    print("The divided value is: ",tot)
    return tot

def otpt_spkr(txt_say=None):
    saiy = pyttsx3.init()
    voices = saiy.getProperty('voices')
    saiy.setProperty('voice', voices[2].id)
    saiy.say(txt_say)
    saiy.runAndWait()


def gugl_spch_rcg():
    # google speach recognizer
    global cnd
    cnd=True
    while cnd==True:
        r = sr.Recognizer()
        with sr.Microphone() as mic:
            r.adjust_for_ambient_noise(mic, duration=0.5)
            if not chk_cnction():
                vosk_spch_rcg()
            print("Listening....")
            audio = r.listen(mic, phrase_time_limit=6)
            global text
            text=" "
            try:
                text = r.recognize_google(audio)
                text = text.lower()
                print(f"USER: {text}")
            except:
                pass
            if text == "stop the assistant" or text == "stop assistant":
                print("ASSISTANT STOPPED!!!")
                otpt_spkr("thankyou for using the assistant, assistant stopped")
                text=" "
                cnd=False
            if not text==" ":
                do_tsk()


def vosk_spch_rcg():
    # vosk speech recognizer
    global text, cnd
    while cnd==True:
        if chk_cnction():
            gugl_spch_rcg()
        model = Model("./Vosk/vosk-model-small-en-us-0.15")
        recognizer = KaldiRecognizer(model, 16000)

        mic = pyaudio.PyAudio()
        stream = mic.open(rate=16000, channels=1, format=pyaudio.paInt16, input=True, frames_per_buffer=8192)
        stream.start_stream
        chk = True
        print("Listening..")
        while chk == True:
            if chk_cnction():
                gugl_spch_rcg()
                chk = False
            data = stream.read(4096)
            if recognizer.AcceptWaveform(data):
                txt = recognizer.Result()
                text = txt[14:-3]
                if not text==" ":
                    print(f"USER(Offline recognition): {text}")
                if text == "turn on internet" or text == "turned on internet" or text == "don't have internet" or text == "internet":
                    otpt_spkr("turning wifi on")
                    subprocess.run(['netsh', 'interface', 'set', 'interface', 'Wi-Fi', 'ENABLED'])
                    chk = False
                    text="-"
                    # print(text)
                if text == "stop the assistant" or text == "stop assistant" or text=="stop":
                    print("ASSISTANT STOPPED!!!")
                    otpt_spkr("thankyou for using the assistant, assistant stopped")
                    chk=False
                    cnd = False
                    text = " "
                if not text == " ":
                    do_tsk()


def do_tsk():
    global text
    if (text == "open chrome" or text == "open browser"):
        otpt_spkr("opening chrome")
        open('chrome', match_closest=True, output=False)
        text = " "
    elif (text == "close chrome" or text == "close browser"):
        otpt_spkr("closing chrome")
        close('chrome', match_closest=True, output=False)
        text = " "
    elif (text == "open telegram"):
        otpt_spkr("opening telegram")
        open('telegram', match_closest=True, output=False)
        text = " "
    elif (text == "close telegram"):
        otpt_spkr("closing telegram")
        close('telegram', match_closest=True, output=False)
        text = " "
    elif (text == "open ms word" or text == "open word"):                       # microsoft office softwares
        otpt_spkr("opening ms word")
        open('msword', match_closest=True, output=False)
        text = ""
    elif (text == "close ms word" or text == "close word"):
        otpt_spkr("closing ms word")
        subprocess.call(["taskkill", "/F", "/IM", "winword.exe"])
        text = ""
    elif (text == "open ms excel" or text == "open excel"):
        otpt_spkr("opening ms excel")
        open('excel', match_closest=True, output=False)
        text = ""
    elif (text == "close ms excel" or text == "close excel"):
        otpt_spkr("closing ms excel")
        subprocess.call(["taskkill", "/F", "/IM", "excel.exe"])
        text = ""
    elif (text == "open ms powerpoint" or text == "open powerpoint"):
        otpt_spkr("opening ms powerpoint")
        open('powerpoint', match_closest=True, output=False)
        text = ""
    elif (text == "close ms powerpoint" or text == "close powerpoint"):
        otpt_spkr("closing ms powerpoint")
        subprocess.call(["taskkill", "/F", "/IM", "powerpnt.exe"])
        text = ""
    elif (text == "open onenote" or text=="open ms onenote" or text=="open microsoft onenote"):
        otpt_spkr("opening microsoft onenote")
        open('onenote', match_closest=True, output=False)
        text = ""
    elif (text == "close one note" or text=="close ms one note" or text=="close microsoft one note"):
        otpt_spkr("closing microsoft onenote")
        subprocess.call(["taskkill", "/F", "/IM", "onenote.exe"])
        text = ""
    elif (text == "open outlook" or text=="open ms outlook" or text=="open microsoft outlook"):
        otpt_spkr("opening microsoft outlook")
        open('outlook', match_closest=True, output=False)
        text = ""
    elif (text == "close outlook" or text=="close ms outlook" or text=="close microsoft outlook"):
        otpt_spkr("closing microsoft outlook")
        subprocess.call(["taskkill", "/F", "/IM", "outlook.exe"])
        text = ""
    elif (text == "open onedrive"):                                                     # it can be closed through file explorer
        otpt_spkr("opening microsoft onedrive")
        open('onedrive', match_closest=True, output=False)
        text = ""
    elif (text == "open vlc" or text == "open vlc media player" or text == "open media player"):
        otpt_spkr("opening vlc media player")
        open('vlc media player', match_closest=True, output=False)
        text = ""
    elif (text == "close vlc" or text == "close vlc media player" or text == "close media player"):
        otpt_spkr("closing vlc media player")
        close('vlc', match_closest=True, output=False)
        text = ""
    elif (text == "open file explorer"):                                                # system apps
        otpt_spkr("opening file explorer")
        open('file explorer', match_closest=True, output=False)
        text = ""
    elif (text == "close file explorer"):
        otpt_spkr("closing file explorer")
        close('file explorer', match_closest=True, output=False)
        text = ""
    elif (text == "open notepad"):
        otpt_spkr("opening notepad")
        open('notepad', match_closest=True, output=False)
        text = ""
    elif (text == "close notepad"):
        otpt_spkr("closing notepad")
        close('notepad', match_closest=True, output=False)
        text = ""
    elif (text == "open calculator"):
        otpt_spkr("opening calculator")
        open('calculator', match_closest=True, output=False)
        text = ""
    elif (text == "close calculator"):
        otpt_spkr("closing calculator")
        close('calculator', match_closest=True, output=False)
        text = ""
    elif (text == "open camera"):
        otpt_spkr("opening camera")
        open('camera', match_closest=True, output=False)
        text = ""
    elif (text == "close camera"):
        otpt_spkr("closing camera")
        subprocess.call(["taskkill", "/F", "/IM", "Windowscamera.exe"])
        text = ""
    elif (text == "open task manager"):
        otpt_spkr("opening task manager")
        open('task manager', match_closest=True, output=False)
        text = ""
    elif (text == "close task manager"):
        otpt_spkr("closing task manager")
        close('task manager', match_closest=True, output=False)
        # subprocess.call(["taskkill","/F","/IM","taskmgr.exe"])
        text = ""
    elif (text == "open settings"):
        otpt_spkr("opening settings")
        open('settings', match_closest=True, output=False)
        text = ""
    elif (text == "close settings"):
        otpt_spkr("closing settings")
        close('settings', match_closest=True, output=False)
        text = ""
    elif (text == "open this pc"):                                                  # it can be closed through file explorer
        otpt_spkr("opening this pc")
        open('this pc', match_closest=True, output=False)
        text = ""
    elif (text == "open control panel"):
        otpt_spkr("opening control panel")
        open('control panel', match_closest=True, output=False)
        text = ""
    elif (text == "open paint"):
        otpt_spkr("opening paint")
        open('paint', match_closest=True, output=False)
        text = ""
    elif (text == "close paint"):
        otpt_spkr("closing paint")
        close('paint', match_closest=True, output=False)
        text = ""
    elif (text == "open snipping tool"):
        otpt_spkr("opening snipping tool")
        open('snipping tool', match_closest=True, output=False)
        text = ""
    elif (text == "close snipping tool"):
        otpt_spkr("closing snipping tool")
        close('snipping tool', match_closest=True, output=False)
        text = ""
    elif (text == "open powershell"):
        otpt_spkr("opening powershell")
        open('powershell', match_closest=True, output=False)
        text = ""
    elif (text == "close powershell"):
        otpt_spkr("closing powershell")
        close('powershell', match_closest=True, output=False)
        text = ""
    elif (text == "open voice recorder"):
        otpt_spkr("opening voice recorder")
        open('voice recorder', match_closest=True, output=False)
        text = ""
    elif (text == "close voice recorder"):
        otpt_spkr("closing voice recorder")
        close('soundrec', match_closest=True, output=False)
        text = ""
    elif (text == "open weather"):
        otpt_spkr("opening weather software")
        open('weather', match_closest=True, output=False)
        text = ""
    elif (text == "close weather"):
        otpt_spkr("closing weather software")
        close('microsoft.msn.weather', match_closest=True, output=False)
        text = ""
    elif (text == "open maps"):
        otpt_spkr("opening maps")
        open('maps', match_closest=True, output=False)
        text = ""
    elif (text == "close maps"):
        otpt_spkr("closing maps")
        close('maps', match_closest=True, output=False)
        text = ""
    elif (text == "open sticky notes"):
        otpt_spkr("opening sticky notes")
        open('sticky notes', match_closest=True, output=False)
        text = ""
    elif (text == "close sticky notes"):
        otpt_spkr("closing sticky notes")
        close('microsoft.notes', match_closest=True, output=False)
        text = ""
    elif (text == "open solitaire game"):
        otpt_spkr("opening solitaire game")
        open('solitaire & casual game', match_closest=True, output=False)
        text = ""
    elif (text == "close solitaire game"):
        otpt_spkr("closing solitaire game")
        close('solitaire & casual game', match_closest=True, output=False)
        text = ""
    elif (text == "open photos"):
        otpt_spkr("opening photos")
        open('photos', match_closest=True, output=False)
        text = ""
    elif (text == "close photos"):
        otpt_spkr("closing photos")
        close('photos', match_closest=True, output=False)
        text = ""
    elif (text == "open spotify"):
        otpt_spkr("opening spotify")
        open('spotify', match_closest=True, output=False)
        text = ""
    elif (text == "close spotify"):
        otpt_spkr("closing spotify")
        close('spotify', match_closest=True, output=False)
        text = ""
    elif (text == "open clock"):
        otpt_spkr("opening clock")
        open('clock', match_closest=True, output=False)
        text = ""
    elif (text == "close clock"):
        otpt_spkr("closing clock")
        close('time', match_closest=True, output=False)
        text = ""
    elif (text == "open command prompt" or text=="open cmd"):
        otpt_spkr("opening command prompt")
        open('command prompt', match_closest=True, output=False)
        text = ""
    elif (text == "close command prompt" or text=="close cmd"):
        otpt_spkr("closing command prompt")
        close('command prompt', match_closest=True, output=False)
        text = ""
    elif (text == "open task manager"):
        otpt_spkr("opening task manager")
        open('task manager', match_closest=True, output=False)
        text = ""
    elif (text == "close task manager"):
        otpt_spkr("closing task manager")
        close('task manager', match_closest=True, output=False)
        text = ""
    elif (text == "open recycle bin"):  # can be closed by file explorer
        otpt_spkr("opening recycle bin")
        open('recycle bin', match_closest=True, output=False)
        text = ""
    elif (text == "open vs code" or text=="open visual studio code"):               # non system apps
        otpt_spkr("opening visual studio code")
        open('visual studio code', match_closest=True, output=False)
        text = ""
    elif (text == "close vs code" or text=="close visual studio code"):
        otpt_spkr("closing visual studio code")
        close('code', match_closest=True, output=False)
        text = ""
    elif (text == "open edge" or text=="open microsoft edge"):
        otpt_spkr("opening microsoft edge")
        open('microsoft edge', match_closest=True, output=False)
        text = ""
    elif (text == "close edge" or text=="close microsoft edge"):
        otpt_spkr("closing microsoft edge")
        close('microsoft edge', match_closest=True, output=False)
        text = ""
    elif (text == "open libreoffice"):
        otpt_spkr("opening libreoffice")
        open('libreoffice', match_closest=True, output=False)
        text = ""
    elif (text == "close libreoffice"):
        otpt_spkr("closing libreoffice")
        close('soffice', match_closest=True, output=False)
        text = ""
    elif (text == "open eclipse ide" or text=="open eclipse"):
        otpt_spkr("opening eclipse ide for java developers")
        open('eclipse ide for java developers', match_closest=True, output=False)
        text = ""
    elif (text == "close eclipse" or text == "close eclipse ide"):
        otpt_spkr("closing eclipse ide for java developers")
        close('eclipse ide for java developers', match_closest=True, output=False)
        text = ""
    elif (text == "open nearby share"):
        otpt_spkr("opening nearby share beta from google")
        open('nearby share beta from google', match_closest=True, output=False)
        text = ""
    elif (text == "close nearby share"):
        otpt_spkr("closing nearby share beta from google")
        close('nearby share beta from google', match_closest=True, output=False)
        text = ""
    elif (text == "open python idle" or text=="open python"):                       # it cant close because it'll stop assistant
        otpt_spkr("opening python idle")
        open('idle python', match_closest=True, output=False)
        text = ""
    elif "open site".lower() in text.lower() or "open website".lower() in text.lower():              # website opening
        open_lnk(text)
        text = ""
    elif (text == "open google drive"):
        otpt_spkr("opening google drive")
        open('google drive', match_closest=True, output=False)
        text = ""
    # elif (text == "open youtube"):
    #     otpt_spkr("opening youtube")
    #     webbrowser.open('https://youtube.com')
    #     text = ""
    elif (text == "turn on internet" or text == "turned on internet"):              # WIFI toggle
        otpt_spkr("turning wifi on")
        subprocess.run(['netsh', 'interface', 'set', 'interface', 'Wi-Fi', 'ENABLED'])
        text = ""
    elif (text == "turn off wi-fi" or text == "turn off internet" or text == "don't have internet"):
        otpt_spkr("turning wifi off")
        subprocess.run(['netsh', 'interface', 'set', 'interface', 'Wi-Fi', 'DISABLED'])
        text = ""
    elif (text == "what is the current time" or text == "what is the time"):    # date & time part
        tim = datetime.datetime.now().strftime("%Ihours:%Mminutes:%Sseconds %p")
        print(tim)
        otpt_spkr(tim)
        text = ""
    elif (text == "what is the date of today" or text == "what is the date"):
        dt = datetime.datetime.now().strftime("%d %B %Y")
        print(dt)
        otpt_spkr(dt)
        text = ""
    elif (text == "which day is today"):
        dt = datetime.datetime.now().strftime("%A")
        print(dt)
        otpt_spkr(dt)
        text = ""
    elif (text == "what is my ip address" or text == "what is my ip"):              # checking IP address
        hostname = socket.gethostname()
        IPAddr = socket.gethostbyname(hostname)
        print(IPAddr)
        otpt_spkr(IPAddr)
        text = ""
    elif (text == "what is the speed" or text == "what is the internet speed"):     # checking internet speed
        speed_test = speedtest.Speedtest(secure=True)
        otpt_spkr("Let me check this will take few seconds")
        print("Checking....")
        download_speed = speed_test.download()
        upload_speed = speed_test.upload()
        download_speed /= 1000000
        download_speed = round(download_speed, 2)
        upload_speed /= 1000000
        upload_speed = round(upload_speed, 2)
        print(f"Download speed is: {download_speed}mbps")
        print(f"Upload speed is: {upload_speed}mbps")
        otpt_spkr(f"download speed is, {download_speed}mbps, and, upload speed is {upload_speed}mbps")
        text = ""
    elif "location" in text.lower():                                                # getting device location
        otpt_spkr(f"the location is {lction()}")
        text=""
    elif "add" in text.lower() or "plus" in text.lower() or "+" in text.lower():    # Calculation part
        otpt_spkr(f"the added value is {cal_addp(text)}")
        text=""
    elif "subtract" in text.lower() or "minus" in text.lower():
        otpt_spkr(f"the subtracted value is {cal_subp(text)}")
        text=""
    elif "multiply" in text.lower() or "x" in text.lower() or "into" in text.lower():
        otpt_spkr(f"the multiplied value is {cal_mulp(text)}")
        text=""
    elif "divide" in text.lower() or "by" in text.lower():
        otpt_spkr(f"the divided value is {cal_dvd(text)}")
        text=""


if __name__ == '__main__':
    try:
        gugl_spch_rcg()
    except:
        pass




    # print(os.getlogin())

    # subprocess.run(['netsh', 'interface', 'set', 'interface', 'Wi-Fi', 'ENABLED']) # to turn on wifi

    # subprocess.Popen(['C:\Program Files\Google\Chrome\Application\chrome.exe']) # to open app

    # subprocess.call(["taskkill","/F","/IM","chrome.exe"]) # to close app

    # speech=pyttsx3.init() # for saying anything
    # speech.say("Hello")
    # speech.runAndWait()

#                               \FOR CHECKING AVAILABLE VOICES IN SYSTEM/
#     engine = pyttsx3.init()
#     voices = engine.getProperty('voices')
#     for voice in voices:
#         print(voice, voice.id)
#         engine.setProperty('voice', voices[3].id)
#         engine.say("Hello World!")
#         engine.runAndWait()
#         engine.stop()
#     DAVID 0
#     KALPANAM 1
#     HEERAM 2
#     ZIRA 3
