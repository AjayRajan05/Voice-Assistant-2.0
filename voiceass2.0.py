import subprocess
import wolframalpha
import pyttsx3
import json
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import smtplib
import ctypes
import time
import requests
import getpass
import wmi
from pathlib import Path
from ecapture import ecapture as ec
from urllib.request import urlopen

# Initialize the speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)  # Use voices[0].id for male voice, voices[1].id for female voice

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning Sir!")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Sir!")   
    else:
        speak("Good Evening Sir!")  

    assname = "Asta 1 point o"
    speak("I am your assistant")
    speak(assname)

def usrname():
    speak("What should I call you, sir?")
    uname = cmd()
    speak("Welcome, Sir")
    speak(uname)
    print("#####################")
    print("Welcome", uname)
    print("#####################")

def cmd():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")
    except Exception as e:
        print(e)    
        print("Unable to Recognize your voice.")  
        return "None"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremail@email.com', 'your-password')
    server.sendmail('youremail@email.com', to, content)
    server.close()

def organize():
    FILE_FORMATS = {
        '.jpg': 'Images', '.jpeg': 'Images', '.png': 'Images', 
        '.pdf': 'Documents', '.doc': 'Documents', '.docx': 'Documents'
        # Add more formats and directories as needed
    }
    for entry in os.scandir():
        if entry.is_dir():
            continue
        file_path = Path(entry.name)
        file_format = file_path.suffix.lower()
        if file_format in FILE_FORMATS:
            directory_path = Path(FILE_FORMATS[file_format])
            directory_path.mkdir(exist_ok=True)
            file_path.rename(directory_path.joinpath(file_path))
    try:
        os.mkdir("OTHER")
    except:
        pass
    for dir in os.scandir():
        try:
            if dir.is_dir():
                os.rmdir(dir)
            else:
                os.rename(os.getcwd() + '/' + str(Path(dir)), os.getcwd() + '/OTHER/' + str(Path(dir)))
        except:
            pass

if __name__ == '__main__':
    clear = lambda: os.system('cls')
    clear()
    wishMe()
    usrname()
    speak("How can I help you, Sir?")
    
    while True:
        query = cmd().lower()
        assname = "Asta 1 point o"

        if "good morning" in query:
            speak("A warm " + query)
            speak("How are you Mister")
            speak(assname)

        elif 'open youtube' in query:
            speak("Taking you to YouTube\n")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("Taking you to Google\n")
            webbrowser.open("google.com")

        elif "change brightness to" in query:
            query = query.replace("change brightness to", "")
            brightness = int(query)
            c = wmi.WMI(namespace='wmi')
            methods = c.WmiMonitorBrightnessMethods()[0]
            methods.WmiSetBrightness(brightness, 0)

        elif "organize files" in query:
            organize()

        elif 'open geeksforgeeks' in query:
            speak("Here you go to GeeksforGeeks. Happy Learning!")
            webbrowser.open("geeksforgeeks.com")

        elif "send a whatsapp message" in query:
            driver = webdriver.Chrome('Path to Web Driver')
            driver.get('https://web.whatsapp.com/')
            speak("Scan QR code before proceeding")
            time.sleep(15)
            speak("Enter name of group or user")
            name = cmd()
            speak("Enter your message")
            msg = cmd()
            user = driver.find_element_by_xpath('//span[@title = "{}"]'.format(name))
            user.click()
            msg_box = driver.find_element_by_class_name('_3uMse')
            msg_box.send_keys(msg)
            button = driver.find_element_by_class_name('_1U1xa')
            button.click()

        elif 'play music' in query or "play song" in query:
            username = getpass.getuser()
            music_dir = f"C:\\Users\\{username}\\Music"
            songs = os.listdir(music_dir)
            print(songs)    
            os.startfile(os.path.join(music_dir, songs[0]))

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Sir, the time is {strTime}")

        elif 'open microsoft edge' in query:
            codePath = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
            os.startfile(codePath)

        elif 'email to sha' in query:
            try:
                speak("What should I say?")
                content = cmd()
                to = "sha04@gmail.com"    
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'how are you' in query:
            speak("I am fine, thank you")
            speak("How are you, Sir?")

        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query

        elif "shall i change your name" in query:
            speak("What would you like to call me, Sir?")
            assname = cmd()
            speak("Thanks for naming me, Sir")

        elif "what's your name" in query or "what is your name" in query:
            speak("My dear, call me")
            speak(assname)
            print("My dear, call me", assname)

        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()

        elif "who made you" in query or "who created you" in query: 
            speak("I have been created by Ajay Rajan")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif "calculate" in query:
            app_id = "Your WolframAlpha API ID"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)

        elif 'search' in query or 'play' in query: 
            query = query.replace("search", "") 
            query = query.replace("play", "")
            webbrowser.open(query)

        elif "who am i" in query:
            speak("If you talk then definitely you are a human.")

        elif "why were you created" in query:
            speak("Thanks to Ajay Rajan. To help others and be a real friend")

        elif 'what is love' in query:
            speak("It is the 7th sense that destroys all other senses")

        elif "who are you" in query:
            speak("I am your virtual assistant created by Ajay Rajan to help others")

        elif "what is the reason for creating you" in query:
            speak("As a mini project to help Ajay")

        elif 'google news' in query:
            try:
                jsonObj = urlopen('''https://newsapi.org/v2/top-headlines?sources=google-news-in&apiKey=Your-API-Key''')
                data = json.load(jsonObj)
                i = 1
                speak('Here are some top news from Google News')
                print('''=============== Google News ============''' + '\n')
                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                print(str(e))

        elif 'lock window' in query:
            speak("Locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
            speak("Hold on a sec! Your system is on its way to shut down")
            subprocess.call('shutdown /p /f')

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin emptied")

        elif "don't listen" in query or "stop listening" in query:
            speak("For how much time you want to stop me from listening commands?")
            a = int(cmd())
            time.sleep(a)
            print(a)

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl/maps/place/" + location + "")

        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Asta Camera", "img.jpg")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown /h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "write a note" in query:
            speak("What should I write, Sir?")
            note = cmd()
            file = open('asta.txt', 'w')
            speak("Sir, Should I include date and time?")
            snfm = cmd()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
            file.close()

        elif "show note" in query:
            speak("Showing Notes")
            file = open("asta.txt", "r")
            print(file.read())
            speak(file.read(6))

        elif "wikipedia" in query:
            speak("Searching Wikipedia...")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=3)
            speak("According to Wikipedia")
            print(results)
            speak(results)
