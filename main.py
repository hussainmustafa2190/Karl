from Karl import KarlAssistant
import re
import os
import random
import pprint
import datetime
import requests
import sys
import urllib.parse  
import pyjokes
import time
import pyautogui
import pywhatkit
import wolframalpha
import pickle
from PIL import Image
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import QTimer, QTime, QDate, Qt
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from Karl.features.gui import Ui_MainWindow
from Karl.config import config
import win32com.client as win
import subprocess
from Karl.features import bot1 
from Karl.features.bot1 import NeighborSampler


obj = KarlAssistant()

# ================================ MEMORY ===========================================================================================================

GREETINGS = ["hello Karl", "Karl", "wake up Karl", "you there Karl", "time to work Karl", "hey Karl",
             "ok Karl", "are you there"]
GREETINGS_RES = ["always there for you sir", "i am ready sir",
                 "your wish my command", "how can i help you sir?", "i am online and ready sir"]

EMAIL_DIC = {
    'myself': 'hussainmustafa2190@gmail.com',
    'my official email': 'mhussain3032@gmai.com',
    'my second email': 'hussainmustafa2190@gmail.com',
}

CALENDAR_STRS = ["what do i have", "do i have plans", "am i busy"]
# =======================================================================================================================================================


def speak(text):
    obj.tts(text)


app_id = config.wolframalpha_id


def computational_intelligence(question):
    try:
        client = wolframalpha.Client("EQLEWU-XVQKH8YTR2")
        answer = client.query(question)
        answer = next(answer.results).text
        print(answer)
        return answer
    except:
        speak("Sorry sir I couldn't fetch your question's answer. Please try again ")
        return None
    
def startup():
    """speak("Initializing Karl")
    speak("Starting all systems applications")
    speak("Installing and checking all drivers")
    speak("Caliberating and examining all the core processors")
    speak("Checking the internet connection")
    speak("Wait a moment sir")
    speak("All drivers are up and running")
    speak("All systems have been activated")
    speak("Now I am online")
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<=12:
        speak("Good Morning")
    elif hour>12 and hour<18:
        speak("Good afternoon")
    else:
        speak("Good evening")
    c_time = obj.tell_time()
    speak(f"Currently it is {c_time}")
    speak("I am Karl. Online and ready sir. Please tell me how may I help you")"""
    



def wish():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<=12:
        speak("Good Morning")
    elif hour>12 and hour<18:
        speak("Good afternoon")
    else:
        speak("Good evening")
    c_time = obj.tell_time()
    speak(f"Currently it is {c_time}")
    #speak("I am Karl. Online and ready sir. Please tell me how may I help you")
# if __name__ == "__main__":


class MainThread(QThread):
    def __init__(self):
        super(MainThread, self).__init__()

    def run(self):
        self.TaskExecution()

    def TaskExecution(self):
        startup()
        wish()

        while True:
            command = obj.mic_input()

            if re.search('date', command):
                date = obj.tell_me_date()
                print(date)
                speak(date)
                continue
            
            elif command in ("who made you","who created you","who is your creator","what is the name of your creator"):
                speak("I was Created by Genesis")
                continue
            
            
            elif command in ("who is genesis","what is genesis"):
                speak("genesis are the master  minds that created me they are Mustafa, Mohtashim and Pritam")
                continue    
            
            elif "time" in command:
                time_c = obj.tell_time()
                print(time_c)
                speak(f"Sir the time is {time_c}")
                continue

            elif re.search('launch', command):
                dict_app = {
                    'chrome': 'C:/Program Files/Google/Chrome/Application/chrome',
                    'powerpoint':'C:/Program Files/Microsoft Office/root/Office16/POWERPNT',
                    'excel':'C:/Program Files/Microsoft Office/root/Office16/EXCEL',
                    'word':'C:/Program Files/Microsoft Office/root/Office16/WINWORD',
                    'outlook':'C:/Program Files/Microsoft Office/root/Office16/OUTLOOK',
                    'one note':'C:/Program Files/Microsoft Office/root/Office16/ONENOTE',
                    'windows media player':'C:/Program Files/Windows Media Player\wmplayer',
                    'firefox':'C:/Program Files/Mozilla Firefox/firefox',
                    'notepad':'C://Program Files (x86)//Notepad++//notepad++',
                    'internet explorer':'C:/Program Files/Internet Explorer/iexplore'
                }

                app = command.split(' ', 1)[1]
                path = dict_app.get(app)

                if path is None:
                    speak('Application path not found')
                    print('Application path not found')
                    continue

                else:
                    speak('Launching: ' + app + 'for you sir!')
                    obj.launch_any_app(path_of_app=path)
                    continue

            elif command in GREETINGS:
                speak(random.choice(GREETINGS_RES))
                continue

            elif re.search('open', command):
                domain = command.split(' ')[-1]
                open_result = obj.website_opener(domain)
                speak(f'Alright sir !! Opening {domain}')
                print(open_result)
                continue

            elif re.search('weather', command):
                city = command.split(' ')[-1]
                weather_res = obj.weather(city=city)
                print(weather_res)
                speak(weather_res)
                continue

            elif re.search('tell me about', command):
                topic = command.split(' ')[-1]
                if topic:
                    wiki_res = obj.tell_me(topic)
                    print(wiki_res)
                    speak(wiki_res)
                    continue
                else:
                    speak(
                        "Sorry sir. I couldn't load your query from my database. Please try again")
                    continue

            elif "buzzing" in command or "news" in command or "headlines" in command:
                news_res = obj.news()
                #speak('Source: The Times Of India')
                speak('Todays Top Headlines are..')
                for index, articles in enumerate(news_res):
                    pprint.pprint(articles['title'])
                    speak(articles['title'])
                    if index == 4 :
                        break
                speak('These were the top headlines, Have a nice day Sir!!..')
                continue

            elif 'search google for' in command:
                obj.search_anything_google(command)
                continue
            
            elif "play music" in command or "hit some music" in command:
                music_dir = "F://Songs//Imagine_Dragons"
                songs = os.listdir(music_dir)
                for song in songs:
                    os.startfile(os.path.join(music_dir, song))
                continue

            elif 'youtube' in command:
                video = command.split(' ')[1]
                speak(f"Okay sir, playing {video} on youtube")
                pywhatkit.playonyt(video)
                continue

            elif "email" in command or "send email" in command:
                sender_email = config.email
                sender_password = config.email_password
                

                try:
                    speak("Whom do you want to email sir ?")
                    recipient = obj.mic_input()
                    receiver_email = EMAIL_DIC.get(recipient)
                    if receiver_email:

                        speak("What is the subject sir ?")
                        subject = obj.mic_input()
                        speak("What should I say?")
                        message = obj.mic_input()
                        msg = 'Subject: {}\n\n{}'.format(subject, message)
                        obj.send_mail(sender_email, sender_password,
                                      receiver_email, msg)
                        speak("Email has been successfully sent")
                        time.sleep(2)
                        continue

                    else:
                        speak(
                            "I coudn't find the requested person's email in my database. Please try again with a different name")
                        continue

                except:
                    speak("Sorry sir. Couldn't send your mail. Please try again")
                    continue
            elif "calculate" in command:
                question = command
                answer = computational_intelligence(question)
                speak(answer)
                continue
            
            elif "what is" in command or "who is" in command:
                question = command
                answer = computational_intelligence(question)
                speak(answer)
                continue

            #elif "what do i have" in command or "do i have plans" or "am i busy" in command:
             #   obj.google_calendar_events(command)

            if "make a note" in command or "write this down" in command or "remember this" in command:
                speak("What would you like me to write down?")
                note_text = obj.mic_input()
                obj.take_note(note_text)
                speak("I've made a note of that")
                continue

            elif "close the note" in command or "close notepad" in command:
                speak("Okay sir, closing notepad")
                os.system("taskkill /f /im notepad++.exe")
                continue

            if "joke" in command:
                joke = pyjokes.get_joke()
                print(joke)
                speak(joke)
                continue

            elif "system" in command:
                sys_info = obj.system_info()
                print(sys_info)
                speak(sys_info)
                continue

            elif "where is" in command:
                place = command.split('where is ', 1)[1]
                current_loc, target_loc, distance = obj.location(place)
                city = target_loc.get('city', '')
                state = target_loc.get('state', '')
                country = target_loc.get('country', '')
                time.sleep(1)
                try:

                    if city:
                        res = f"{place} is in {state} state and country {country}. It is {distance} km away from your current location"
                        print(res)
                        speak(res)
                        continue

                    else:
                        res = f"{state} is a state in {country}. It is {distance} km away from your current location"
                        print(res)
                        speak(res)
                        continue

                except:
                    res = "Sorry sir, I couldn't get the co-ordinates of the location you requested. Please try again"
                    speak(res)
                    continue
                
            elif "ip address" in command:
                ip = requests.get('https://api.ipify.org').text
                print(ip)
                speak(f"Your ip address is {ip}")
                continue

            elif "switch the window" in command or "switch window" in command:
                speak("Okay sir, Switching the window")
                pyautogui.keyDown("alt")
                pyautogui.press("tab")
                time.sleep(1)
                pyautogui.keyUp("alt")
                continue

            elif "where i am" in command or "current location" in command or "where am i" in command:
                try:
                    city, state, country = obj.my_location()
                    print(city, state, country)
                    speak(
                        f"You are currently in {city} city which is in {state} state and country {country}")
                    continue
                except Exception as e:
                    speak(
                        "Sorry sir, I coundn't fetch your current location. Please try again")
                    continue

            elif "take screenshot" in command or "take a screenshot" in command or "capture the screen" in command:
                speak("By what name do you want to save the screenshot?")
                name = obj.mic_input()
                speak("Alright sir, taking the screenshot")
                img = pyautogui.screenshot()
                name = f"{name}.png"
                img.save(name)
                speak("The screenshot has been succesfully captured")
                continue

            elif "show me the screenshot" in command:
                try:
                    img = Image.open('D://Karl//Karl_2.0//' + name)
                    img.show(img)
                    speak("Here it is sir")
                    time.sleep(2)
                    continue

                except IOError:
                    speak("Sorry sir, I am unable to display the screenshot")
                    continue

            elif "hide all files" in command or "hide this folder" in command:
                os.system("attrib +h /s /d")
                speak("Sir, all the files in this folder are now hidden")
                continue
                

            elif "visible" in command or "make files visible" in command:
                os.system("attrib -h /s /d")
                speak("Sir, all the files in this folder are now visible to everyone. I hope you are taking this decision in your own peace")
                continue
           
            

            elif command in  ("goodbye","bye","offline","close","shutdown","abort","terminate","quit","power off"):
                speak("Alright sir, going offline. It was nice working with you")
                sys.exit()
                continue
            
            elif command:
                with open("pipe","rb") as f:
                    model = pickle.load(f)
                    z = model.predict(['command'])
                speak(z)
                print(z)
            
            
            """elif "launch microsoft edge" in command or "launch internet explorer" in command or "explorer" in command:
                speak(" opening internet explorer")
                subprocess.call("C:\Program Files\Internet Explorer\iexplore.exe")
                
                
            elif "launch notepad" in command or "notepad" in command :
                speak("Openeing Notepad")
                subprocess.call("C://Program Files (x86)//Notepad++//notepad++.exe")
                break
                
            elif"open windows excel" in command or "open excel" in command or "excel" in command :
                excel = win.gencache.EnsureDispatch('Excel.Application')
                excel.Visible = True
                #_ = input("press Enter to quit")
                #excel.Application.Quit()
                

            elif command in ("firefox","open firefox"):
                subprocess.call("C://Program Files//Mozilla Firefox//firefox.exe")
                
            elif command in ("launch wmplayer","wmplayer","windows media player"):
               subprocess.call("C:\Program Files\Windows Media Player\wmplayer.exe")
               
            elif "launch windows powerpoint" in command or "launch powerpoint" in command or "powerpoint" in command :
                excel = win.gencache.EnsureDispatch('Powerpoint.Application')
                excel.Visible = True
            
            elif "launch windows word" in command or "launch word" in command or "word" in command or "windows word" in command :
                excel = win.gencache.EnsureDispatch('Word.Application')
                excel.Visible = True"""
                










startExecution = MainThread()


class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.startTask)
        self.ui.pushButton_2.clicked.connect(self.close)

    def __del__(self):
        sys.stdout = sys.__stdout__

    # def run(self):
    #     self.TaskExection
    def startTask(self):
        self.ui.movie = QtGui.QMovie("Karl/utils/images/live_wallpaper.gif")
        self.ui.label.setMovie(self.ui.movie)
        self.ui.movie.start()
        self.ui.movie = QtGui.QMovie("Karl/utils/images/initiating.gif")
        self.ui.label_2.setMovie(self.ui.movie)
        self.ui.movie.start()
        timer = QTimer(self)
        timer.timeout.connect(self.showTime)
        timer.start(1000)
        startExecution.start()

    def showTime(self):
        current_time = QTime.currentTime()
        current_date = QDate.currentDate()
        label_time = current_time.toString('hh:mm:ss')
        label_date = current_date.toString(Qt.ISODate)
        self.ui.textBrowser.setText(label_date)
        self.ui.textBrowser_2.setText(label_time)


app = QApplication(sys.argv)
Karl = Main()
Karl.show()
exit(app.exec_())
