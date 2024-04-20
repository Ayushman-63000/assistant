import ctypes  # lock or shut down pc
import json
import os
import random  #
import re  # to find integer
import subprocess  # also to lock shutdown pc
import webbrowser
from datetime import date as D, datetime as DT
from functools import reduce  # to subtract
import time  # for sleep function
import keyboard  # to press space so that it can play song
import pygame
import requests  # to get news and articles
import speech_recognition as sr
import wikipedia
import winshell  # to empty recycle bin and lock / shut down pc
from ecapture import ecapture as ec  # to record audio and video and take photo
from requests.exceptions import ConnectionError, Timeout
from voice import SpeakText  # def function to speak my custom made

# Initialize the recognizer
r = sr.Recognizer()

# Function to set an alarm
def set_alarm(alarm_time):
    while True:
        now = DT.now()
        if now.strftime("%H:%M") == alarm_time:
            SpeakText("Time's up!")
            break
        time.sleep(1)

# API for Google News
api_key = '415b5d7afbb84a9f8dc2038d2121bfea'

def get_news(category=None):
    if category:
        url = f'https://newsapi.org/v2/top-headlines?country=in&category={category}&apiKey={api_key}'
    else:
        url = f'https://newsapi.org/v2/top-headlines?country=in&apiKey={api_key}'
    try:
        response = requests.get(url)
        data = response.json()
        for i in range(5):
            SpeakText(data['articles'][i]['title'])
    except requests.exceptions.RequestException as e:
        SpeakText("Error: Could not connect to the server")
    except ValueError as e:
        SpeakText("Error: Invalid response received from the server")

# Get current date and time
today = D.today()
now = DT.now()
current_time = now.strftime("%H:%M:%S")

SpeakText("Hello sir! Welcome, Today is {} and the time is {}".format(today, current_time))

# Loop infinitely for user to speak
while True:
    try:
        # use the microphone as source for input.
        with sr.Microphone() as source:
            r.adjust_for_ambient_noise(source, duration=0.2)
            audio = r.listen(source, timeout=3)
            M = r.recognize_google(audio).lower()
            print("Processing task " + M)

            # data file handling
            with open("data.txt", "a") as file:
                file.write("\n")
                file.write(M)

            # Websites
            if "open" in M:
                if "youtube" in M:
                    SpeakText("opening youtube")
                    webbrowser.open("www.youtube.com")
                elif "instagram" in M:
                    SpeakText("opening instagram")
                    webbrowser.open("https://www.instagram.com")
                elif "facebook" in M:
                    SpeakText("opening facebook")
                    webbrowser.open("https://www.facebook.com")
                elif "whatsapp" in M:
                    SpeakText("opening whatsapp")
                    webbrowser.open("https://web.whatsapp.com")
                elif "twitter" in M:
                    SpeakText("opening twitter")
                    webbrowser.open("https://www.twitter.com")
                elif "netflix" in M:
                    SpeakText("opening netflix")
                    webbrowser.open("https://www.netflix.com/in")
                elif "tik tok" in M:
                    SpeakText("opening tik tok")
                    webbrowser.open("https://www.tiktok.com/en")

                # Apps
                elif "excel" in M:
                    SpeakText("opening excel")
                    os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE")
                elif "ms word" in M:
                    SpeakText("opening word")
                    os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE")
                elif "powerpoint" in M:
                    SpeakText("opening powerpoint")
                    os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE")
                elif "chrome" in M:
                    SpeakText("opening chrome")
                    os.startfile("C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe")
                elif "edge" in M:
                    SpeakText("opening edge")
                    os.startfile("C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe")
                elif "obs" in M:
                    SpeakText("opening OBS")
                    os.startfile("C:\\Program Files\\obs-studio\\bin\\64bit\\obs64.exe")
                elif "task manager" in M:
                    SpeakText("opening task manager")
                    os.startfile("C:\\WINDOWS\\system32\\Taskmgr.exe")
                elif "my sql" in M:
                    SpeakText("opening my SQL")
                    os.startfile("C:\\Program Files\\MySQL\\MySQL Server 8.0\\bin\\mysql.exe")

            # Time
            elif "time" in M:
                SpeakText(current_time)

            # Date
            elif "date" in M:
                SpeakText(today)

            # Calculator
            numbers = re.findall(r'-?\d+', M)
            numbers = [float(x) for x in numbers]
            if "sum" in M or "+" in M:
                result = sum(numbers)
                c = "the sum is" if "sum" in M else "your result is"
                SpeakText(c)
                SpeakText(result)
                print(f"{c} {result}")

            elif "-" in M or "subtract" in M:
                result = reduce(lambda x, y: x - y, numbers)
                c = "the substraction is" if "-" in M else "your result is"
                SpeakText(c)
                SpeakText(result)
                print(f"{c} {result}")

            elif "multiply" in M or "multiplication" in M:
                result = reduce(lambda x, y: x * y, numbers)
                c = "the multiplication is" if "multiply" in M else "your result is"
                SpeakText(c)
                SpeakText(result)
                print(f"{c} {result}")

            elif "divide" in M or "/" in M:
                result = reduce(lambda x, y: x / y, numbers)
                c = "the division is" if "divide" in M else "your result is"
                SpeakText(c)
                SpeakText(result)
                print(f"{c} {result}")

            # Music
            elif "spotify" in M:
                webbrowser.open("https://open.spotify.com/collection/tracks")
                time.sleep(2)
                keyboard.press_and_release('space')
                time.sleep(6)
                keyboard.press_and_release('space')

            elif "music" in M:
                path = "C:\\Users\\RAJEEV\\Music\\Playlists"
                files = [f for f in os.listdir(path) if f.endswith(".mp3")]
                pygame.init()
                pygame.mixer.init()

                while True:
                    music = random.choice(files)
                    pygame.mixer.music.load(os.path.join(path, music))
                    pygame.mixer.music.play()

                    while pygame.mixer.music.get_busy():
                        if keyboard.is_pressed("p"):
                            pygame.mixer.music.pause()
                        elif keyboard.is_pressed("r"):
                            pygame.mixer.music.unpause()
                        elif keyboard.is_pressed("n"):
                            pygame.mixer.music.fadeout(2000)
                        elif keyboard.is_pressed("s"):
                            pygame.quit()
                            break

            # Search on web
            elif "search" in M:
                if "youtube" in M:
                    words = ["youtube", "search", "on"]

                    for word in words:
                        M = M.replace(word, "")
                    print(M)
                    SpeakText("searching on youtube")
                    M = ("https://www.youtube.com/results?search_query=" + M)
                    webbrowser.open(M)

                elif "wikipedia" in M:
                    try:
                        words = ["wikipedia", "search", "on", "about"]

                        for word in words:
                            M = M.replace(word, "")
                        print(M)
                        # Search for an article
                        result = wikipedia.search(M)

                        # Extract the summary of the first article
                        summary = wikipedia.summary(result[0])
                        print(summary)
                        SpeakText(summary)
                    except wikipedia.exceptions.PageError:
                        SpeakText("sorry I cannot find that")

                # Google search
                elif "google" in M:
                    words = ["google", "search", "on"]
                    for word in words:
                        M = M.replace(word, "")
                    print(M)
                    SpeakText("searching on google")
                    M = ("https://www.google.com/search?q=" + M)
                    webbrowser.open(M)
                elif "bing" in M:
                    words = ["bing", "search", "on"]
                    for word in words:
                        M = M.replace(word, "")
                    print(M)
                    SpeakText("searching on bing")
                    M = ("https://www.bing.com/search?q=" + M)
                    webbrowser.open(M)
                elif "article" in M:
                    words = ["article", "search", "about"]
                    for word in words:
                        M = M.replace(word, "")
                    print(M)
                    SpeakText("searching article on the web")
                    # Make the API call
                    url = f'https://newsapi.org/v2/everything?q={M}&apiKey={api_key}'
                    response = requests.get(url)
                    S = response.json()
                    # Parse the JSON response
                    data = json.loads(response.text)

                    # Extract the articles from the response
                    articles = data['articles']

                    # Iterate through the articles
                    for article in articles:
                        # Extract the content of each article
                        content = article['content']
                    print(content)
                    SpeakText(content)

            elif "what" in M or "who" in M or "when" in M:
                try:
                    words = ["who", "what", "when", "is", "was", "search", "on", "about"]
                    for word in words:
                        M = M.replace(word, "")
                    print(M)
                    # Search for an article
                    result = wikipedia.search(M)
                    # Extract the summary of the first article
                    summary = wikipedia.summary(result[0])
                    print(summary)
                except wikipedia.exceptions.PageError:
                    print("no result")

            elif "news" in M:
                if "business" in M:
                    get_news("business")
                elif "entertainment" in M:
                    get_news("entertainment")
                elif "general" in M:
                    get_news("general")
                elif "health" in M:
                    get_news("health")
                elif "science" in M:
                    get_news("science")
                elif "sports" in M:
                    get_news("sports")
                elif "tech" in M:
                    get_news("technology")
                else:
                    get_news()

            # Game not complete
            elif "bored" in M and "i" in M:
                SpeakText("here are some games for you that you might like")
                webbrowser.open("https://www.crazygames.com/")

            # Alarm
            elif "set alarm" in M:
                SpeakText("What time do you want to set the alarm for?")
                alarm_time = input("What time do you want to set the alarm for?")
                set_alarm(alarm_time)

            # Extras
            elif 'empty recycle bin' in M:
                winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
                SpeakText("Recycle Bin Recycled")
            elif 'lock window' in M:
                SpeakText("locking the device")
                ctypes.windll.user32.LockWorkStation()
            elif 'shutdown' in M:
                SpeakText("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')
            # Record
            elif "take" in M:
                if "photo" in M or "picture" in M:
                    ec.capture(0, "Camera ", "img.jpg")
                elif "video" in M:
                    ec.record("video.mp4")
                    time.sleep(5)
                    ec.stop()
                elif "audio" in M:
                    ec.record_audio("audio.mp3")
                    time.sleep(5)
                    ec.stop()
            elif "record" in M:
                if "video" in M:
                    ec.record("video.mp4")
                    time.sleep(5)
                    ec.stop()
                elif "audio" in M:
                    ec.record_audio("audio.mp3")
                    time.sleep(5)
                    ec.stop()
                elif "photo" in M or "picture" in M:
                    ec.capture(0, "Jarvis Camera ", "img.jpg")
            # Stop the program
            elif "stop" in M or "sleep" in M:
                SpeakText("bye sir")
                break

            SpeakText("complete")
    except sr.RequestError as e:
        print("Could not request results; {0}".format(e))

    except sr.UnknownValueError:
        print("unknown error occurred")
    except sr.WaitTimeoutError:
        print("timeout")
