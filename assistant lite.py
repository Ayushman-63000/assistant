import os
import random
import re
import time
from datetime import datetime, date
from functools import reduce
from multiprocessing.connection import wait

import pyttsx3
import pygame
import pyaudio
import speech_recognition as sr
import webbrowser
import wikipediaapi
import keyboard

# Initialize the recognizer
recognizer = sr.Recognizer()

# Function to convert text to speech
def speak_text(text):
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[3].id)
    engine.say(text)
    engine.runAndWait()

def get_current_time():
    now = datetime.now()
    return now.strftime("%H:%M:%S")

def get_today_date():
    return date.today()

# Welcome message
today = get_today_date()
current_time = get_current_time()
welcome_message = f"Hello sir! Welcome. Today is {today} and the time is {current_time}"
speak_text(welcome_message)

# Wikipedia setup
wiki = wikipediaapi.Wikipedia(language='en', extract_format=wikipediaapi.ExtractFormat.WIKI)

# Loop infinitely for user to speak
while True:
    try:
        with sr.Microphone() as source:
            recognizer.adjust_for_ambient_noise(source, duration=0.2)
            audio = recognizer.listen(source)

        user_input = recognizer.recognize_google(audio).lower()
        print("Processing task: ", user_input)

        # Process user input
        if "youtube" in user_input:
            webbrowser.open("https://www.youtube.com")
        elif "instagram" in user_input:
            webbrowser.open("https://www.instagram.com")
        elif "facebook" in user_input:
            webbrowser.open("https://www.facebook.com")
        elif "whatsapp" in user_input:
            webbrowser.open("https://web.whatsapp.com")
        
        elif "time" in user_input:
            speak_text(get_current_time())
        elif "date" in user_input:
            speak_text(get_today_date())

        elif "excel" in user_input:
            os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE")
        elif "ms word" in user_input:
            os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE")
        elif "powerpoint" in user_input:
            os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE")
        elif "chrome" in user_input:
            os.startfile("C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe")
        elif "edge" in user_input:
            os.startfile("C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe")
        elif "OBS" in user_input:
            os.startfile("C:\\Program Files\\obs-studio\\bin\\64bit\\obs64.exe")
        elif "task manager" in user_input:
            os.startfile("C:\\WINDOWS\\system32\\Taskmgr.exe")
        elif "my sql" in user_input:
            os.startfile("C:\\Program Files\\MySQL\\MySQL Server 8.0\\bin\\mysql.exe")

        # Calculations
        if "sum" in user_input:
            numbers = re.findall(r'\d+', user_input)
            numbers = [int(x) for x in numbers]
            result = sum(numbers)
            speak_text(f"The sum is {result}")
        elif "subtract" in user_input:
            numbers = re.findall(r'\d+', user_input)
            numbers = [int(x) for x in numbers]
            result = reduce(lambda x, y: y - x, numbers)
            speak_text(f"The subtraction is {result}")
        # Add other calculations here

        # Music playing
        if "music" in user_input:
            path = "C:\\Users\\RAJEEV\\Music\\Playlists"
            files = [f for f in os.listdir(path) if f.endswith(".mp3")]
            music = random.choice(files)
            pygame.init()
            while True:
                pygame.mixer.music.load(os.path.join(path, music))
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy():
                    for event in pygame.event.get():
                        if event.type == pygame.KEYDOWN:
                            if event.unicode == 's':
                                pygame.mixer.music.fadeout(3000)
                                break
                break

        # Stop commands
        if user_input in ["exit", "sleep", "shut down", "go to sleep", "stop xper"]:
            speak_text("Goodbye sir")
            break

        # Wikipedia search
        if "wikipedia" in user_input:
            query = re.search("wikipedia (.*)", user_input)
            if query:
                page = wiki.page(query.group(1))
                if page.exists():
                    speak_text("According to Wikipedia, " + page.summary[0:60])
                else:
                    speak_text("Sorry, I couldn't find any information about " + query.group(1))

        # Google search
        if "google" in user_input and "search" in user_input:
            webbrowser.open("https://www.google.com/search?q=" + user_input)

        speak_text("Task completed")

    except sr.RequestError as e:
        print("Could not request results: ", e)
    except sr.UnknownValueError:
        print("Unknown error occurred")
