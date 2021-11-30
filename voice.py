import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import ctypes
import time
import requests
import shutil
from clint.textui import progress
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from tkinter import *

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishme():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("hello,Good Morning!")

    elif hour >= 12 and hour < 18:
        speak("hello,Good Afternoon!")

    else:
        speak("hi,Good Evening !")

    assname = ("avatar")
    speak("I am your Assistant")
    speak(assname)


def usrname():
    speak("What should i do call you sir ")
    uname = takeCommand()
    speak("Welcome Mister")
    speak(uname)
    columns = shutil.get_terminal_size().columns

    print("##################################################################".center(columns))
    print("Welcome Mr.", uname)
    print("##################################################################".center(columns))

    speak("How can i help you, Sir")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:

        print("listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

        try:
            print("recognizing....")
            query = r.recognize_google(audio, language='en-in')
            print(f"User said :{query}\n")


        except Exception as e:
            print(e)
            print("Unable to recognize your voice.")
            return "None"

        return query


def sendEmail(to, content):
    server = smptlib.SMTP('smpt.gmail.com', 587)
    server.ehlo()
    server.sterttls()

    # Enable low security in gmail
    server.login('your email id ', 'your email password')
    server.sendmail('your email id', to, content)
    server.close()


if __name__ == '__main__':
    clear = lambda: os.system('cls')

    # this function will clean any
    # command before execution of this python file
    clear()
    wishme()
    usrname()

    while 'true':

        query = takeCommand().lower()

        # all commands said by user will be
        # stored here in ;query' and will be
        # converted to lower case for easily
        # recognition of command
        if 'wikipedia' in query:
            speak('Searching wikipedia....')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=3)
            speak("according to wikipedia")
            print(results)
            speak(results)

        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")

        elif 'open youtube' in query:
            speak("here you go to Youtube/n")
            webbrowser.open("youtube.com")

        elif 'play music' in query or 'play song' in query:
            speak("here you go with your music ")
            # music_dir = "g:\\song"
            music_dir = "F:\\Music"
            songs = os.listdir(music_dir)
            print(songs)
            random = os.startfile(os.path.join(music_dir, songs[1]))

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif 'email to pannaga' in query:
            try:
                speak("What should I say")
                content = takeCommand()
                to = "Reciever email address"
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("Iam not able to send this mail")

        elif 'send a mail' in query:
            try:
                speak("What should I say ?")
                content = takeCommand()
                speak("whom should I say")
                to = input()
                sendEmail(to, content)
                speak("mail has been sent !")
            except Exception as e:
                print(e)
                speak("I am unable to send this mail")

        elif 'how are you' in query:
            speak("Iamfine,Thank you")
            speak("how about you, Sir")

        elif 'fine' in query or "good" in query:
            speak("It's happy to know, that you are fine")

        elif 'change name' in query:
            speak("sorry to say this, you have no option to change my name")

        elif 'Whats your name' in query or 'what is your name' in query:
            speak("My friends call me")
            assname = "avatar"
            speak(assname)
            print("My friends call me", assname)

        elif 'exit' in query:
            speak("Thanks for lending your time")
            exit()

        elif 'who created you' in query:
            speak("i was created by Pannaga, the rest is beautiful history")

        elif 'will you marry me' in query:
            speak("if i was not AI, I shall be ready for it")

        elif 'open presentation' in query:
            speak("opening Power Point presentation")
            power = r"C:\\Users\\GAURAV\\Desktop\\Minor Project\\Presentation\\Voice Assistant.pptx"
            os.startfile(power)

        elif 'what is love' in query:
            speak("It is 7th sense that destroy all other senses")

        elif "Good Morning" in query:
            speak("A warm" + query)
            speak("How are you Mister")

        elif 'open youtube' in query:
            speak("here you go to you tube enjoy")
            webbrowser.open("youtube.com/n")


        elif 'shutdown system' in query:
            speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')



        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "what can you do" in query:
             speak("i can do several tasks")
             speak("try saying open google")

        



















