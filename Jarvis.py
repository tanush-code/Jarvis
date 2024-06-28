from tkinter import *
import pyperclip
from passwordgenerator import pwgenerator
import os
import speech_recognition as sr
from win32com.client import Dispatch
import webbrowser
import datetime
import time
import pandas as p
import pyautogui
import PyDictionary
#data and algorithm work
def password_genrator():
    try:
        speak("What should I name the app?")
        t2.insert(END,"What should I name the app?")
        r = sr.Recognizer()
        with sr.Microphone() as source:
            audio = r.listen(source)
            try:
                text = r.recognize_google(audio)
                password_name = format(text)

            except:
                speak("Sorry I can't recognize it")
                t2.insert("Sorry I can't recognize it")
        content = password_name
        password = pwgenerator.generate()
        Myfile = open('Main.txt', 'a')
        Myfile.write(f"{content} -- {password} \n")
        Myfile.close()
        t2.delete(1.0,END)
        speak("password is genrated press ctrl +v to paste")
        t2.insert("password is genrated press ctrl+V to paste")
        pyperclip.copy(password)
    except Exception as e:
        print(e)
        speak("password wasn't genrated")


def wishthemininsta():
    webbrowser.open('www.instagram.com/' + item['insta id'])
    time.sleep(8)
    pyautogui.moveTo(760, 170)
    pyautogui.click()
    time.sleep(3)
    pyautogui.write(item['wish'])
    pyautogui.press('enter')
def speak(str):
    speak = Dispatch(("sapi.SpVoice"))
    speak.Speak(str)
def say():
    t1.delete(1.0,END)
    t2.delete(1.0,END)
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            text = r.recognize_google(audio)
            t1.insert(END,format(text))
        except:
            t2.insert(END,"Sorry could not recognize what you said")

    questions = {"hello\n": "hi","who are you\n": "I am Jarvis",
                "what can you do\n": "I can do many things like opening website etc :-)"
        , "will you marry me\n": "nhaa i need robot", "who is your creator\n": "my creator is tanush",
                "what is your version\n": "it is 3.0 with gui", "do you know something else\n": "i don't know",
                "hi\n": "hello"}

    if t1.get(1.0,END) in questions:
        t2.insert(END,questions[t1.get(1.0,END)]+":\n")
        speak(questions[t1.get(1.0,END)]+":\n")
    elif "generate password" in t1.get(1.0,END):
        password_genrator()
    elif "open password" in t1.get(1.0,END):
        os.startfile("Main.txt")
    elif "open Google" in t1.get(1.0,END):
        speak("opening Google")
        t2.insert(END, "opening Google")
        webbrowser.open('www.Google.com')

    elif "open Instagram" in t1.get(1.0,END):
        speak("opening Instagram")
        t2.insert(END, "opening instagram")
        webbrowser.open('www.Instagram.com')

    elif "open YouTube" in t1.get(1.0,END):
        speak("opening Youtube")
        t2.insert(END, "opening youtube")
        webbrowser.open('www.youtube.com')

    elif "open SST class" in t1.get(1.0,END):
        speak("opening Sst class")
        t2.insert(END, "opening SST class")
        webbrowser.open('https://classroom.google.com/u/1/c/OTAwMDY1NzAwNDRa')

    elif "open science class" in t1.get(1.0,END):
        speak("opening science class")
        t2.insert(END, "opening science class")
        webbrowser.open('https://classroom.google.com/u/1/c/OTAwMTM4OTQ3NDBa')

    elif "open Maths class" in t1.get(1.0,END):
        speak("opening Maths class")
        t2.insert(END, "opening maths class")
        webbrowser.open('https://classroom.google.com/u/1/c/MTE2MTkyMDcyOTcz')

    elif "open English class" in t1.get(1.0,END):
        speak("opening English class")
        t2.insert(END, "opening English class")
        webbrowser.open('https://classroom.google.com/u/1/c/OTQ2MTAyMDgyMzla')

    elif "open Hindi class" in t1.get(1.0,END):
        speak("opening Hindi class")
        t2.insert(END, "opening hindi class")
        webbrowser.open('https://classroom.google.com/u/1/c/OTA3NzA5NzUxMTJa')

    elif "open French class" in t1.get(1.0,END):
        speak("opening French class")
        t2.insert(END, "opening french class")
        webbrowser.open('https://classroom.google.com/u/1/c/ODA4OTE5MzE1MTFa')

    elif "open computer class" in t1.get(1.0,END):
        speak("opening Coumputer class")
        t2.insert(END, "opening coumputer class")
        webbrowser.open('https://classroom.google.com/u/1/c/OTMyNTY5MjM0Njla')

    elif "open school mail" in t1.get(1.0,END):
        speak("opening gmail")
        t2.insert(END, "opening gmail")
        webbrowser.open('https://mail.google.com/mail/u/1/?pli=1')

    elif "check for birthday" in t1.get(1.0,END):
        speak("Checking for birthdays sir")
        t2.insert(END, "Checking for birthday")
        br = p.read_excel("Birthday.xlsx")
        today = datetime.datetime.now().strftime('%d-%m')
        # print(today)
        global item
        for index, item in br.iterrows():
            # print(index, item['date'])
            birthday_date = item['date'].strftime('%d-%m')
            name = item['name']
            print(item['name'])
        if str(birthday_date) in today:
            t2.delete(1.0,END)
            print(item['name'])
            speak("Today is the birthday of" + item['name'] + 'wishing her in 5 seconds')
            t2.insert(END, "Today is the birthday of" + item['name'] + "wishing her in 5 seconds")
            time.sleep(5)
            wishthemininsta()
        else:
            speak("no one's birthday is today Thank You")
            t2.insert(END, " no ones birthday today")
            print(today)
            print(birthday_date)
            print(type(birthday_date))
            print(type(today))
    elif "antonym" in t1.get(1.0,END):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            speak("please say the antoym sir")
            t2.insert(END,"Pls say the antoynm sir")
            audio = r.listen(source)
            try:
                text = r.recognize_google(audio)
                antonym_word = format(text)

            except:
                speak("Sorry I can't recognize it")
                t2.insert(END, "Sorry i can't recognize it")

        antonym = PyDictionary.PyDictionary.antonym(antonym_word)

        speak(f"the synonym of the word is {antonym}")

    elif "the meaning" in t1.get(1.0,END):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            speak("please say the word sir")
            t2.insert(END, "please say the words sir")
            audio = r.listen(source)
            try:
                text = r.recognize_google(audio)
                theword = format(text)
                word = PyDictionary.PyDictionary.meaning(theword)
                speak("The meaning of the word is" + str(word))
                t2.insert(END, "The meaning of the word is " + str(word))

            except:
                speak("Sorry I can't recognize it")
                t2.insert(END, "sorry i can't recognize it ")
    else:
        speak("sorry I can't do that")
        t2.insert(END, 'Sorry i cant do that')



#the tkinter work
root = Tk()
root.maxsize(width=400,height=400)
root.title("jarvis")
l1 = Label(root,text="You said:",font=("Bold",20)).place(x=0,y=0)
t1 = Text(root)
t1.place(x=1,y=40,height=100,width=400)
global t1value
t1value = t1.get(1.0,END)
b1 = Button(root,text="speak",command=say).place(x=1,y=150)
l2 = Label(root,text="Answer:",font=("Bold",20)).place(x=1,y=180)
t2 = Text(root)
t2.place(x=1,y=220,height=100,width=400)
root.mainloop()
