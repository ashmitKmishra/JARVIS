import pyttsx3 #pip install pyttsx3
import speech_recognition as sr #pip install speechRecognition
import datetime
import wikipedia #pip install wikipedia

import requests
from bs4 import BeautifulSoup
import webbrowser
import os
import smtplib
import pyjokes
import sys
import gtts
from win32com.client import Dispatch
r = sr.Recognizer()

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning Mr Ak!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon Mr Ak!")   

    else:
        speak("Good Evening Mr Ak!")  

    speak("Jarvisis now available for you, sir")       

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"Mr. Ashmit(Creator) said: {query}\n")

    except Exception as e:
        # print(e)    
        speak("Sir, please Say that again please...")  
        return "None"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('.....')
    server.sendmail('....', to, content)
    server.close()


if __name__ == "__main__":
    wishMe()
    while True:
    # if 1:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'do for me' in query:
            speak(".....from sending mail through voice command   to controlling my highness laptop,,   as per his wish,, of course except hacking nuclear codes, I can do everything for you sir. Just command me. ")

        elif 'going' in query:
            speak("AAAA..... It feels like I was in deep sleep,, but now I am glad that I got a chance to serve you, sir!")
            

        elif 'How can you serve' in query:
            speak("from sending mail through voice command   to controlling my highness laptop,,   as per his wish,, of course except hacking nuclear codes, I can do everything for you sir. Just command me. ")

        elif 'thank you' in query:
            speak('Its my duty to serve you sir')


        elif 'hello' in query:
            speak('I am listening sir, what can i do?')
            print('I am listening sir, what can i do?')



        
        elif 'happening in world' in query:
            url = 'https://www.bbc.com/news'
            response = requests.get(url)
  
            samachar = BeautifulSoup(response.text, 'html.parser')
            headlines = samachar.find('body').find_all('h3')
            for x in headlines:
                speak(x.text.strip())
                print(x.text.strip())
    
        elif 'open youtube' in query:
            speak("As youwish sir")
            webbrowser.open("https://www.youtube.com/")


        elif 'practice sat' in query:
            speak("yes sir")
            webbrowser.open("https://www.khanacademy.org/mission/sat/overview")

        # elif 'hospitals' in query:
        #     speak("Analyzing your location sir,,............ done!")
        #     webbrowser.open("https://www.justdial.com/Delhi/Private-Hospitals-in-Rohini/nct-10390288")
        #     speak("here it is, sir ,, please be safe sir!")

        # elif 'physician' in query:
        #     webbrowser.open("https://www.google.com/search?rlz=1C1CHBF_enIN1003IN1003&tbs=lf:1,lf_ui:2&tbm=lcl&sxsrf=ALiCzsa-s5RuBWXC4_SknCj7te4ItdZi9g:1660404328302&q=physician+near+me&rflfq=1&num=10&sa=X&ved=2ahUKEwiy0K-RkMT5AhW47DgGHaH7D-MQjGp6BAgLEAE&biw=1536&bih=714&dpr=1.25#rlfi=hd:;si:;mv:[[28.732946499999997,77.0965727],[28.710890299999996,77.0831408]];tbs:lrf:!1m4!1u3!2m2!3m1!1e1!1m4!1u2!2m2!2m1!1e1!2m1!1e2!2m1!1e3!2m4!1e17!4m2!17m1!1e2!3sIAE,lf:1,lf_ui:2")
        #     speak("Sir, I can provide the details of chemist shops too.")

        # elif 'dermatologist' in query:
        #     webbrowser("https://www.google.com/search?q=dermatologist+near+me&rlz=1C1CHBF_enIN1003IN1003&biw=1536&bih=714&tbm=lcl&sxsrf=ALiCzsaHIOMqjv7pdzYRvp0TTaU4TNnVfQ%3A1660404472238&ei=-ML3YsmMDrKq4-EP4vma8A4&oq=derma+near+me&gs_l=psy-ab.3.0.0i30i7k1l3j0i512k1j0i30i7k1l6.31166.31840.0.33567.5.5.0.0.0.0.170.606.0j4.4.0....0...1c.1.64.psy-ab..1.4.604...38j0i13k1j0i67k1.0.6JUz8X6rLV4#rlfi=hd:;si:;mv:[[28.7367489,77.1361032],[28.6909962,77.0670816]];tbs:lrf:!1m4!1u3!2m2!3m1!1e1!1m4!1u2!2m2!2m1!1e1!2m1!1e2!2m1!1e3!2m4!1e17!4m2!17m1!1e2!3sIAE,lf:1,lf_ui:2")


        elif 'open sat' in query:
            speak("yes sir")
            webbrowser.open("https://www.khanacademy.org/mission/sat/overview")

        elif 'open my classes' in query:
            speak("sure sir")
            webbrowser.open("https://study.physicswallah.live/tabs/tabs/batch-tab")

        elif 'open google' in query:
            speak("roger that")
            webbrowser.open("www.google.com")

        elif 'search something' in query:
            speak("roger that")
            webbrowser.open("www.google.com")

        elif 'open my website' in query:
            speak("Indeed Sir") 
            webbrowser.open("https://finoledgee.blogspot.com/")
        
        elif 'glimpse of my website' in query:
            speak("Indeed Sir") 
            webbrowser.open("https://finoledgee.blogspot.com/")

        elif 'my state' in query:
            speak("here it is sir") 
            webbrowser.open("https://my.msstate.edu/?check_logged_in=1")

        elif 'canvas' in query:
            speak("here it is sir") 
            webbrowser.open("https://canvas.msstate.edu/")

        elif 'college assignments' in query:
            speak("here it is sir") 
            webbrowser.open("https://canvas.msstate.edu/")

        elif 'housing portal' in query:
            webbrowser.open("https://msstate.starrezhousing.com/StarRezPortalX/622EBF66/1/1/Home-Welcome_to_myHousing?UrlToken=CDE3366A")   

        elif 'cheer up' in query:
            speak("Sir, i will try my best to cheer your mood, here is the joke::")
            speak(pyjokes.get_joke())

        elif 'check my mails' in query:
            speak("Let me do it for you sir")
            webbrowser.open("https://mail.google.com/mail/u/0/#inbox")



        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Sir, the time is {strTime}")
            print(f"Sir, the time is {strTime}")
            speak("sir by the way the time will always be in your favour only")

        elif 'send email' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "recipent@gmail.com"    
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry. I am not able to send this email")

        elif 'exit' in query:
            sys.exit(speak("INDEED SIR,  I will be always there to serve you sir"))              
        elif 'quit' in query:
            sys.exit(speak("INDEED SIR,  I will be always there to serve you"))

        