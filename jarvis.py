import pyttsx3 #pip install pyttsx3
import datetime
import pyaudio
import speech_recognition as sr # pip install speechRecognition
import wikipedia # pip install wikipedia
import webbrowser
import os
import smtplib
import json

engine = pyttsx3.init('sapi5') # taking voice form window
voices = engine.getProperty('voices')
# print(voices[].id)
engine.setProperty('voice', voices[1].id)



def speak(audio):
    engine.say(audio) # speak function 
    engine.runAndWait()

def wishMe():
    # jarvis will greet according to date and time
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        # he will wish you
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")

    else:
        speak("Good Evening")        

    speak("I am jarvis Sir. Please tell me how may I help you")    

def takecommand():
    # It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        # he will wait till you complete & take a pause
        audio = r.listen(source) # speech reco... module

    try:
        print("Recognizing...")    
        # if any chance of error 
        query = r.recognize_google(audio, language='en-in')
        print(f"User said:, {query}\n")

    except Exception as e:
        #print(e)    

        print("Say that again please....")
        return"None" # returning string 
    return query

def sendEmail(do, content):
    server = smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.login('yourgmail.com', 'password')
    server.sendmail('yourgmail.com', to_addrs , msg)
    # from to msg 
    server.close()

# to read the newspaper
def speak(str):
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.SpVoice")    
    speak.Speak(str)

if __name__ == "__main__":
    #speak("pranjuli is a good girl")   
    wishMe()

    #while True:
    if 1:
        # it will continue listen you
        query = takecommand().lower()

         # logic for executing task based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
                webbrowser.open("youtube.com")

        elif 'open google' in query:
                webbrowser.open("google.com")

        elif 'open facebook' in query:
                webbrowser.open("facebook.com")                

        elif 'open stackoverflow' in query:
                webbrowser.open("stackoverflow.com") 

        elif'play music' in query:
            music_dir = 'E:\\music'
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir, songs[2]))

        elif 'the time' in query:
            strTime= datetime.datetime.now().strftime("%H:%M:%S")    
            # it tell will tell you about current time
            speak(f"Sweetheart, the time is {strTime}")

        elif 'open code' in query:
            codePath = "C:\\Users\\PlusOne\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)   


        #elif 'play music in youtube' in query:
        #    mPath = "https://www.youtube.com//watch?v=VxDm4brYPg8"  
        #    os.startfile(mPath)  
        
        
        elif 'email to pranjuli' in query:
            try:
                speak("what should I say")
                content = takecommand()
                to = "your@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent!")

            except Exception as e:
                print(e)
                speak("sorry my friend pranjuli. I am not able to send mail.")
            
        elif "shutdown" in query:
            check = input("Want to shutdown your computer ? (y/n): ");
            if check == 'y' or check == "Y":
                os.system("shutdown /s /t 1");

        elif "restart" in query:
            check = input("Want to restart your computer ? (y/n): ");
            if check == 'y' or check == "Y":
                os.system("shutdown /r /t 1");

        elif " repeat after me " in query:
            query=query. replace("repeat after me "," ")
            Speak (query)       

