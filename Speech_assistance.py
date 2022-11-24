import datetime
# Install speachRecognition
import speech_recognition as sr
import pyttsx3    # text to speach conversion
import wikipedia
import webbrowser
import os
import random



engine=pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
print(voices[1].id)
engine.setProperty('voice',voices[0].id)  #to set the voice . 0-male and 1-female

def speak(audio):
    engine.say(audio)
    engine.runAndWait()


    # from win32com.client import Dispatch
    # speak = Dispatch("SAPI.Spvoice")
    # speak.Speak(audio)

def wishme():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Goodmorning")
    elif hour>=12 and hour<=18:
        speak("Goodafternoon")
    else:
        speak("goodevening")
    speak("This is Sharwan how may I help you")


def takecommand():         # it will take microphone input from the user and return string
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening......")
        r.pause_threshold = 1  # it gives the time between the speaker saying and the result shown
        r.energy_threshold = 300  # on increasing we need to increase the voice level then only the system can listen
        audio = r.listen(source)
    try:
        print("Recognizing")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said : {query}\n")
    except Exception as e:
        # print(e)
        print("say that again please")
    return query
if __name__=='__main__':
    wishme()
    if 1:
        query = takecommand().lower()
        if 'wikipedia' in query:
            speak('Hritika is searching.....')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query)
            # results = wikipedia.summary(query, sentence=2)   sentence 2 mean it will read only 2 lines
            speak("Accoring to hritika")
            speak(results)

        elif 'open youtube' in query:
            webbrowser.open('www.youtube.com')
        elif 'open google' in query:
            webbrowser.open('www.google.com')
        elif 'open my instagram ' in query:
            webbrowser.open('www.instagram.com/sharwannnn/')
        elif 'open instagram' in query:
            webbrowser.open('www.instagram.com')
        elif 'open facebook' in query:
            webbrowser.open('www.facebook.com')
        elif 'play music' or 'play song' in query:
            # music = 'C:\\Users\\Sharwan Kumar\\Desktop\\music'  # \\ is used so that we can excape the chracter
            # songs = os.listdir(music)
            # playsongs = random.choice()
            # print(playsongs)
            music = 'C:\\Users\\Sharwan Kumar\\Desktop\\music'
            songs = os.listdir(music)
            print(songs)
            os.startfile(os.path.join(music,songs[0]))

        elif 'what is the time' in query:
            time = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"The time is {time}")
            # you can try to convert the 24 hour time to 12 hour time

        #  opening apps
        elif 'open vs code' in query:
            vscodepath = "C:\\Users\\Sharwan Kumar\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(vscodepath)