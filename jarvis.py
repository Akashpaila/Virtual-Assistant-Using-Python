import pyttsx3 as pyt
import speech_recognition as sr
import datetime as dt
import os
import subprocess
import cv2
import threading
import webbrowser
import requests
from requests import get
import wikipedia
import pywhatkit as kit
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import sys
import time
import pyjokes
import pyautogui
import geocoder
# import geopy.geocoders
from geopy.geocoders import Nominatim


engine = pyt.init('sapi5')

voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
rate = engine.getProperty('rate')
engine.setProperty('rate', rate - 10)

####################################  Text to Speech ###################################
def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()

####################################  Basic Greeting Message ###########################
def greet():
    hour = dt.datetime.now().hour
    # tt = time.strftime("%I:%M:%p")

    if 0 <= hour < 12:
        speak("Good morning Akash")
    elif 12 <= hour < 16:
        speak("Good afternoon Akash")
    elif 16 <= hour < 20:
        speak("Good evening Akash")
    else:
        speak("Good night Akash")
    speak("How was your day? This is your Virtual Assistant Aku. How can I help you?")

####################################  Taking Input from User ###########################
def take_input():
    rec = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening ...")
        rec.adjust_for_ambient_noise(source)
        audio = None
        try:
            audio = rec.listen(source, timeout=5, phrase_time_limit=10)
        except sr.WaitTimeoutError:
            print("Listening timed out while waiting for phrase to start.")
            return "none"

    try:
        print("Recognizing ...")
        query = rec.recognize_google(audio, language='en-in')
        print(f"User said: {query}")
    except sr.UnknownValueError:
        speak("I didn't understand what you said. Please say that again.")
        return "none"
    except sr.RequestError:
        speak("Could not request results from Google Speech Recognition service.")
        return "none"

    return query.lower()

####################################  Open Applications ################################
def open_application(query):
    if 'notepad' in query:
        speak("Opening Notepad")
        subprocess.Popen(['notepad'])
    elif 'command prompt' in query or 'cmd' in query:
        speak("Opening Command Prompt")
        subprocess.Popen(['cmd', '/c', 'start'])
    elif 'calculator' in query:
        speak("Opening Calculator")
        subprocess.Popen(['calc'])
    elif 'chrome' in query or 'browser' in query:
        speak("Opening Google Chrome")
        subprocess.Popen(['start', 'chrome'], shell=True)
    elif 'camera' in query:
        speak("Opening Camera")
        open_camera()
    elif 'play music' in query:
        speak("Sure, which platform do you suggest?")
        platform_query = take_input()
        play_music(platform_query)
    elif 'open google' in query:
        speak("Opening Google. What should I search?")
        search_query = take_input()
        search_google(search_query)
    elif 'ip address' in query:
        speak("Fetching your IP address")
        ip = get('http://api.ipify.org').text
        speak(f"Your IP address is {ip}")
    elif 'wikipedia' in query:
        speak("Searching Wikipedia...")
        query = query.replace('wikipedia',"")
        results = wikipedia.summary(query, sentences=2)
        speak("According to Wikipedia")
        speak(results)
        print(results)
    elif 'youtube' in query:
        webbrowser.open('http://www.youtube.com')
    elif 'facebook' in query:
        webbrowser.open('http://www.facebook.com')
    elif 'stack overflow' in query:
        webbrowser.open('http://www.stackoverflow.com')
    elif 'message' in query or 'whatsapp' in query:
        send_whatsapp_message(query)
    elif 'email' in query or 'gmail' in query:
        send_email()
    elif 'terminate' in query or 'dismiss' in query:
        terminate()
    elif 'set alarm' in query :
        set_alarm()
    elif 'joke' in query or 'tell me some jokes' in query:
        jokes()
    elif 'shut down the system' in query:
        os.system('shutdown /s /t 5')
    elif 'restart the system' in query:
        os.system('shutdown /r /t 5')
    elif 'sleep' in query:
        os.system('rundll32.exe powrprof.dll, SetSuspendState 0,1,0')
    elif 'switch tab' in query or 'switch window' in query:
        switch()
    elif 'news' in query :
        news()
    elif 'location' in query:
        if 'current location' in query or 'where am i' in query:
         get_location()
        else:
            search_location(query)

    else:
        speak(f"Searching Google for {query}")
        search_google(query)

####################################  Open Camera #####################################
def open_camera():
    def camera_thread():
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            speak("Could not open camera")
            return

        while True:
            ret, frame = cap.read()
            if not ret:
                break
            cv2.imshow('Camera', frame)

            # Press 'q' to exit the camera view
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

        cap.release()
        cv2.destroyAllWindows()

    camera_thread = threading.Thread(target=camera_thread)
    camera_thread.start()

####################################  Play Music ######################################
def play_music(platform_query):
    if 'spotify' in platform_query:
        speak("Opening Spotify")
        webbrowser.open_new_tab('https://open.spotify.com/')
    elif 'wink' in platform_query:
        speak("Opening Wynk")
        webbrowser.open_new_tab('https://wynk.in/')
    elif 'my playlist' in platform_query:
        speak("Opening local playlist")
    elif 'amazon music' in platform_query:
        speak("Opening Amazon Music")
        webbrowser.open_new_tab('https://music.amazon.com/')
    elif 'jio saavn' in platform_query or 'jiosaavn' in platform_query:
        speak("Opening JioSaavn")
        webbrowser.open_new_tab('https://www.jiosaavn.com/')
    else:
        speak("Sorry, I can't play music from that platform")

####################################  Search in Google ################################
def search_google(search_query):
    url = f"https://www.google.com/search?q={search_query}"
    webbrowser.open_new_tab(url)

####################################  Send WhatsApp Message ###########################
def send_whatsapp_message(query):
    speak("Sending WhatsApp message")
    
    # Extract the recipient's name from the query
    recipient_name = query.split("send message to ")[-1].strip()
    
    # Prompt user for the message content
    speak(f"What message would you like to send to {recipient_name}?")
    message = take_input()

    # Open WhatsApp Web for sending the message
    webbrowser.open("https://web.whatsapp.com/send?phone=&text=" + message)
    speak("Please select the contact manually and the message will be sent automatically.")

####################################  Send Email ######################################
def send_email():
    speak("Please type the email address of the recipient.")
    recipient_email = input("Enter the recipient email address: ")

    speak(f"You entered {recipient_email}. Is that correct? Please type yes or no.")
    confirmation = input("Is that correct: ")

    if 'yes' in confirmation:
        speak("What is the subject of the email?")
        subject = take_input()

        speak("What should I say in the email?")
        content = take_input()

        try:
            send_email_smtp(recipient_email, subject, content)
            speak("Email has been sent successfully.")
        except Exception as e:
            print(e)
            speak("I am not able to send the email at the moment. Please try again later.")
    else:
        speak("Okay, let's try again.")
        send_email()

def send_email_smtp(to, subject, body):
    from_email = "cs21b1017@iiitr.ac.in"
    from_password = "Akash597143618"

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_email, from_password)
    text = msg.as_string()
    server.sendmail(from_email, to, text)
    server.quit()

################################ Set Alarm #############################################
# def set_alarm():
#     nn = int(dt.datetime.now().hour)
#     if nn==

################################ Jokes Telling ##########################################
def jokes():
    joke = pyjokes.get_joke()
    speak(joke)

################################ Tabs Switching ##########################################
def switch():
    pyautogui.keyDown("alt")
    pyautogui.press("tab")
    time.sleep(1)
    pyautogui.keyUp("alt")

################################ News telling ##########################################
def news():
    speak("Please wait, fetching the latest news...")
    url = ('https://newsapi.org/v2/top-headlines?'
       'sources=bbc-news&'
       'apiKey=98b32b7c5911455ab1ebf780d8c86a9f')
    main_page = requests.get(url).json()
    articles = main_page["articles"]
    head = []
    day = ["first","second","third","fourth","fifth","sixth","seventh","eighth","ninth","tenth"]
    for ar in articles:
        head.append(ar["title"])
    for i in range (len(day)):
        speak(f"today's {day[i]} news is : {head[i]}")

####################################  Get Location ################################
def get_location():
    speak("Fetching your current location")
    g = geocoder.ip('me')
    latlng = g.latlng
    location = geocoder.osm(latlng, method='reverse')
    if location:
        address = location.address
        speak(f"You are currently at {address}")
    else:
        speak("Sorry, I couldn't determine your location.")

####################################  Search Location ################################
def search_location(query):
    speak(f"Searching for {query}")
    geolocator = Nominatim(user_agent="AkuVirtualAssistant")
    location = geolocator.geocode(query)
    if location:
        address = location.address
        latitude = location.latitude
        longitude = location.longitude
        speak(f"Here is the information for {query}:")
        speak(f"Address: {address}")
        speak(f"Latitude: {latitude}")
        speak(f"Longitude: {longitude}")
    else:
        speak(f"Sorry, I couldn't find any information for {query}.")

################################ Terminate Virtual Assistant ############################
def terminate():
    speak("Thanks for using me, Have a good day")
    sys.exit()



if __name__ == '__main__':
    greet()
    while True:
        query = take_input()
        if query == "none":
            continue
        open_application(query)
    

    
