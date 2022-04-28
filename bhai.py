# Imports
import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import smtplib
import pyautogui
import psutil
import wmi
import requests
import pyjokes
import speedtest
import urllib
from bs4 import BeautifulSoup as bs
import fbchat
import cv2
from word2number import w2n
import time



# Initializations
engine = pyttsx3.init("sapi5")
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[0].id)

MUSIC_DIR = "C:\\Users\\ahamm\\Music"
MOVIE_DIR = "D:\\Movies"
CODE_PATH = "C:\\Users\\ahamm\\AppData\\Local\\Programs\\Microsoft VS Code"
OFFICE_PATH = "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16"
WORD_PATH = os.path.join(OFFICE_PATH, "WINWORD.EXE")
PP_PATH = os.path.join(OFFICE_PATH, "POWERPNT.EXE")
EXCEL_PATH = os.path.join(OFFICE_PATH, "EXCEL.EXE")
EDGE_PATH = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
MY_EMAIL = "ahammadshawki8@gmail.com"
MY_PASSWORD = os.environ.get("EMAIL_PASS")
EMAIL_ADDRESS = {"shawki":"ahammadshawki8@gmail.com", "organisation":"team.as8.org@gmail.com"}
LOCATION = "DHAKA"
WEATHER_API_KEY = os.environ.get("WEATHER_API")
FB_USERNAME = "ahammadshawki8"
FB_PASS = os.environ.get("FB_PASS")

MATH_SYMBOLS_MAPPING = {
    "equal": "=",
    "plus": "+",
    "minus": "-",
    "into": "*",
    "divide": "/",
    "modulo": "mod",
    "power": "**",
}

STUFF_BHAI_CAN_DO = "I can do a lot of stuffs. I can continue a basic conversation with you. " + \
                    "My other capabilities include: " + \
                    "Providing system report, weather report, and network report. " + \
                    "Telling current time. " + \
                    "Perform Dynamic news searching. " + \
                    "Opening Web Browser, Visual Studio Code, Calculator, and Office Softwares like Word, Powerpoint, Excel. " + \
                    "Searching on Google, Youtube and Wikipedia. " + \
                    "Opening Facebook, LinkedIn and Gmail. " + \
                    "Capturing your selfie. " + \
                    "Telling you jokes. " + \
                    "Sending Messages to your Facebook Friends. " + \
                    "Sending Emails. " + \
                    "Play Music. " + \
                    "Pause or Resume the current track or move to next/previous track. " + \
                    "Start a movie for you. " + \
                    "Increase or decrease system output volume. " + \
                    "Do basic calculation. " + \
                    "And finally sleep when I am done."

ANS_QUES_MAP = {
        "Hello Bhai": ["hello", "hi", "hey", "hello there"],
        "I am Fine Bhai. What about you?": ["how are you", "how are you doing"],
        "It is nice to hear bhai.": ["fine", "I am also fine", "I am well", "fine thank you", "i am doing well", "pretty good"],
        "I am your bhai": ["who are you", "what is your identity", "what is your name"],
        "Ahammad Shawki has created me": ["who made you", "who created you", "who is your creator"],
        "Welcome": ["thanks", "thank you"],
        "Should I tell you a joke or play any music for you?": ["my mood is off", "i am not feeling great"],
        STUFF_BHAI_CAN_DO : ["tell me what can you do", "what can you do", "which tasks you can perform"],
    }



# Functions
def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")

    speak("How can I help you?")
     

def takeCommand(lang = "en-in"):
    rec = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening....")
        rec.energy_threshold = 500
        audio = rec.listen(source)

    try:
        print("Recognizing....")
        query = rec.recognize_google(audio, language=lang)
        print(f"User said: {query}")
        return query
    except Exception as e:
        print(e)
        print("Say that again please...")
        return "None"


def sendEmail(to, content):
    # enable less secure apps
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.ehlo()
    server.starttls()
    server.login(MY_EMAIL, MY_PASSWORD)
    server.sendmail(from_addr=MY_EMAIL, to_addrs=to, msg=content)
    server.close()


def systemReport():
    battery = psutil.sensors_battery()
    percent = battery.percent
    plugged = "plugged in" if battery.power_plugged else "not plugged in"
    
    inst = wmi.WMI()
    num_process = len(inst.Win32_Process())
    return f"All systems are at 100 percent. Battery percentage: {percent} percent. Battery state: {plugged}. {num_process} processes are currently running."
    

def weatherReport(CITY, API_KEY):
    BASE_URL = "https://api.openweathermap.org/data/2.5/weather?"
    URL = BASE_URL + "q=" + CITY + "&appid=" + API_KEY
    
    response = requests.get(URL)
    if response.status_code == 200:
        data = response.json()
        main = data["main"]
        temperature = main["temp"]
        humidity = main["humidity"]
        pressure = main["pressure"]
        report = data["weather"]
        speak(f"Weather Report of Dhaka")
        speak(f"The Temperature is: {temperature}")
        speak(f"The Humidity is: {humidity}")
        speak(f"The Pressure is: {pressure}")
        speak(f"The Weather Condition is: {report[0]['description']}")
    else:
        print("Error in the HTTP request")


def networkReport():
    st = speedtest.Speedtest()
    try:
        speak("Creating Network Report...")
        server_names = []
        st.get_servers(server_names)

        downlink_bps = st.download()
        uplink_bps = st.upload()
        ping = st.results.ping
        up_mbps = uplink_bps / 1000000
        down_mbps = downlink_bps / 1000000

        speak("Speedtest results:\n"
            "The ping is: %s ms \n"
            "The upling is: %0.2f Mbps \n"
            "The downling is: %0.2f Mbps" % (ping, up_mbps, down_mbps)
            )

    except Exception as e:
        speak("Sorry, I coudn't run a speedtest")


def newsReport():
    # Can't find rss for Bangladeshi Newspapers
    try:
        news_url = "https://news.google.com/news/rss"
        client = urllib.request.urlopen(news_url)
        xml_page = client.read()
        client.close()
        soup = bs(xml_page, "xml")
        news_list = soup.findAll("item")
        response = ""
        for news in news_list[:5]:
            data = news.title.text.encode("utf-8")
            response += data.decode()
        speak(response)
    except Exception as e:
        print(e)
        speak("I can't find about daily news..")


def sendMessage():
    client = fbchat.Client(FB_USERNAME, FB_PASS)
    friend_name = input("Enter your friends username: ")
    friends = client.searchForUsers(friend_name)
    friend = friends[0]
    print("User's ID: {}".format(friend.uid))
    print("User's name: {}".format(friend.name))
    print("User's profile picture URL: {}".format(friend.photo))
    print("User's main URL: {}".format(friend.url))
    speak("What will be your message?")
    msg = takeCommand("bn")
    sent = client.sendMessage(msg, thread_id=friend.uid)
    if sent:
        print("Message sent successfully!")


def googleSearch(query):
    if query.startswith("google search "):
        query = query.replace("google search ", "")
    keyword_list = query.split(" ")
    new_query = "+".join(keyword_list)
    webbrowser.open(f"https://www.google.com/search?q={new_query}")


def youtubeSearch(query):
    if query.startswith("youtube search "):
        query = query.replace("youtube search ", "")
    if query.startswith("search on youtube "):
        query = query.replace("search on youtube ", "")
    keyword_list = query.split(" ")
    new_query = "+".join(keyword_list)
    webbrowser.open(f"https://www.youtube.com/results?search_query={new_query}")


def replace_words_with_numbers(transcript):
        transcript_with_numbers = ""
        for word in transcript.split():
            try:
                number = w2n.word_to_num(word)
                transcript_with_numbers += " " + str(number)
            except ValueError as e:
                transcript_with_numbers += " " + word
        return transcript_with_numbers


def clear_transcript(transcript):
        """
        Keep in transcript only numbers and operators
        """
        cleaned_transcript = ""
        for word in transcript.split():
            if word.isdigit() or word in MATH_SYMBOLS_MAPPING.values():
                # Add numbers
                cleaned_transcript += word
            else:
                # Add operators
                cleaned_transcript += MATH_SYMBOLS_MAPPING.get(word, "")
        return cleaned_transcript


def do_calculations(voice_transcript, **kwargs):
        transcript_with_numbers = replace_words_with_numbers(voice_transcript)
        math_equation = clear_transcript(transcript_with_numbers)
        try:
            result = str(eval(math_equation))
            speak(f"The asnwer of your math equation is {result}.")
        except Exception as e:
            print('Failed to eval the equation --> {0} with error message {1}'.format(math_equation, e))



# Main loop
if __name__ == "__main__":
    wishMe()
    while True:
        query = takeCommand().lower()

        for key, value in ANS_QUES_MAP.items():
            if query in value:
                speak(key)
                break 
        
        if 'system report' in query:
            speak(systemReport())

        elif "weather" in query:
            weatherReport(LOCATION, WEATHER_API_KEY)

        elif "network report" in query or "internet speed" in query or "speed test" in query:
            networkReport()

        elif "news" in query:
            newsReport()

        elif "joke" in query:
            speak(pyjokes.get_joke())
        
        elif 'wikipedia' in query:
            speak("Searching Wikipedia...")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)
        
        elif "what is" in query or "google" in query:
            googleSearch(query)

        elif "youtube search" in query or "search on youtube" in query:
            youtubeSearch(query)

        elif "open youtube" in query:
            webbrowser.open("youtube.com")

        elif "open google" in query:
            webbrowser.open("google.com")

        elif "open facebook" in query:
            webbrowser.open("facebook.com")
        
        elif "open linkedin" in query:
            webbrowser.open("linkedin.com")

        elif "open gmail" in query:
            webbrowser.open("gmail.com")

        elif "play music" in query:
            songs = os.listdir(MUSIC_DIR)
            os.startfile(os.path.join(MUSIC_DIR, songs[0]))

        elif "volume up" in query:
            pyautogui.press("volumeup", presses = 10)

        elif "volume down" in query:
            pyautogui.press("volumedown", presses = 10)

        elif "mute" in query:
            pyautogui.press("volumemute")

        elif "stop" in query or "resume" in query:
            pyautogui.press("playpause")

        elif "next" in query:
            pyautogui.press("nexttrack")

        elif "previous" in query:
            pyautogui.press("prevtrack")

        elif "close" in query:
            pyautogui.keyDown('alt')
            pyautogui.press("f4")

        elif "movie" in query:
            print(os.listdir(MOVIE_DIR))
            speak("Which movie do you want to watch?")
            movie_name = takeCommand().lower()
            while movie_name == None:
                movie_name = takeCommand().lower()
            try:
                for movie in os.listdir(MOVIE_DIR):
                    if movie_name in movie:
                        os.startfile(os.path.join(MOVIE_DIR, movie))
            except Exception as e:
                speak("Sorry movie doesn't exist.")

        elif "time" in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Bhai The time is {strTime}.")

        elif "visual studio" in query:
            os.startfile(CODE_PATH)

        elif "open calculator" in query:
            os.system("calc.exe")

        elif "open word" in query:
            os.startfile(WORD_PATH)

        elif "open powerpoint" in query:
            os.startfile(PP_PATH)

        elif "open excel" in query:
            os.startfile(EXCEL_PATH)

        elif "browser" in query or "edge" in query:
            os.startfile(EDGE_PATH)

        elif "selfie" in query:
            cam = cv2.VideoCapture(0)
            result, image = cam.read()
            if result:
                cv2.imwrite("selfie.jpg", image)
                os.startfile("selfie.jpg")
                cv2.waitKey(1000)
                cv2.destroyWindow("selfie")
        
        elif "message" in query:
            sendMessage()

        elif "email" in query:
            try:
                speak("To whom do you want to send the mail?")
                to = EMAIL_ADDRESS[takeCommand().lower()]
                speak("What should I say?")
                content = takeCommand()
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry. I cant send the email at this moment.")

        elif "sleep" in query:
            break

        else:
            do_calculations(query)