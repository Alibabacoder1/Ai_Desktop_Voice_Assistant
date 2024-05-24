# Importing the necessary modules
import threading
import random
import time
import subprocess
import speech_recognition as sr
from win32com.client import Dispatch
import pyttsx3
import os
import pyautogui
import datetime
import webbrowser
import requests
import winshell
import json
import sqlite3
#brightness control
import cv2 as cv
import numpy as np
import screen_brightness_control as scb
from cvzone.HandTrackingModule import HandDetector
import time 
#volume control
import cv2
import numpy as np
from cvzone.HandTrackingModule import HandDetector
from pywintypes import com_error
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import time 


# Connecting the Assistant file with the database
with sqlite3.connect('Desktop_Ai_Assistant.db') as conn:
    cursor = conn.cursor()
        
# Initialize the recognizer
recognizer = sr.Recognizer()

# Initialize text-to-speech engine
engine = pyttsx3.init()

# To change the rate of Speech
engine.setProperty('rate', 145)

# To change the gender(speech)  0 -> male and 1 -> female
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

# Global variable to indicate whether the assistant is active
is_active = False


# Function to open the camera
def open_camera():
    speak("Opening camera")
    os.system("start microsoft.windows.camera:")
    sleep() 



# Function to empty the recycle bin        
def empty_recycle_bin():
    try:
        winshell.recycle_bin().empty(confirm=False, show_progress=False)
        speak("Recycle bin emptied successfully.")
        sleep()
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        

                       
# Function to tell the date
def tell_date():
    current_date = datetime.datetime.now().strftime("%A, %d %B %Y")
    speak(f"Today is {current_date}")
    sleep() 



# Function to tell time
def tell_time():
    current_time = datetime.datetime.now().strftime("%I:%M %p")
    speak(f"The time is {current_time}")
    sleep()  



# Function for web searching
def search_web():
    speak(f"{data} What do you want to search for?")
    with sr.Microphone() as source:
        audio = recognizer.listen(source, timeout=10)
    try:
        query = recognizer.recognize_google(audio)
        url = f"https://www.google.com/search?q={query}"
        webbrowser.open(url)
        speak(f"Searching the web for {query}...")
        sleep()  
    except sr.UnknownValueError:
        speak("Could not understand audio. Please repeat.")
    except sr.RequestError as e:
        speak(f"Could not request results; {e}")
        

        
# Function to terminate/close the application
def terminate():
    global is_active
    is_active = False
    speak("Ok  Closing  the  application.")
    speak(f" Have a nice day  {data}")
    exit()



# Function to get weather information
def get_weather_info():
    speak("Please tell me the  city name")
    with sr.Microphone() as source:
        audio = recognizer.listen(source, timeout=10)
    try:
        city_name = recognizer.recognize_google(audio)
        api_key = "aec4737c284ebc75c33726a06a3a35b5"
        base_url = "http://api.openweathermap.org/data/2.5/weather?"
        complete_url = f"{base_url}q={city_name}&appid={api_key}"

        response = requests.get(complete_url)
        x = response.json()

        if x["cod"] != "404":
            y = x["main"]
            current_temperature_kelvin = y["temp"]
            current_pressure = y["pressure"]
            current_humidity = y["humidity"]
            z = x["weather"]
            weather_description = z[0]["description"]

            # Convert temperature from Kelvin to Celsius
            current_temperature_celsius = round(current_temperature_kelvin - 273.15)

            #speak(f"Weather in {city_name}:")
            print(f"Temperature is {current_temperature_celsius} degrees Celsius")
            print(f"And it is currently {weather_description}")
            speak(f"Temperature: {current_temperature_celsius} degrees Celsius")
            #speak(f"Atmospheric pressure: {current_pressure} hPa")
            #speak(f"Humidity: {current_humidity}%")
            speak(f"And it is currently: {weather_description}")
            print()
            sleep()  # Enter sleep mode after performing the action
        else:
            speak("City not found.")

    except sr.UnknownValueError:
        speak("Could not understand your city name. Please repeat.")
    except sr.RequestError as e:
        speak(f"Could not request results; {e}")



# Function to get top news
def NewsFromBBC():
    # BBC news api
    # following query parameters are used
    # source, sortBy, and apiKey
    query_params = {
        "source": "bbc-news",
        "sortBy": "top",
        "apiKey": "0ee0aeedfb5d4573b07f2e78bd5e5c0e"
    }
    main_url = "https://newsapi.org/v1/articles"

    # fetching data in JSON format
    res = requests.get(main_url, params=query_params)
    open_bbc_page = res.json()

    # getting all articles in a string article
    article = open_bbc_page["articles"]

    # empty list which will contain all trending news
    results = []
    
    for ar in article:
        results.append(ar["title"])
        
    for i in range(len(results)):
        
        # printing all trending news
        print(i + 1, results[i])

    # to read the news out loud for us
    print()
    speak = Dispatch("SAPI.Spvoice")
    speak.Speak(results)  # This should work after importing Dispatch
    sleep()  # Enter sleep mode after performing the action



# Function to open a web application based on user input
def open_web_app():
    speak("Which web application would you like to open?")
    with sr.Microphone() as source:
        audio = recognizer.listen(source, timeout=10)

    try:
        user_input = recognizer.recognize_google(audio).lower()
        try:
            url = "https://www." + user_input + ".com"
            webbrowser.open(url)
            speak(f"Opening {user_input} in the default web browser.")
            sleep()  # Enter sleep mode after performing the action
        except Exception as e:
            speak(f"An error occurred while opening {user_input}: {str(e)}")
    except sr.UnknownValueError:
        speak("Could not understand which web application to open. Please repeat.")
    except sr.RequestError as e:
        speak(f"Could not request results; {e}")



# Function to play the music
def play_music():
    # Check if the folder exists
    cursor.execute("SELECT data FROM user WHERE user=?", ("music",))
    music_data = cursor.fetchone()
    if music_data:
        folder_path = music_data[0]
            
    if not os.path.exists(folder_path) or not os.path.isdir(folder_path):
        print(f"Error: Folder '{folder_path}' not found.")
        return

    # Get a list of video files in the folder
    video_files = [f for f in os.listdir(folder_path) if f.endswith(('.mp4','.mp3', '.avi', '.mkv', '.mov', '.wmv','.mpeg'))]

    if not video_files:
        print(f"No video files found in the folder '{folder_path}'.")
        return

    # Select a random video file
    random_video = random.choice(video_files)

    # Construct the full path to the selected video file
    video_path = os.path.join(folder_path, random_video)

    # Open the default web browser to play the video file
    webbrowser.open(video_path)


# Function to locate a place on map
def locate_place_on_maps():
    speak("Please say the name of the place:")
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)

    try:
        place_name = recognizer.recognize_google(audio)
        if place_name:
            speak(f"You said: {place_name}")
            google_maps_url = "https://www.google.com/maps/search/?api=1&query=" + place_name
            webbrowser.open(google_maps_url)
            speak(f"Opening the map for {place_name}")
            sleep()
        else:
            speak("Sorry, I could not understand your voice.")
    except sr.UnknownValueError:
        speak("Sorry, I could not understand your voice.")
    except sr.RequestError as e:
        speak(f"Could not request results: {e}")



# Function to restart the computer
def restart_computer():
    speak("Restarting your computer")
    os.system("shutdown /r /t 0")



# Function which tells the users email    
def my_email():
    cursor.execute("SELECT data FROM user WHERE user=?", ("email",))
    email = cursor.fetchone()
    if email:
        found = email[0]
        speak(f"your email is {found}")
        print(f"your email is : {found}")



# Function which tells the users phone number
def my_number():
    cursor.execute("SELECT data FROM user WHERE user=?", ("phone",))
    number = cursor.fetchone()   
    if number:
        found=number[0]
        speak(f"your phone number is {found}")
        print(f"your mobile number is {found}")



# Function to adjust the brightness
def brightness():
    speak("ok adjust the brightness by using your hand gesture")
    cap = cv2.VideoCapture(0)
    hd = HandDetector()
    val = 0

    start_time = time.time()  # Get the current time

    while True:
        _, img = cap.read()
        hands, img = hd.findHands(img)

        if hands:
            lm = hands[0]['lmList']
            length, info, img = hd.findDistance(lm[8][0:2], lm[4][0:2], img)
            blevel = np.interp(length, [25, 145], [0, 100])
            val = np.interp(length, [0, 100], [400, 150])
            blevel = int(blevel)

            scb.set_brightness(blevel)

            cv2.rectangle(img, (20, 150), (85, 400), (0, 255, 255), 4)
            cv2.rectangle(img, (20, int(val)), (85, 400), (0, 0, 255), -1)
            cv2.putText(img, str(blevel) + '%', (20, 430), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 0, 0), 3)

        cv2.imshow('Brightness Control', img)

        elapsed_time = time.time() - start_time  # Calculate elapsed time

        if elapsed_time >= 12.0:  # Terminate the program after 5 seconds
            break

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    # Release the camera and close the camera window
    cap.release()
    cv2.destroyAllWindows()
    cap.release()
    sleep()  # Enter sleep mode after performing the action
    

    
# Function to adjust the volume
def set_volume(volume_level):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMasterVolumeLevelScalar(volume_level, None)

def volume():
    speak("ok adjust the volume by using your hand gesture")
    # Open the camera window
    cap = cv2.VideoCapture(0)
    hd = HandDetector()

    start_time = time.time()  # Get the current time

    while True:
        _, img = cap.read()
        hands, img = hd.findHands(img)

        if hands:
            lm = hands[0]['lmList']
            p1 = lm[4]  # Tip of thumb
            p2 = lm[8]  # Tip of index finger

            # Calculate the distance between two points (p1 and p2)
            length = np.linalg.norm(np.array(p2) - np.array(p1))

            # Map the length to a volume level between 0 and 1
            volume_level = np.interp(length, [20, 150], [0, 1])

            # Set the volume level
            set_volume(volume_level)

            cv2.putText(img, f'Volume: {int(volume_level * 100)}%', (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)

        cv2.imshow('Volume Control', img)

        elapsed_time = time.time() - start_time  # Calculate elapsed time

        if elapsed_time >= 12.0:  # Terminate the program after 5 seconds
            break

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    # Release the camera and close the camera window
    cap.release()
    cv2.destroyAllWindows()
    cap.release()
    sleep()  # Enter sleep mode after performing the action



# Function to shutdown the computer
def shut_down():
    speak("Shutting down the computer...")
    os.system("shutdown /s /t 0")  # This command is for Windows    



# Function which takes the screenshot
def take_screenshot():
    speak("Taking the picture of the current screen")
    # Specify the directory where you want to save the screenshot
    #save_directory = "C:\\Users\\sahil\\OneDrive\\Pictures\\assistant_screenshot"
    
    cursor.execute("SELECT data FROM user WHERE user=?", ("screenshot",))
    screenshot_data = cursor.fetchone()
    if screenshot_data:
        save_directory = screenshot_data[0]

    # Ensure the save directory exists
    if not os.path.exists(save_directory):
        os.makedirs(save_directory)

    # Create a unique filename using a timestamp
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"screenshot_{timestamp}.png"

    # Check if the file already exists
    while os.path.exists(os.path.join(save_directory, filename)):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"screenshot_{timestamp}.png"

    # Construct the full path
    save_path = os.path.join(save_directory, filename)

    # Capture the entire screen
    screenshot = pyautogui.screenshot()
    screenshot.save(save_path)

    speak(f"Screenshot saved sucessfully")
    sleep()



# Function which opens desktop app and system apps
def open_apps():
    speak("Which application do you want to open?")
    
    # Listen for the application name
    with sr.Microphone() as source:
        audio = recognizer.listen(source, timeout=10)
    
    try:
        app_name = recognizer.recognize_google(audio).lower()
        if app_name:
            speak(f"Opening {app_name}")
            
            # Use the provided logic to open the application
            command = f"open {app_name}"
            command = command.replace("open", "").replace("pluto", "")
            pyautogui.press("super")
            time.sleep(1)
            pyautogui.typewrite(command)
            time.sleep(1)
            pyautogui.press("enter")
            
            # Go to sleep mode
            sleep()
        else:
            speak("Sorry, I could not understand the application name.")
    except sr.UnknownValueError:
        speak("Sorry, I could not understand the application name. Please repeat.")
    except sr.RequestError as e:
        speak(f"Could not request results; {e}")
        

         
# Function which tells users name                
def my_name():
    speak(f"your name is {data}")
    
    
# Chat_gtp functions searches the info from the open_ai from the openai_api

def search_chat_gpt():
    print("complete this function")
    
    
 
       
# Function to write the note        
def speech_to_text_and_save():
    try:
        cursor.execute("SELECT data FROM user WHERE user=?", ("note",))
        note_data = cursor.fetchone()
        if note_data:
            output_folder = note_data[0]
        
        # Use the default microphone as the audio source
        with sr.Microphone() as source:
            speak("Please start speaking...")

            # Adjust for ambient noise and record the audio with a 20-second timeout
            recognizer.adjust_for_ambient_noise(source)
            audio = recognizer.listen(source, timeout=20)

            if audio:
                print("Taking a note...")

                # Use the Google Web Speech API to transcribe the audio
                text = recognizer.recognize_google(audio)

                # Print the transcribed text
                print("Transcription:")
                print(text)

                # Generate a file name with the current date and time
                now = datetime.datetime.now()
                current_time = now.strftime("%Y-%m-%d_%H-%M-%S")
                file_name = f"transcribed_note_{current_time}.txt"

                # Create the full path to the file in the specified folder
                full_path = os.path.join(output_folder, file_name)

                # Save the transcribed text to the specified folder
                with open(full_path, "w") as file:
                    file.write(text)
                    speak(f"Transcribed note saved as '{full_path}'")
                    sleep()

            else:
                print("No speech detected for more than 20 seconds. File not saved.")

    except sr.UnknownValueError:
        print("Could not understand audio")
    except sr.RequestError as e:
        print(f"Could not request results; {e}")
        
 
        
# Function to register a new user                 
def add_face():
    speak("OK  adding a new face ")
    speak("after  hundred  frames  are  captured  the  application  will  closed  and  you  need  to  restart  the  application")
    speak("please  align  the  person  in  front  of  the  camera  to  capture  the  frames")
    script_path = './Collect_Data.py'
    subprocess.Popen(['python', script_path])
    exit()


    
# Define a dictionary to map keywords to functions
keyword_actions = {
    'camera': open_camera,
    'empty recycle bin': empty_recycle_bin,
    'date': tell_date,
    'time': tell_time,
    'search': search_web,
    'terminate': terminate,
    'close': terminate,
    'news': NewsFromBBC,
    'music': play_music,
    'song': play_music,
    'web': open_web_app,
    'weather': get_weather_info,
    'recycle': empty_recycle_bin,
    'bin': empty_recycle_bin,
    'locate': locate_place_on_maps,
    'location': locate_place_on_maps,
    'map': locate_place_on_maps,
    'where': locate_place_on_maps,
    'restart': restart_computer,
    'shutdown':shut_down,
    'shut down':shut_down,
    'switch off':shut_down,
    'turn off':shut_down,
    'screenshot':take_screenshot,
    'snapshot':take_screenshot,
    'desktop':open_apps,
    'note' : speech_to_text_and_save,
    'brightness':brightness,
    'volume':volume,
    'name':my_name,
    'email':my_email,
    'mobile':my_number,
    'register':add_face,
    'chat_gpt':search_chat_gpt,
    'chatgpt':search_chat_gpt,
    'add':add_face,
    'face':add_face
}



# Function to perform actions based on recognized keywords
def perform_action(command):
    for keyword, action in keyword_actions.items():
        if keyword in command:
            action()
            return



# Function to enter sleep mode
def sleep():
    global is_active
    is_active = False
    speak(" Going  to  sleep.")



# Function to activate the assistant
def activate_assistant():
    global is_active
    is_active = True
    speak("Pluto here , what can I do for you?")



# Function to speak text
def speak(text,rate=145):
    engine.setProperty('rate', rate)
    engine.say(text)
    engine.runAndWait()



                            # Main loop for voice recognition
# Greet Script
current_time = datetime.datetime.now().time()
morning_start = datetime.time(5, 0, 0)
afternoon_start = datetime.time(12, 0, 0)
evening_start = datetime.time(17, 0, 0)


if morning_start <= current_time < afternoon_start:
    greet="Good morning!"
elif afternoon_start <= current_time < evening_start:
    greet="Good afternoon!"
else:
    greet="Good evening!"
    
cursor.execute("SELECT data FROM user WHERE user = ?", ("name",))
row = cursor.fetchone()  # Fetch the first matching row
if row is not None:
    data = row[0]
""" print(f" {greet}  Welcome  {data} ,  My name is Pluto,  I am your desktop assistant , Call me by saying  'Hey Pluto' for any assistance") """
speak(f" {greet}  Welcome  {data} ,  My name is Pluto,  I am your desktop assistant , Call me by saying  'Hey Pluto' for any assistance")
     


# Loop execution unitl Terminated
while True:
    try:
        with sr.Microphone() as source:
            if not is_active:
                print("Listening for the wake word...")
                print()
                recognizer.adjust_for_ambient_noise(source)  # Adjust for ambient noise
                audio = recognizer.listen(source , timeout=5) 
                wake_word = recognizer.recognize_google(audio).lower()
                if "hey" in wake_word or "pluto" in wake_word or "heypluto" in wake_word or "hey pluto" in wake_word:
                    activate_assistant()
             
            elif is_active:
                print("Listening for a command...")
                recognizer.adjust_for_ambient_noise(source)  # Adjust for ambient noise
                audio = recognizer.listen(source  , timeout=10)  # Extend the timeout for commands 
                print("Recognizing...")

                # Recognize speech using Google Web Speech API
                command = recognizer.recognize_google(audio).lower()
                print("You said:", command)
                

                # Extract keywords and perform actions
                perform_action(command)

    except sr.UnknownValueError:
        if is_active:
            speak("Could not understand audio. Please repeat.")
    except sr.RequestError as e:
        speak(f"Could not request results; {e}")
    except KeyboardInterrupt:
        break



# Stop the text-to-speech engine on exit
conn.close()
engine.stop()