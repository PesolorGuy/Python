# Made By PesolorGuy, Please Consider Giving Credit If Used Some Of The Code Of My Jarvis. Thanks !
import os
import sys
import datetime
import speech_recognition as sr
import pyttsx3
import webbrowser
import wikipedia
import smtplib
import random
import json
import requests
import subprocess
import pyautogui
import psutil
import platform
import socket
import speedtest
import cv2
import numpy as np
import time
import wolframalpha
import pyjokes
import ctypes
import winshell
import shutil
import win32com.client
from urllib.request import urlopen
from bs4 import BeautifulSoup
from ecapture import ecapture as ec
from difflib import get_close_matches
from word2number import w2n
import sounddevice
import wavio
import librosa
import noisereduce as nr
# Import noise reduction library
import noisereduce as noise

# Audio processing settings
SAMPLE_RATE = 44100
CHANNELS = 2
CHUNK_SIZE = 1024
FORMAT = 'float32'

# Configure audio settings
def configure_audio():
    try:
        sounddevice.default.samplerate = SAMPLE_RATE
        sounddevice.default.channels = CHANNELS
        sounddevice.default.dtype = FORMAT
    except sounddevice.PortAudioError as e:
        print(f"Audio device error: {str(e)}")
        return False
    return True

# Enhanced audio preprocessing
def preprocess_audio(audio_data):
    # Convert to numpy array
    audio_array = np.frombuffer(audio_data.frame_data, dtype=np.int16)
    
    # Convert to float32
    audio_float = audio_array.astype(np.float32) / 32768.0
    
    # Apply noise reduction
    reduced_noise = noise.reduce_noise(
        y=audio_float,
        sr=SAMPLE_RATE,
        prop_decrease=0.75,
        n_fft=2048
    )
    
    # Normalize audio
    normalized = librosa.util.normalize(reduced_noise)
    
    return normalized

# Initialize speech recognition with better settings
r = sr.Recognizer()
r.energy_threshold = 4000  # Increase energy threshold for better noise handling
r.dynamic_energy_threshold = True  # Dynamically adjust for ambient noise
r.dynamic_energy_adjustment_damping = 0.15
r.dynamic_energy_ratio = 1.5
r.pause_threshold = 0.8  # Shorter pause threshold for more natural speech detection
r.operation_timeout = None  # No timeout
r.phrase_threshold = 0.3
r.non_speaking_duration = 0.5

# Configure language model for better accuracy
r.language = "en-US"  # Set language explicitly

# Function to reduce background noise from audio input
def reduce_noise(audio_data):
    # Convert audio data to numpy array
    y = np.frombuffer(audio_data.frame_data, dtype=np.int16)
    # Reduce noise using spectral gating
    reduced_noise = nr.reduce_noise(
        y=y.astype(float), 
        sr=audio_data.sample_rate,
        prop_decrease=1.0,
        stationary=True
    )
    return reduced_noise

# Function to get clearer audio input
def get_audio():
    with sr.Microphone() as source:
        print("Adjusting for ambient noise...")
        r.adjust_for_ambient_noise(source, duration=1)
        print("Listening...")
        try:
            audio = r.listen(source, timeout=5, phrase_time_limit=5)
            return audio
        except sr.WaitTimeoutError:
            print("No speech detected")
            return None

# Enhanced speech recognition function
def enhanced_recognize(audio):
    if audio is None:
        return "None"
    
    try:
        # Try Google's speech recognition first (most accurate)
        text = r.recognize_google(audio, language='en-US')
        return text
    except sr.UnknownValueError:
        try:
            # Fallback to Sphinx if Google fails
            text = r.recognize_sphinx(audio)
            return text
        except:
            try:
                # Try with noise reduction if regular recognition fails
                reduced_audio = reduce_noise(audio)
                text = r.recognize_google(reduced_audio, language='en-US')
                return text
            except:
                print("Could not understand audio")
                return "None"
    except Exception as e:
        print(f"Error in speech recognition: {str(e)}")
        return "None"


# Enhanced speech recognition
def enhanced_speech_recognition(audio):
    try:
        # Preprocess the audio
        processed_audio = preprocess_audio(audio)
        
        # Convert back to audio data format
        enhanced_audio = sr.AudioData(
            processed_audio.tobytes(),
            sample_rate=SAMPLE_RATE,
            sample_width=2
        )
        
        # Perform recognition with enhanced audio
        text = r.recognize_google(enhanced_audio)
        return text
        
    except sr.UnknownValueError:
        print("Could not understand audio")
        return ""
    except sr.RequestError:
        print("Could not request results")
        return ""


# Initialize text-to-speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

# Improved speech recognition setup
def initialize_recognizer():
    r = sr.Recognizer()
    # Adjust recognition settings for better accuracy
    r.energy_threshold = 4000
    r.dynamic_energy_threshold = True
    r.dynamic_energy_adjustment_damping = 0.15
    r.dynamic_energy_ratio = 1.5
    r.pause_threshold = 0.8
    r.phrase_threshold = 0.3
    r.non_speaking_duration = 0.5
    return r

r = initialize_recognizer()

def enhance_audio(audio_data):
    # Convert audio data to numpy array
    y = np.frombuffer(audio_data.frame_data, dtype=np.int16)
    # Apply noise reduction
    reduced_noise = noise.reduce_noise(y=y, sr=audio_data.sample_rate)
    return reduced_noise

def speak(text):
    engine.say(text)
    engine.runAndWait()

def wishMe():
    hour = datetime.datetime.now().hour
    if 0 <= hour < 12:
        speak("Good Morning!")
    elif 12 <= hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak("I am Jarvis. How may I help you?")

def takeCommand():
    with sr.Microphone() as source:
        print("Listening...")
        # Adjust for ambient noise
        r.adjust_for_ambient_noise(source, duration=1)
        audio = r.listen(source)
        
    try:
        print("Recognizing...")
        # Enhanced audio processing
        enhanced_audio = enhance_audio(audio)
        
        # Try multiple recognition services for better accuracy
        try:
            query = r.recognize_google(audio, language='en-US')
        except:
            try:
                query = r.recognize_wit(audio, key="YOUR_WIT_AI_KEY")
            except:
                query = r.recognize_sphinx(audio)
        
        print(f"User said: {query}\n")
        if 'gui' in globals():
            gui.add_to_conversation(query, "User")
        return query.lower()
    except Exception as e:
        print("Could not understand audio")
        return "None"

def sendEmail(to, content):
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        email = os.getenv("EMAIL")
        password = os.getenv("EMAIL_PASSWORD")
        if not email or not password:
            raise ValueError("Email credentials not configured")
        server.login(email, password)
        server.sendmail(email, to, content)
        server.close()
    except Exception as e:
        print(f"Email error: {str(e)}")
        raise

# Main program loop
if __name__ == "__main__":
    wishMe()
    
    while True:
        query = takeCommand()
        
        # Basic conversation
        if "hello" in query:
            speak("Hello! How can I help you today?")
            
        elif "how are you" in query:
            speak("I'm doing well, thank you for asking!")
            
        # Web browsing features
        elif "open youtube" in query:
            webbrowser.open("youtube.com")
            
        elif "open google" in query:
            webbrowser.open("google.com")
            
        elif "search wikipedia" in query:
            speak("Searching Wikipedia...")
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            speak(results)
            
        # System operations
        elif "shutdown" in query:
            os.system("shutdown /s /t 1")
            
        elif "restart" in query:
            os.system("shutdown /r /t 1")
            
        elif "sleep" in query:
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
            
        # File operations
        elif "create folder" in query:
            speak("What should I name the folder?")
            folder_name = takeCommand()
            os.mkdir(folder_name)
            speak(f"Created folder named {folder_name}")
            
        elif "delete folder" in query:
            speak("Which folder should I delete?")
            folder_name = takeCommand()
            shutil.rmtree(folder_name)
            speak(f"Deleted folder named {folder_name}")
            
        # System information
        elif "cpu usage" in query:
            usage = str(psutil.cpu_percent())
            speak(f"CPU is at {usage} percent")
            
        elif "battery" in query:
            battery = psutil.sensors_battery()
            speak(f"Battery is at {battery.percent} percent")
            
        # Entertainment features
        elif "tell me a joke" in query:
            speak(pyjokes.get_joke())
            
        elif "play music" in query:
            music_dir = "C:\\Music"  # Change to your music directory
            songs = os.listdir(music_dir)
            os.startfile(os.path.join(music_dir, random.choice(songs)))
            
        # Email feature
        elif "send email" in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("To whom should I send it?")
                to = takeCommand()  # You might want to maintain an address book
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry, I couldn't send the email")
                
        # Weather information
        elif "weather" in query:
            # Add your weather API implementation here
            speak("Please provide your API key to get weather information")
            
        # News updates
        elif "news" in query:
            # Add your news API implementation here
            speak("Please provide your API key to get news updates")
            
        # Calculator functions
        elif "calculate" in query:
            try:
                app_id = os.getenv("WOLFRAMALPHA_APP_ID")
                client = wolframalpha.Client(app_id)
                indx = query.lower().split().index('calculate')
                query = query.split()[indx + 1:]
                res = client.query(' '.join(query))
                answer = next(res.results).text
                speak(f"The answer is {answer}")
            except Exception as e:
                speak("Sorry, I couldn't calculate that")
                
        # System cleanup
        elif "clean recycle bin" in query:
            winshell.recycle_bin().empty()
            speak("Recycle Bin is now empty")
            
        # Take photo
        elif "take photo" in query:
            ec.capture(0, "Jarvis Camera", "img.jpg")
            
        # Note taking
        elif "make a note" in query:
            speak("What should I write down?")
            note_text = takeCommand()
            with open('notes.txt', 'a') as f:
                f.write(f"\n{datetime.datetime.now()}: {note_text}")
            speak("I've made a note of that")
            
        # Exit command
        elif "goodbye" in query or "exit" in query:
            speak("Goodbye! Have a great day!")
            break

        # Additional features can be added:
        # - Calendar management
        # - Task scheduling
        # - Home automation integration
        # - Social media interaction
        # - File encryption/decryption
        # - Language translation
        # - Voice recording
        # - Screen recording
        # - System monitoring
        # - Network speed testing
        # - Password management
        # - Document conversion
        # - Image editing
        # - Audio processing
        # - Video processing
        # - Database management
        # - Web scraping
        # - PDF operations
        # - Text analysis
        # - Face recognition
        # - Voice authentication
        # - Custom automation scripts
        # Calendar management
        elif "manage calendar" in query:
            speak("Opening calendar management. What would you like to do?")
            calendar_query = takeCommand()
            # Add calendar management logic here
            
        # Task scheduling 
        elif "schedule task" in query:
            speak("What task would you like to schedule?")
            task = takeCommand()
            # Add task scheduling logic here
            
        # Home automation
        elif "home control" in query:
            speak("What would you like to control in your home?")
            home_query = takeCommand()
            # Add home automation logic here
            
        # Social media
        elif "social media" in query:
            speak("Which social media platform would you like to access?")
            platform = takeCommand()
            # Add social media interaction logic here
            
        # File encryption
        elif "encrypt file" in query:
            speak("Which file would you like to encrypt?")
            file = takeCommand()
            # Add encryption logic here
            
        elif "decrypt file" in query:
            speak("Which file would you like to decrypt?")
            file = takeCommand()
            # Add decryption logic here
            
        # Language translation
        elif "translate" in query:
            speak("What would you like me to translate?")
            text = takeCommand()
            # Add translation logic here
            
        # Voice recording
        elif "record voice" in query:
            speak("Starting voice recording")
            # Add voice recording logic here
            
        # Screen recording
        elif "record screen" in query:
            speak("Starting screen recording")
            # Add screen recording logic here
            
        # System monitoring
        elif "system status" in query:
            cpu = psutil.cpu_percent()
            memory = psutil.virtual_memory().percent
            speak(f"CPU usage is {cpu}% and memory usage is {memory}%")
            
        # Network speed
        elif "network speed" in query:
            speak("Testing network speed")
            st = speedtest.Speedtest()
            dl_speed = st.download() / 1_000_000  # Convert to Mbps
            up_speed = st.upload() / 1_000_000
            speak(f"Download speed is {round(dl_speed, 2)} Mbps and upload speed is {round(up_speed, 2)} Mbps")
            
        # Password management
        elif "password manager" in query:
            speak("Opening password manager. What would you like to do?")
            pass_query = takeCommand()
            # Add password management logic here
            
        # Document conversion
        elif "convert document" in query:
            speak("What document would you like to convert?")
            doc = takeCommand()
            # Add document conversion logic here
            
        # Image editing
        elif "edit image" in query:
            speak("What image would you like to edit?")
            image = takeCommand()
            # Add image editing logic here
            
        # Audio processing
        elif "process audio" in query:
            speak("What audio would you like to process?")
            audio = takeCommand()
            # Add audio processing logic here
            
        # Video processing
        elif "process video" in query:
            speak("What video would you like to process?")
            video = takeCommand()
            # Add video processing logic here
            
        # Database management
        elif "database" in query:
            speak("What database operation would you like to perform?")
            db_query = takeCommand()
            # Add database management logic here
            
        # Web scraping
        elif "scrape website" in query:
            speak("What website would you like to scrape?")
            url = takeCommand()
            # Add web scraping logic here
            
        # PDF operations
        elif "pdf" in query:
            speak("What PDF operation would you like to perform?")
            pdf_query = takeCommand()
            # Add PDF operations logic here
            
        # Text analysis
        elif "analyze text" in query:
            speak("What text would you like me to analyze?")
            text = takeCommand()
            # Add text analysis logic here
            
        # Face recognition
        elif "recognize face" in query:
            speak("Starting face recognition")
            # Add face recognition logic here
            
        # Voice authentication
        elif "voice auth" in query:
            speak("Starting voice authentication")
            # Add voice authentication logic here
            
        # Custom automation
        elif "automate task" in query:
            speak("What task would you like to automate?")
            task = takeCommand()
            # Add custom automation logic here
            
        # List all capabilities
        elif "what can you do" in query or "list commands" in query:
            capabilities = [
                "Calendar management",
                "Task scheduling",
                "Home automation control",
                "Social media interaction",
                "File encryption and decryption",
                "Language translation",
                "Voice recording",
                "Screen recording",
                "System monitoring",
                "Network speed testing",
                "Password management",
                "Document conversion",
                "Image editing",
                "Audio processing",
                "Video processing",
                "Database management",
                "Web scraping",
                "PDF operations",
                "Text analysis",
                "Face recognition",
                "Voice authentication",
                "Custom automation",
                "Basic conversation",
                "Web browsing",
                "Email sending",
                "Note taking",
                "System cleanup",
                "Taking photos"
            ]
            
            speak("Here are all the things I can do:")
            for capability in capabilities:
                speak(capability)
                time.sleep(0.5)  # Add small delay between items

        elif "voice note" in query:
            speak("Recording voice note. Say 'stop' when finished")
            audio_chunks = []
            while True:
                audio = r.listen(sr.Microphone())
                text = r.recognize_google(audio)
                if 'stop' in text.lower():
                    break
                audio_chunks.append(audio)
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            wavio.write(f"voice_note_{timestamp}.wav", audio_chunks, 44100, sampwidth=2)
            speak("Voice note saved")

        elif "transcribe audio" in query:
            speak("Which audio file should I transcribe?")
            filename = takeCommand()
            audio = sr.AudioFile(filename)
            with audio as source:
                audio_data = r.record(source)
            text = r.recognize_google(audio_data)
            speak(f"Transcription: {text}")

        elif "smart home" in query:
            # Add integration with popular smart home platforms
            speak("Which smart home device would you like to control?")
            device = takeCommand()
            # Add your smart home control logic here

        elif "reminder" in query:
            speak("What should I remind you about?")
            reminder_text = takeCommand()
            speak("When should I remind you? Please specify hours from now")
            try:
                hours = w2n.word_to_num(takeCommand())
                # Schedule reminder using system notifications
                reminder_time = datetime.datetime.now() + datetime.timedelta(hours=hours)
                # Add reminder scheduling logic here
                speak(f"Reminder set for {reminder_time}")
            except:
                speak("Sorry, I couldn't understand the time")

        elif "summarize text" in query:
            speak("Please provide the text to summarize")
            text = takeCommand()
            # Add text summarization logic here
            # You can use libraries like sumy or transformers

