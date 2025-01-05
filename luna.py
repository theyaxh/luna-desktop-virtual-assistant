import pyttsx3 as tts
import datetime as dt
import speech_recognition as asr
import wikipedia
import webbrowser
import os
import sys
import random
import pyautogui
import psutil
import pygame
import time
import socket
from datetime import datetime, timedelta
import subprocess
import platform
from glob import glob

class LunaAssistant:
    def __init__(self):
        # Initialize text-to-speech engine
        self.tts_engine = tts.init('sapi5')
        self.voices = self.tts_engine.getProperty('voices')
        self.tts_engine.setProperty('voice', self.voices[1].id)
        self.tts_engine.setProperty('rate', 175)
        
        # Initialize speech recognizer
        self.recognizer = asr.Recognizer()
        self.recognizer.pause_threshold = 1
        
        # Store user preferences
        self.user_name = "User"
        
        # Initialize pygame for audio
        pygame.mixer.init()
        
        # Command history
        self.command_history = []
        
        # Todo list and notes
        self.todos = []
        self.notes = {}
        
        # Timer and alarm storage
        self.timers = {}
        self.alarms = {}
        
        # Random facts and quotes
        self.facts = [
            "A day on Venus is longer than its year!",
            "Honey never spoils. Archaeologists found 3000-year-old honey in Egyptian tombs!",
            "The first oranges weren't orange - they were green!",
            "Octopuses have three hearts!",
            "A cloud can weigh more than a million pounds!"
        ]
        
        self.quotes = [
            "The only way to do great work is to love what you do. - Steve Jobs",
            "Innovation distinguishes between a leader and a follower. - Steve Jobs",
            "The future belongs to those who believe in the beauty of their dreams. - Eleanor Roosevelt",
            "Success is not final, failure is not fatal. - Winston Churchill",
            "The best way to predict the future is to create it. - Peter Drucker"
        ]

        self.joke_responses = [
            "Why don't scientists trust atoms? Because they make up everything!",
            "What do you call a bear with no teeth? A gummy bear!",
            "Why did the scarecrow win an award? He was outstanding in his field!",
            "I told my computer I needed a break, and now it won't stop sending me KitKats.",
            "Why don't skeletons fight each other? They don't have the guts.",
            "I'm reading a book on anti-gravity—it's impossible to put down.",
            "Parallel lines have so much in common… it's a shame they'll never meet.",
            "I told my plants a joke, but I think they needed thyme to get it.",
            "I tried to write a joke about pizza, but it was too cheesy.",
            "I used to play piano by ear, but now I use my hands.",
            "Why don't eggs tell jokes? They'd crack each other up.",
            "I only know 25 letters of the alphabet. I don't know Y.",
            "I was going to make a belt out of watches, but it was a waist of time."
        ]

        # Common application paths
        self.apps = {
            "notepad": "notepad.exe",
            "calculator": "calc.exe",
            "paint": "mspaint.exe",
            "word": "winword.exe",
            "excel": "excel.exe"
        }

    def speak(self, audio):
        """Convert text to speech"""
        try:
            self.tts_engine.say(audio)
            self.tts_engine.runAndWait()
        except Exception as e:
            print(f"Speech error: {str(e)}")

    def take_command(self):
        """Listen for and recognize voice commands"""
        try:
            with asr.Microphone() as source:
                print("Luna is Listening...")
                self.recognizer.adjust_for_ambient_noise(source)
                audio = self.recognizer.listen(source, timeout=5)
                
            query = self.recognizer.recognize_google(audio, language='en-in')
            print(f"User said: {query}")
            self.command_history.append(query)
            return query.lower()
            
        except Exception as e:
            print(f"Error: {str(e)}")
            return "None"

    # System Operations
    def get_system_info(self):
        """Get system resource information"""
        cpu = psutil.cpu_percent()
        memory = psutil.virtual_memory()
        battery = psutil.sensors_battery()
        
        info = f"CPU usage is {cpu}%. Memory usage is {memory.percent}%."
        if battery:
            info += f" Battery is at {battery.percent}%"
            if battery.power_plugged:
                info += " and charging."
            else:
                info += f" with {battery.secsleft // 3600} hours remaining."
        
        self.speak(info)

    def control_system_volume(self, action):
        """Control system volume"""
        from ctypes import cast, POINTER
        from comtypes import CLSCTX_ALL
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
        
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        
        if "up" in action:
            volume.VolumeUp(None)
            self.speak("Volume increased")
        elif "down" in action:
            volume.VolumeDown(None)
            self.speak("Volume decreased")
        elif "mute" in action:
            volume.SetMute(1, None)
            self.speak("Audio muted")
        elif "unmute" in action:
            volume.SetMute(0, None)
            self.speak("Audio unmuted")

    def empty_recycle_bin(self):
        """Empty the Windows Recycle Bin"""
        try:
            winshell.recycle_bin().empty(confirm=False)
            self.speak("Recycle bin emptied successfully")
        except Exception as e:
            self.speak("Could not empty the recycle bin")

    # Productivity Features
    def create_todo(self):
        """Create a todo item"""
        self.speak("What would you like to add to your todo list?")
        todo = self.take_command()
        if todo != "None":
            self.todos.append(todo)
            self.speak(f"Added {todo} to your todo list")

    def show_todos(self):
        """Show all todos"""
        if not self.todos:
            self.speak("Your todo list is empty")
        else:
            self.speak("Here are your todos:")
            for i, todo in enumerate(self.todos, 1):
                self.speak(f"{i}. {todo}")

    def take_note(self):
        """Take a note with a title"""
        self.speak("What's the title of your note?")
        title = self.take_command()
        if title != "None":
            self.speak("What's the content of your note?")
            content = self.take_command()
            if content != "None":
                self.notes[title] = content
                self.speak(f"Note saved with title: {title}")

    def read_note(self):
        """Read a specific note"""
        if not self.notes:
            self.speak("You don't have any notes")
            return
            
        self.speak("Which note would you like me to read? Here are your note titles:")
        for title in self.notes:
            self.speak(title)
            
        title = self.take_command()
        if title in self.notes:
            self.speak(f"Here's your note titled {title}: {self.notes[title]}")
        else:
            self.speak("I couldn't find that note")

    # Entertainment Features
    def tell_fact(self):
        """Tell a random fact"""
        fact = random.choice(self.facts)
        self.speak(fact)

    def tell_joke(self):
        """Tell a randomg joke"""
        joke = random.choice(self.joke_responses)
        self.speak(joke)

    def tell_quote(self):
        """Tell a random quote"""
        quote = random.choice(self.quotes)
        self.speak(quote)

    def play_number_game(self):
        """Play a simple number guessing game"""
        number = random.randint(1, 100)
        attempts = 0
        max_attempts = 7
        
        self.speak("Let's play a number guessing game! I'm thinking of a number between 1 and 100.")
        
        while attempts < max_attempts:
            self.speak("What's your guess?")
            try:
                guess = int(self.take_command())
                attempts += 1
                
                if guess < number:
                    self.speak("Higher! Try again.")
                elif guess > number:
                    self.speak("Lower! Try again.")
                else:
                    self.speak(f"Congratulations! You got it in {attempts} attempts!")
                    return
                    
            except ValueError:
                self.speak("Please say a number.")
                
        self.speak(f"Game over! The number was {number}.")

    def process_command(self, query):
        """Process and execute voice commands"""
        if not query or query == "None":
            return True

        # System Commands
        if "system info" in query or "status" in query:
            self.get_system_info()
            
        elif "volume" in query:
            if "up" in query:
                self.control_system_volume("up")
            elif "down" in query:
                self.control_system_volume("down")
            elif "mute" in query:
                self.control_system_volume("mute")
            elif "unmute" in query:
                self.control_system_volume("unmute")

        elif "empty recycle bin" in query:
            self.empty_recycle_bin()

        # Productivity Commands
        elif "add todo" in query:
            self.create_todo()
            
        elif "show todos" in query:
            self.show_todos()
            
        elif "take note" in query:
            self.take_note()
            
        elif "read note" in query:
            self.read_note()

        # Entertainment Commands
        elif "tell me a fact" in query:
            self.tell_fact()
            
        elif "tell me a quote" in query:
            self.tell_quote()

        elif "tell me a joke" in query:
            self.tell_joke()
            
        elif "play game" in query:
            self.play_number_game()

        # Previous commands remain unchanged...
        elif "wikipedia" in query:
            self.speak("Searching Wikipedia...")
            query = query.replace("wikipedia", "")
            try:
                results = wikipedia.summary(query, sentences=2)
                self.speak("According to Wikipedia")
                self.speak(results)
            except:
                self.speak("Sorry, I couldn't find that on Wikipedia")

        # Exit command
        elif any(word in query for word in ["stop", "quit", "exit", "bye"]):
            self.speak("Goodbye! Have a great day!")
            return False

        return True

def main():
    luna = LunaAssistant()
    luna.wish_me()
    
    running = True
    while running:
        query = luna.take_command()
        running = luna.process_command(query)

if __name__ == "__main__":
    main()