import pyttsx3 as tts
import datetime as dt
import speech_recognition as asr
import wikipedia
import webbrowser
import os
import sys
import json
import requests
import random
import pyautogui
from datetime import datetime, timedelta

class LunaAssistant:
    def __init__(self):
        self.tts_engine = tts.init('sapi5')
        self.voices = self.tts_engine.getProperty('voices')
        self.tts_engine.setProperty('voice', self.voices[1].id)
        self.tts_engine.setProperty('rate', 175)  # Adjust speaking rate
        
        self.recognizer = asr.Recognizer()
        self.recognizer.pause_threshold = 1
        
        self.user_name = "User"
        
        self.command_history = []
        
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
            
        except asr.RequestError:
            self.speak("Sorry, I'm having trouble connecting to the speech service.")
            return "None"
        except asr.WaitTimeoutError:
            self.speak("I didn't hear anything. Could you please speak again?")
            return "None"
        except Exception as e:
            print(f"Error: {str(e)}")
            self.speak("I didn't catch that. Could you please repeat?")
            return "None"

    def wish_me(self):
        """Greet user based on time of day"""
        hour = dt.datetime.now().hour
        greetings = {
            range(0, 5): "Yo night owl, What's keeping you up!",
            range(5, 8): "Yo, you're up before the sun!",
            range(8, 12): "Heyy, how's your morning shaping up!",
            range(12, 16): "Yo, what's the afternoon scene looking like!",
            range(16, 21): "Heyy, how's your evening rolling!",
            range(21, 24): "Heyy, how's your night going?"
        }
        
        greeting = next((msg for time_range, msg in greetings.items() 
                        if hour in time_range), "Hello!")
        self.speak(f"{greeting} What can I do for you?")

    def take_screenshot(self):
        """Capture and save a screenshot"""
        try:
            screenshot = pyautogui.screenshot()
            filename = f"screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            screenshot.save(filename)
            self.speak(f"Screenshot saved as {filename}")
        except Exception as e:
            self.speak("Sorry, I couldn't take the screenshot")

    def tell_joke(self):
        """Tell a random joke"""
        joke = random.choice(self.joke_responses)
        self.speak(joke)

    def process_command(self, query):
        """Process and execute voice commands"""
        if not query or query == "None":
            return True

        if "wikipedia" in query:
            self.speak("Searching Wikipedia...")
            query = query.replace("wikipedia", "")
            try:
                results = wikipedia.summary(query, sentences=2)
                self.speak("According to Wikipedia")
                self.speak(results)
            except:
                self.speak("Sorry, I couldn't find that on Wikipedia")
                
        elif "open youtube" in query:
            webbrowser.open("youtube.com")
            self.speak("Opening YouTube")
            
        elif "open google" in query:
            webbrowser.open("google.com")
            self.speak("Opening Google")

        elif "screenshot" in query:
            self.take_screenshot()
            
        elif "tell me a joke" in query:
            self.tell_joke()
            
        elif "weather" in query:
            self.get_weather()

        elif "time" in query:
            current_time = datetime.now().strftime("%I:%M %p")
            self.speak(f"The current time is {current_time}")
            
        elif "set reminder" in query:
            self.speak("What should I remind you about?")
            reminder_text = self.take_command()
            self.speak("In how many minutes?")
            try:
                minutes = int(self.take_command())
                self.set_reminder(reminder_text, minutes)
            except:
                self.speak("Sorry, I couldn't understand the time")

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