import os
import sys

import random
import subprocess
import re
import datetime as dt
import threading
import pyttsx3 as tts
import speech_recognition as asr
import wikipedia
import webbrowser
import psutil
import pygame
import requests
import queue
from bs4 import BeautifulSoup
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

class LunaAssistant:
    """
    A voice-controlled personal assistant that can perform various tasks including:
    - System operations (volume control, system info)
    - Weather updates and forecasts
    - Task management (todos, notes)
    - Entertainment features (jokes, facts, quotes)
    - Application and website management
    """

    def __init__(self):
        self._initialize_speech_engine()
        self._initialize_recognizer()
        self._initialize_storage()
        self._initialize_systems()
        self._load_resource_paths()
        self._load_content_data()

    def _initialize_speech_engine(self):
        """Initialize and configure the text-to-speech engine"""
        try:
            self.tts_engine = tts.init('sapi5')
            self.voices = self.tts_engine.getProperty('voices')
            self.tts_engine.setProperty('voice', self.voices[1].id)
            self.tts_engine.setProperty('rate', 175)
        except Exception as e:
            print(f"Error initializing speech engine: {e}")
            sys.exit(1)

    def _initialize_recognizer(self):
        """Initialize speech recognition components"""
        self.recognizer = asr.Recognizer()
        self.recognizer.pause_threshold = 1
        pygame.mixer.init()
        self.command_queue = queue.Queue()

    def _initialize_storage(self):
        """Initialize storage for user data and preferences"""
        self.user_name = "User"
        self.default_location = "Ahmedabad"
        self.command_history = []
        self.todos = []
        self.notes = {}
        self.timers = {}
        self.alarms = {}

    def _initialize_systems(self):
        self.devices = AudioUtilities.GetSpeakers()
        self.interface = self.devices.Activate(IAudioEndpointVolume._iid_, 1, None)
        self.volume = self.interface.QueryInterface(IAudioEndpointVolume)

    def _load_resource_paths(self):
        """Load system paths and common directories"""
        self.apps = {
            "notepad": "notepad.exe",
            "calculator": "calc.exe",
            "paint": "mspaint.exe",
            "word": r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
            "excel": r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
            "powerpoint": r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
            "chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            "vlc": r"C:\Program Files\VideoLAN\VLC\vlc.exe",
            "control panel": "control",
            "file explorer": "explorer.exe"
        }

        self.common_dirs = {
            "desktop": os.path.join(os.path.expanduser("~"), "Desktop"),
            "documents": os.path.join(os.path.expanduser("~"), "Documents"),
            "downloads": os.path.join(os.path.expanduser("~"), "Downloads"),
            "music": os.path.join(os.path.expanduser("~"), "Music"),
            "pictures": os.path.join(os.path.expanduser("~"), "Pictures"),
            "videos": os.path.join(os.path.expanduser("~"), "Videos")
        }

        self.websites = {
            "google": "https://www.google.com",
            "youtube": "https://www.youtube.com",
            "github": "https://www.github.com",
            "wikipedia": "https://www.wikipedia.org",
            "gmail": "https://www.gmail.com",
            "amazon": "https://www.amazon.com",
            "netflix": "https://www.netflix.com"
        }

    def _load_content_data(self):
        """Load entertainment content and response data"""
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

        self.jokes = [
            "Why don't scientists trust atoms? Because they make up everything!",
            "What do you call a bear with no teeth? A gummy bear!",
            "Why did the scarecrow win an award? He was outstanding in his field!",
            "I told my computer I needed a break, and now it won't stop sending me KitKats.",
            "Why don't skeletons fight each other? They don't have the guts."
        ]

        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }

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

    def speak(self, audio):
        """Speak text while listening for stop command"""
        try:
            # Set speaking flag and start background listener
            self.is_speaking = True
            listener_thread = threading.Thread(target=self.background_listener)
            listener_thread.daemon = True  # Thread will end when main program ends
            listener_thread.start()

            # Split text into sentences for better interrupt handling
            sentences = audio.split('. ')
            
            for sentence in sentences:
                if not self.is_speaking:
                    break
                    
                try:
                    # Check if stop command received
                    try:
                        if self.command_queue.get_nowait() == "stop":
                            print("Stopping speech...")
                            self.is_speaking = False
                            break
                    except queue.Empty:
                        pass

                    # Speak the sentence
                    self.tts_engine.say(sentence)
                    self.tts_engine.runAndWait()
                    
                except Exception as e:
                    print(f"Error speaking sentence: {e}")
                    break

        finally:
            # Clean up
            self.is_speaking = False
            if listener_thread.is_alive():
                listener_thread.join(timeout=1)

    def background_listener(self):
        """Continuously listen for the stop command while speaking"""
        try:
            with asr.Microphone() as source:
                while self.is_speaking:
                    try:
                        # Short timeout to allow checking is_speaking flag
                        audio = self.recognizer.listen(source, timeout=0.5, phrase_time_limit=1)
                        try:
                            text = self.recognizer.recognize_google(audio, language='en-in')
                            if "stop" in text.lower():
                                print("Stop command detected!")
                                self.command_queue.put("stop")
                                break
                        except asr.UnknownValueError:
                            continue  # Ignore unrecognized speech
                    except asr.WaitTimeoutError:
                        continue  # Keep listening if timeout occurs
        except Exception as e:
            print(f"Background listener error: {e}")

    def take_command(self):
        """Listen for and recognize voice commands with improved error handling"""
        try:
            with asr.Microphone() as source:
                print("Luna is Listening...")
                # Dynamically adjust for ambient noise
                self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=5)

            query = self.recognizer.recognize_google(audio, language='en-in')
            print(f"User said: {query}")
            self.command_history.append(query)
            return query.lower()

        except asr.WaitTimeoutError:
            print("Listening timed out. Please try again.")
            return "None"
        except asr.UnknownValueError:
            print("Could not understand audio")
            return "None"
        except asr.RequestError as e:
            print(f"Could not request results; {e}")
            return "None"
        except Exception as e:
            print(f"Error: {e}")
            return "None"
    
    def tell_fact(self, query=None):
        """Speak a random fact from the loaded facts list."""
        fact = random.choice(self.facts)
        self.speak(f"Here's a fun fact: {fact}")
        return True

    def tell_joke(self, query=None):
        """Speak a random joke from the loaded jokes list."""
        joke = random.choice(self.jokes)
        self.speak(f"Here's a fun fjoke: {joke}")
        return True
    
    def tell_quote(self, query=None):
        """Speak a random quote from the loaded quotes list."""
        quote = random.choice(self.quotes)
        self.speak(f"Here's a fun quote: {quote}")
        return True

    def open_website(self, url_or_name):
        """
        Opens a website in the default browser. Can handle both direct URLs 
        and common website names from the self.websites dictionary.
        
        Args:
            url_or_name (str): Website URL or common name (e.g., 'google', 'youtube')
        """
        try:
            # Check if it's a common website name
            if url_or_name.lower() in self.websites:
                url = self.websites[url_or_name.lower()]
            # Check if it's a direct URL
            elif any(url_or_name.lower().startswith(prefix) for prefix in ['http://', 'https://', 'www.']):
                url = url_or_name
                # Add https:// if the URL starts with www.
                if url_or_name.lower().startswith('www.'):
                    url = 'https://' + url_or_name
            else:
                # If it's neither, assume it's a domain and add https://www.
                url = 'https://www.' + url_or_name

            webbrowser.open(url)
            self.speak(f"Opening {url_or_name}")
            return True
            
        except Exception as e:
            self.speak(f"Sorry, I couldn't open that website. {str(e)}")
            return True
        
    def open_directory(self, dir_path):
        """
        Open a directory in File Explorer.
        
        Args:
            dir_path (str): Path or name of the directory to open
        """
        try:
            # Check if it's a common directory
            if dir_path.lower() in self.common_dirs:
                full_path = self.common_dirs[dir_path.lower()]
            else:
                # Try to expand the path if it contains environment variables
                full_path = os.path.expandvars(os.path.expanduser(dir_path))
            
            # Verify the directory exists
            if os.path.exists(full_path):
                # Open the directory in File Explorer
                subprocess.Popen(f'explorer "{full_path}"')
                self.speak(f"Opening {dir_path}")
                return True
            else:
                self.speak(f"Sorry, I couldn't find the directory {dir_path}")
                return True
                
        except Exception as e:
            self.speak(f"Sorry, I couldn't open the directory. {str(e)}")
            return True
    
    def find_application(self, app_name):
        if app_name in self.apps.keys():
            return self.apps[app_name]
        else:
            return None

    def open_application(self, app_name):
        """
        Open a Windows application.
        
        Args:
            app_name (str): Name of the application to open
        """
        try:
            # First try to find the application
            app_path = self.find_application(app_name)
            
            if app_path:
                # Use subprocess to start the application
                subprocess.Popen(app_path)
                self.speak(f"Opening {app_name}")
                return True
            else:
                # Try running directly with shell=True as a last resort
                try:
                    subprocess.Popen(app_name, shell=True)
                    self.speak(f"Opening {app_name}")
                    return True
                except:
                    self.speak(f"Sorry, I couldn't find {app_name}")
                    return True
                    
        except Exception as e:
            self.speak(f"Sorry, I couldn't open {app_name}. {str(e)}")
            return True
        
    def list_applications(self):
        """List all predefined applications that can be opened"""
        self.speak("Here are the applications I can open:")
        for app in sorted(self.apps.keys()):
            self.speak(app)

    def list_directories(self):
        """List all predefined directories that can be opened"""
        self.speak("Here are the common directories I can open:")
        for directory in sorted(self.common_dirs.keys()):
            self.speak(directory)
    
    def control_system_volume(self, action):
        """Control the system volume based on the action"""
        if action == "up":
            # Increase volume by a small increment
            current_volume = self.volume.GetMasterVolumeLevelScalar()
            new_volume = min(current_volume + 0.05, 1.0)  # Cap at max volume
            self.volume.SetMasterVolumeLevelScalar(new_volume, None)
        
        elif action == "down":
            # Decrease volume by a small increment
            current_volume = self.volume.GetMasterVolumeLevelScalar()
            new_volume = max(current_volume - 0.05, 0.0)  # Cap at min volume
            self.volume.SetMasterVolumeLevelScalar(new_volume, None)
        
        elif action == "mute":
            # Mute the system
            self.volume.SetMute(1, None)
        
        elif action == "unmute":
            # Unmute the system
            self.volume.SetMute(0, None)
        
        else:
            print("Unknown action:", action)

    def get_weather(self, location=None):
        """Get current weather information with improved error handling and validation"""
        try:
            location = location or self.default_location
            url = f"https://www.weather.com/weather/today/l/{location}"
            
            response = requests.get(url, headers=self.headers, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Extract weather data with more robust error checking
            temperature = self._safe_extract(soup, 'span', {'data-testid': 'TemperatureValue'})
            condition = self._safe_extract(soup, 'div', {'data-testid': 'wxPhrase'})
            humidity = self._safe_extract(soup, 'span', {'data-testid': 'PercentageValue'})
            
            weather_report = f"Current weather in {location}: {temperature}, {condition}. Humidity: {humidity}"
            self.speak(weather_report)
            return True
            
        except requests.RequestException as e:
            self.speak(f"Sorry, I couldn't fetch the weather information. {str(e)}")
            return True
        except Exception as e:
            self.speak("Sorry, I had trouble processing the weather information.")
            return True

    def _safe_extract(self, soup, tag, attributes):
        """Safely extract data from BeautifulSoup object"""
        element = soup.find(tag, attributes)
        return element.text if element else "Not available"

    def get_system_info(self):
        """Get detailed system resource information"""
        try:
            cpu = psutil.cpu_percent(interval=1)
            memory = psutil.virtual_memory()
            disk = psutil.disk_usage('/')
            battery = psutil.sensors_battery()

            info = (f"CPU usage: {cpu}%. "
                   f"Memory usage: {memory.percent}%. "
                   f"Disk usage: {disk.percent}%.")

            if battery:
                info += (f" Battery at {battery.percent}% "
                        f"({'charging' if battery.power_plugged else 'not charging'})")

            self.speak(info)
        except Exception as e:
            self.speak(f"Error getting system information: {str(e)}")

    def play_number_game(self, query=None):
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
        """Process and execute voice commands with improved command handling"""
        if not query or query == "None":
            self.speak("Sorry, I couldn't hear that. Could you please say it again?")
            return True

        # Create command mapping for better organization
        command_map = {
            "weather": self._handle_weather_command,
            "system info": self._handle_system_command,
            "volume": self._handle_volume_command,
            "todo": self._handle_todo_command,
            "note": self._handle_note_command,
            "fact": self.tell_fact,
            "joke": self.tell_joke,
            "quote": self.tell_quote,
            "game": self.play_number_game,
            "wikipedia": self._handle_wikipedia_command,
            "list:": self._handle_list_command,
        }

        # Check for exit commands first
        if any(word in query for word in ["stop", "quit", "exit", "bye"]):
            self.speak("Goodbye! Have a great day!")
            return False

        # Process commands using command map
        for key, handler in command_map.items():
            if key in query:
                return handler(query)

        # Handle application and website opening
        if query.startswith(("open ", "launch ")):
            return self._handle_open_command(query)

        return True

    def _handle_weather_command(self, query):
        """Handle weather-related commands"""
        if "forecast" in query:
            location_match = re.search(r"forecast (?:for|in) (.+)", query)
            location = location_match.group(1) if location_match else None
            return self.get_weather_forecast(location)
        else:
            location_match = re.search(r"weather (?:for|in) (.+)", query)
            location = location_match.group(1) if location_match else None
            return self.get_weather(location)

    def _handle_system_command(self, query): 
        """Handle system-related commands"""
        self.get_system_info()
        return True

    def _handle_volume_command(self, query):
        """Handle volume control commands"""
        if "up" in query:
            self.control_system_volume("up")
        elif "down" in query:
            self.control_system_volume("down")
        elif "mute" in query:
            self.control_system_volume("mute")
        elif "unmute" in query:
            self.control_system_volume("unmute")
        return True

    def _handle_todo_command(self, query):
        """Handle todo list commands"""
        if "add" in query:
            self.create_todo()
        elif "show" in query:
            self.show_todos()
        return True

    def _handle_note_command(self, query):
        """Handle note-taking commands"""
        if "take" in query:
            self.take_note()
        elif "read" in query:
            self.read_note()
        return True

    def _handle_wikipedia_command(self, query):
        """Handle Wikipedia search commands"""
        self.speak("Searching Wikipedia...")
        query = query.replace("wikipedia", "").strip()
        try:
            results = wikipedia.summary(query, sentences=2)
            self.speak("According to Wikipedia")
            self.speak(results)
        except Exception as e:
            self.speak("Sorry, I couldn't find that on Wikipedia")
        return True

    def _handle_open_command(self, query):
        """Handle commands to open applications or websites"""
        target = query.replace("open ", "").replace("launch ", "").strip()
        
        # Check if it's a website command
        if any(site in target for site in self.websites.keys()):
            return self.open_website(target)
            
        # Check if it's a directory command
        if "folder" in target or "directory" in target:
            target = target.replace("folder", "").replace("directory", "").strip()
            return self.open_directory(target)
            
        # Default to application
        return self.open_application(target)
    
    def _handle_list_command(self, query):
        """Handle user queries about listing applications or directories."""
        
        if "applications" in query or "apps" in query:
            self.list_applications()
            return True
        elif "directories" in query or "folders" in query:
            self.list_directories()
            return True
        else:
            self.speak("I'm sorry, I didn't understand what you want me to list. Please specify applications or directories.")
            return True