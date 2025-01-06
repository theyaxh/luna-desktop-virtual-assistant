import wikipedia
import datetime as dt
from speech_manager import SpeechManager
from system_manager import SystemManager
from app_manager import AppManager
from entertainment import EntertainmentManager

class LunaAssistant:
    def __init__(self):
        self.speech = SpeechManager()
        self.system = SystemManager()
        self.app = AppManager()
        self.entertainment = EntertainmentManager()

    def wish_me(self):
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
        self.speech.speak(f"{greeting} What can I do for you?")

    def process_command(self, query):
        if not query or query == "None":
            self.speech.speak("Sorry, I couldn't hear that. Could you please say it again?")
            return True

        if any(word in query for word in ["stop", "quit", "exit", "bye"]):
            self.speech.speak("Goodbye! Have a great day!")
            return False

        command_handlers = {
            "screenshot": self._handle_screenshot,
            "system info": self._handle_system_info,
            "volume": self._handle_volume,
            "fact": self._handle_entertainment,
            "joke": self._handle_entertainment,
            "quote": self._handle_entertainment,
            "wikipedia": self._handle_wikipedia,
        }

        for key, handler in command_handlers.items():
            if key in query:
                return handler(query)

        if query.startswith(("open ", "launch ")):
            return self._handle_open(query)

        return True

    def _handle_screenshot(self, query):
        filepath = self.system.take_screenshot()
        self.speech.speak("Screenshot saved" if filepath else "Failed to take screenshot")
        return True

    def _handle_system_info(self, query):
        info = self.system.get_system_info()
        self.speech.speak(info)
        return True

    def _handle_volume(self, query):
        for action in ["up", "down", "mute", "unmute"]:
            if action in query:
                self.system.control_system_volume(action)
                break
        return True

    def _handle_entertainment(self, query):
        if "fact" in query:
            self.speech.speak(f"Here's a fun fact: {self.entertainment.get_random_fact()}")
        elif "joke" in query:
            self.speech.speak(f"Here's a fun joke: {self.entertainment.get_random_joke()}")
        elif "quote" in query:
            self.speech.speak(f"Here's a quote: {self.entertainment.get_random_quote()}")
        return True

    def _handle_wikipedia(self, query):
        self.speech.speak("Searching Wikipedia...")
        query = query.replace("wikipedia", "").strip()
        try:
            results = wikipedia.summary(query, sentences=1)
            self.speech.speak("According to Wikipedia")
            self.speech.speak(results)
        except Exception:
            self.speech.speak("Sorry, I couldn't find that on Wikipedia")
        return True

    def _handle_open(self, query):
        target = query.replace("open ", "").replace("launch ", "").strip()
        
        if any(site in target for site in self.app.websites):
            success = self.app.open_website(target)
        elif "folder" in target or "directory" in target:
            target = target.replace("folder", "").replace("directory", "").strip()
            success = self.app.open_directory(target)
        else:
            success = self.app.open_application(target)
        
        if success:
            self.speech.speak(f"Opening {target}")
        else:
            self.speech.speak(f"Sorry, I couldn't open {target}")
        return True