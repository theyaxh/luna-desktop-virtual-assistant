import pyttsx3 as tts
import speech_recognition as asr
import sys

class SpeechManager:
    def __init__(self):
        self._initialize_speech_engine()
        self._initialize_recognizer()

    def _initialize_speech_engine(self):
        try:
            self.tts_engine = tts.init('sapi5')
            self.voices = self.tts_engine.getProperty('voices')
            self.tts_engine.setProperty('voice', self.voices[1].id)
            self.tts_engine.setProperty('rate', 175)
        except Exception as e:
            print(f"Error initializing speech engine: {e}")
            sys.exit(1)

    def _initialize_recognizer(self):
        self.recognizer = asr.Recognizer()

    def speak(self, audio):
        try:
            self.tts_engine.say(audio)
            self.tts_engine.runAndWait()
        except Exception as e:
            print(f"Error speaking sentence: {e}")

    def take_command(self):
        try:
            with asr.Microphone() as source:
                print("Luna is Listening...")
                audio = self.recognizer.listen(source, timeout=10)

            query = self.recognizer.recognize_google(audio, language='en-in')
            print(f"User said: {query}")
            return query.lower()
        except Exception as e:
            print(f"Error: {e}")
            return "None"