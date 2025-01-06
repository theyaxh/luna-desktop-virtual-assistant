import random

class EntertainmentManager:
    def __init__(self):
        self._load_content_data()

    def _load_content_data(self):
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

    def get_random_fact(self):
        return random.choice(self.facts)

    def get_random_joke(self):
        return random.choice(self.jokes)

    def get_random_quote(self):
        return random.choice(self.quotes)