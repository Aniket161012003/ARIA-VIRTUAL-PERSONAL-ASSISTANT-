import pyttsx3
import speech_recognition as sr
from datetime import datetime
import webbrowser
import os
import random
import pyautogui
import time
import subprocess
import smtplib
import pyperclip
import win32com.client as win32
import email
import imaplib
import ssl
import re
import math
import wikipedia
import wolframalpha
import requests
import json

class IntegratedAdvancedAssistant:
    def __init__(self):
        # User and Bot Configuration
        self.USERNAME = os.getenv('USERNAME', 'User')
        self.BOTNAME = "ARIA"  # Keeping ARIA name from the second implementation

        # Initialize Text-to-Speech Engine
        self.engine = pyttsx3.init('sapi5')
        self.engine.setProperty('rate', 180)
        self.engine.setProperty('volume', 1.0)

        # Set Voice (Female voice)
        voices = self.engine.getProperty('voices')
        self.engine.setProperty('voice', voices[1].id)

        # Speech Recognition Setup
        self.recognizer = sr.Recognizer()

        # External API Setup (Optional - you'll need to get these)
        self.wolframalpha_app_id = 'YOUR_WOLFRAMALPHA_APP_ID'
        self.openweather_api_key = 'YOUR_OPENWEATHER_API_KEY'

        # Email Configuration (Optional)
        self.EMAIL_ADDRESS = "your_email@gmail.com"
        self.EMAIL_PASSWORD = "your_email_password"

        # Knowledge and Context Stores
        self.user_context = {
            'name': self.USERNAME,
            'interests': [],
            'preferences': {}
        }
        self.conversation_history = []

    def speak(self, text):
        """Convert text to speech with print output and emotional nuance."""
        try:
            print(f"{self.BOTNAME}: {text}")
            # Optional: Add pauses for more natural speech
            text = text.replace('.', '. ')
            self.engine.say(text)
            self.engine.runAndWait()
        except Exception as e:
            print(f"Speech error: {e}")

    def listen(self):
        """Advanced speech recognition method with multiple fallback mechanisms."""
        with sr.Microphone() as source:
            print("Listening...")
            self.recognizer.adjust_for_ambient_noise(source, duration=1)
            self.recognizer.pause_threshold = 1
            
            try:
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=5)
                
                try:
                    query = self.recognizer.recognize_google(audio, language='en-in')
                    print(f"You said: {query}")
                    return query.lower()
                except sr.UnknownValueError:
                    self.speak("Sorry, I couldn't understand that. Could you repeat?")
                except sr.RequestError:
                    self.speak("Speech recognition service is unavailable.")
            except Exception as e:
                self.speak(f"Listening error: {e}")
        
        return "none"

    def greet_user(self):
        """Contextual greeting based on time of day."""
        hour = datetime.now().hour
        
        if 6 <= hour < 12:
            greeting = f"Good morning, {self.USERNAME}"
        elif 12 <= hour < 16:
            greeting = f"Good afternoon, {self.USERNAME}"
        elif 16 <= hour < 19:
            greeting = f"Good evening, {self.USERNAME}"
        else:
            greeting = f"Hello, {self.USERNAME}"
        
        self.speak(f"{greeting}! I'm {self.BOTNAME}, your intelligent desktop assistant.")
        self.speak("How can I help you today?")

    def open_application(self, app_name):
        """Enhanced application opening with error handling."""
        try:
            if app_name in ['notepad', 'calculator', 'cmd']:
                app_map = {
                    'notepad': 'notepad.exe',
                    'calculator': 'calc.exe',
                    'cmd': 'cmd.exe'
                }
                subprocess.Popen(app_map[app_name])
                self.speak(f"Opening {app_name}")
            else:
                # Try to open any application using Windows search
                pyautogui.hotkey('win')
                time.sleep(0.5)
                pyautogui.typewrite(app_name)
                time.sleep(0.5)
                pyautogui.press('enter')
                self.speak(f"Attempting to open {app_name}")
        except Exception as e:
            self.speak(f"Could not open {app_name}. Error: {e}")

    def close_application(self, app_name):
        """Enhanced application closing with process termination."""
        try:
            app_map = {
                'notepad': 'notepad.exe',
                'calculator': 'calculator.exe',
                'cmd': 'cmd.exe',
                'chrome': 'chrome.exe',
                'firefox': 'firefox.exe'
            }
            
            if app_name in app_map:
                os.system(f'taskkill /F /IM {app_map[app_name]}')
                self.speak(f"Closed {app_name}")
            else:
                self.speak(f"Cannot close {app_name}. Application not recognized.")
        except Exception as e:
            self.speak(f"Error closing {app_name}: {e}")

    def search_web(self, query, site="google"):
        """Advanced web search functionality."""
        search_term = query.replace('search', '').strip()
        
        site_urls = {
            'google': f"https://www.google.com/search?q={search_term}",
            'youtube': f"https://www.youtube.com/results?search_query={search_term}",
            'wikipedia': f"https://en.wikipedia.org/wiki/{search_term.replace(' ', '_')}"
        }
        
        try:
            webbrowser.open(site_urls.get(site, site_urls['google']))
            self.speak(f"Searching {site} for {search_term}")
        except Exception as e:
            self.speak(f"Could not perform web search: {e}")

    def voice_typing_in_notepad(self):
        """Advanced voice typing with more control."""
        self.open_application("notepad")
        time.sleep(1)
        
        self.speak("Voice typing started. Say 'stop typing' to end.")
        
        while True:
            command = self.listen()
            
            if command == 'none':
                continue
            
            if 'stop typing' in command:
                self.speak("Stopping voice typing.")
                break
            
            if command != 'none':
                # Enhanced typing with punctuation and capitalization
                formatted_text = command.capitalize()
                pyautogui.typewrite(formatted_text + ' ')
                pyautogui.press('enter')

    def calculator_by_speech(self):
        """
        Opens the Windows calculator and performs calculations based on voice input.
        Uses speech recognition to capture the expression and interact with the calculator.
        """
        try:
            # Open the calculator application
            self.speak("Opening calculator for you")
            self.open_application("calculator")
            time.sleep(2)  # Allow calculator to fully load
            
            self.speak("Please speak the mathematical expression you want to calculate")
            expression = self.listen()
            
            if expression == 'none':
                self.speak("Sorry, I couldn't hear the expression. Please try again.")
                return
                
            # Clean up the expression by removing words like "calculate" or "what is"
            clean_expression = expression.lower()
            remove_words = ["calculate", "what is", "what's", "compute", "solve"]
            for word in remove_words:
                clean_expression = clean_expression.replace(word, "").strip()
                
            self.speak(f"Computing {clean_expression}")
            
            # Convert spoken numbers and operators to calculator inputs
            # Map spoken operators to keyboard inputs
            operator_map = {
                "plus": "+",
                "add": "+",
                "minus": "-",
                "subtract": "-",
                "times": "*",
                "multiply by": "*",
                "multiplied by": "*",
                "divided by": "/",
                "divide by": "/",
            }
            
            # Process the expression
            for spoken_op, key_op in operator_map.items():
                clean_expression = clean_expression.replace(spoken_op, key_op)
            
            # Enter the expression into calculator
            for char in clean_expression:
                if char == ' ':
                    continue
                elif char == '*':
                    pyautogui.press('*')
                elif char == '/':
                    pyautogui.press('/')
                elif char == '+':
                    pyautogui.press('+')
                elif char == '-':
                    pyautogui.press('-')
                elif char == '.':
                    pyautogui.press('.')
                elif char == '(':
                    pyautogui.press('(')
                elif char == ')':
                    pyautogui.press(')')
                else:
                    pyautogui.press(char)
                    
            # Press enter to get the result
            pyautogui.press('enter')
            
            # Allow time for the calculation to complete
            time.sleep(1)
            
            # Copy result to clipboard (Ctrl+C)
            pyautogui.hotkey('ctrl', 'c')
            result = pyperclip.paste()
            
            self.speak(f"The result is {result}")
            
            # Ask if the user wants to perform another calculation
            self.speak("Would you like to perform another calculation?")
            response = self.listen()
            
            if response == 'none':
                # Close calculator if no response
                self.close_application("calculator")
                return
                
            if any(word in response.lower() for word in ["yes", "yeah", "sure", "okay"]):
                # Clear calculator for next calculation
                pyautogui.press('escape')
                self.calculator_by_speech()  # Recursive call for another calculation
            else:
                # Close calculator
                self.close_application("calculator")
                self.speak("Calculator closed")
                
        except Exception as e:
            self.speak(f"An error occurred while using the calculator: {e}")
            # Try to close calculator in case of error
            try:
                self.close_application("calculator")
            except:
                pass

    def send_whatsapp_message(self):
        """Send WhatsApp message via web."""
        try:
            self.speak("Who do you want to send a message to?")
            contact = self.listen()
            
            if contact == 'none':
                return
            
            self.speak(f"What message do you want to send to {contact}?")
            message = self.listen()
            
            if message == 'none':
                return
            
            # Open WhatsApp Web
            webbrowser.open('https://web.whatsapp.com/')
            self.speak("Please ensure you're logged into WhatsApp Web")
            time.sleep(5)  # Wait for WhatsApp Web to load
            
            # Note: This requires manual interaction currently
            self.speak(f"Please manually select {contact} and send: {message}")
        
        except Exception as e:
            self.speak(f"Could not send WhatsApp message: {e}")

    def send_email(self):
        """Send email via browser."""
        try:
            self.speak("Who do you want to send an email to?")
            recipient = self.listen()
            
            if recipient == 'none':
                return
            
            self.speak("What is the subject of the email?")
            subject = self.listen()
            
            if subject == 'none':
                return
            
            self.speak("What message do you want to send?")
            message = self.listen()
            
            if message == 'none':
                return
            
            # Open Gmail in browser
            webbrowser.open('https://mail.google.com/mail/u/0/#compose')
            time.sleep(3)
            
            self.speak("Please complete the email sending process manually.")
            
        except Exception as e:
            self.speak(f"Could not send email: {e}")

    def system_control(self, command):
        """System control commands."""
        try:
            if 'shutdown' in command:
                os.system('shutdown /s /t 1')
            elif 'restart' in command:
                os.system('shutdown /r /t 1')
            elif 'sleep' in command:
                os.system('rundll32.exe powrprof.dll,SetSuspendState 0,1,0')
            elif 'lock' in command:
                os.system('rundll32.exe user32.dll,LockWorkStation')
        except Exception as e:
            self.speak(f"System control error: {e}")

    def problem_solver(self, query):
        """Advanced problem-solving method with multiple strategies."""
        # Mathematical Problem Solving
        if re.search(r'\b(calculate|solve|math|equation)\b', query):
            try:
                # Extract mathematical expression
                match = re.search(r'([\d+\-*/\(\) ]+)', query)
                if match:
                    expression = match.group(1)
                    result = eval(expression)
                    response = f"The solution is {result}"
                    return response
            except Exception as e:
                return f"I couldn't solve the mathematical problem: {e}"

        # Historical and Biographical Information
        historical_keywords = ['who is', 'tell me about', 'biography of', 'history of']
        if any(keyword in query.lower() for keyword in historical_keywords):
            try:
                # Extract search term by removing keywords
                search_term = query.lower()
                for keyword in historical_keywords:
                    search_term = search_term.replace(keyword, '').strip()
                
                try:
                    # Try Wikipedia first
                    summary = wikipedia.summary(search_term, sentences=3)
                    return summary
                except wikipedia.exceptions.DisambiguationError as e:
                    # If disambiguation, try the first option
                    if e.options:
                        try:
                            summary = wikipedia.summary(e.options[0], sentences=3)
                            return f"Based on {e.options[0]}: {summary}"
                        except:
                            return f"Multiple results found. Try being more specific: {e.options[:3]}"
                except wikipedia.exceptions.PageError:
                    return f"Sorry, I couldn't find detailed information about {search_term}."
            except Exception as e:
                return f"Could not retrieve information: {e}"

        # Wolframalpha Advanced Computation (Fallback)
        try:
            client = wolframalpha.Client(self.wolframalpha_app_id)
            res = client.query(query)
            answer = next(res.results).text
            return answer
        except Exception:
            pass

        # Decision Making Logic
        if re.search(r'\b(should I|what to do|decide)\b', query):
            decisions = [
                "Based on the information, I recommend...",
                "After careful consideration, the best option seems to be...",
                "Weighing the pros and cons, I suggest..."
            ]
            return random.choice(decisions)

        # General Knowledge Fallback
        try:
            # Try Wikipedia with a broader search
            search_term = query.strip()
            summary = wikipedia.summary(search_term, sentences=2)
            return summary
        except:
            pass

        return "I'm sorry, I couldn't find a precise answer. Could you rephrase or provide more context?"

    def get_weather(self, city):
        """Retrieve weather information."""
        try:
            base_url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={self.openweather_api_key}&units=metric"
            response = requests.get(base_url)
            data = response.json()
            
            temp = data['main']['temp']
            description = data['weather'][0]['description']
            return f"Current weather in {city}: {temp}°C, {description}"
        except Exception as e:
            return f"Could not fetch weather: {e}"

    def advanced_conversation(self, query):
        """Contextual and intelligent conversation handler."""
        # Conversation Patterns
        patterns = {
            r'\b(hi|hello|hey)\b': [
                "Hello! I'm ARIA, your intelligent assistant. How can I help you today?",
                "Hi there! What would you like to discuss?",
                "Greetings! I'm ready to assist you with anything."
            ],
            r'\b(how are you)\b': [
                "I'm functioning perfectly and ready to help!",
                "Doing great! My systems are all optimized and ready to go.",
                "I'm in top form, thanks for asking!"
            ],
            r'\b(your name)\b': [
                f"I'm {self.BOTNAME}, which stands for Advanced Responsive Intelligent Assistant.",
                "My name is ARIA, an AI designed to be helpful and engaging."
            ]
        }

        # Check predefined patterns
        for pattern, responses in patterns.items():
            if re.search(pattern, query, re.IGNORECASE):
                return random.choice(responses)

        # Fallback to problem solver
        return self.problem_solver(query)

    def main_loop(self):
        """Main interaction loop with comprehensive command handling."""
        self.greet_user()

        while True:
            query = self.listen()

            if query == 'none':
                continue

            # Application Control
            if 'open' in query:
                app_name = query.replace('open', '').strip()
                self.open_application(app_name)

            elif 'close' in query:
                app_name = query.replace('close', '').strip()
                self.close_application(app_name)

            # Web Search
            elif 'search on google' in query:
                self.speak("What do you want to search?")
                search_query = self.listen()
                self.search_web(search_query, 'google')

            elif 'play on youtube' in query:
                self.speak("What do you want to watch?")
                video_query = self.listen()
                self.search_web(video_query, 'youtube')

            # Special Functions
            elif 'start typing in notepad' in query:
                self.voice_typing_in_notepad()

            # New Calculator Function
            elif any(calc_cmd in query for calc_cmd in ['use calculator', 'open calculator', 'calculate with calculator']):
                self.calculator_by_speech()

            elif 'send whatsapp message' in query:
                self.send_whatsapp_message()

            elif 'send email' in query:
                self.send_email()

            # Weather Query
            elif 'weather' in query:
                city = query.replace('weather', '').strip()
                response = self.get_weather(city)
                self.speak(response)

            # System Control
            elif any(cmd in query for cmd in ['shutdown', 'restart', 'sleep', 'lock']):
                self.system_control(query)

            # Advanced Conversation
            elif 'none' not in query:
                try:
                    response = self.advanced_conversation(query)
                    self.speak(response)
                except Exception as e:
                    self.speak(f"Sorry, I encountered an unexpected error: {e}")

            # Exit Commands
            elif any(exit_word in query for exit_word in ['exit', 'bye', 'goodbye', 'stop']):
                self.speak(f"Goodbye, {self.USERNAME}! Always here if you need me.")
                break

            else:
                self.speak("Sorry, I didn't understand that command.")

def main():
    assistant = IntegratedAdvancedAssistant()
    assistant.main_loop()

if __name__ == "__main__":
    main()