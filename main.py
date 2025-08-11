# jarvis.py
"""
Jarvis - Integrated voice assistant (updated)
- Adds confirmation for weather and email to reduce mis-recognition mistakes.
- Keeps website shortcuts, music playback (musiclibrary.music), news (NewsAPI)
- Uses Gmail API (OAuth) to send emails (no SMTP/App Password)
- Summarization via Gemini (if configured)
- Added get_verified_email() and send_email_with_confirmation() with live transcription and GUI confirmation.
"""

import os
import time
import webbrowser
import subprocess
import platform
import difflib
import requests
import pickle
import base64
import re
from email.mime.text import MIMEText

import speech_recognition as sr
import pyttsx3

# Optional: Porcupine wake-word (kept but fallback used if not available)
try:
    import pvporcupine
    from pvporcupine import Porcupine
    from pvrecorder import PvRecorder
    PICOVOICE_AVAILABLE = True
except Exception:
    PICOVOICE_AVAILABLE = False

# Gmail API libs
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# Gemini (Google Generative AI)
try:
    import google.generativeai as genai
    GENAI_AVAILABLE = True
except Exception:
    GENAI_AVAILABLE = False

# musiclibrary (local)
try:
    import musiclibrary
    MUSIC_LIB_AVAILABLE = hasattr(musiclibrary, "music")
except Exception:
    MUSIC_LIB_AVAILABLE = False

# Try to enable GUI confirmation (Tkinter)
try:
    import tkinter as tk
    from tkinter import messagebox
    TK_AVAILABLE = True
except Exception:
    TK_AVAILABLE = False

# Load .env if present
from dotenv import load_dotenv
load_dotenv()



# -------------------- Speech / TTS --------------------
recognizer = sr.Recognizer()
engine = pyttsx3.init()

def speak(text):
    print(f"Jarvis: {text}")
    try:
        engine.say(text)
        engine.runAndWait()
    except Exception as e:
        print(f"[TTS Error] {e}")

# -------------------- Gmail API (OAuth) --------------------
def authenticate_gmail():
    creds = None
    if os.path.exists(GMAIL_TOKEN_FILE):
        with open(GMAIL_TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)
    if not creds or not getattr(creds, "valid", False):
        if creds and getattr(creds, "expired", False) and getattr(creds, "refresh_token", None):
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"[Gmail] Failed to refresh token: {e}")
                creds = None
        if not creds:
            if not os.path.exists(GMAIL_CREDENTIALS_FILE):
                speak("Gmail credentials.json not found. Place your OAuth client secrets file as credentials.json.")
                raise FileNotFoundError("credentials.json not found")
            flow = InstalledAppFlow.from_client_secrets_file(GMAIL_CREDENTIALS_FILE, GMAIL_SCOPES)
            creds = flow.run_local_server(port=0)
        with open(GMAIL_TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)
    return creds

def create_message(sender, to, subject, message_text):
    msg = MIMEText(message_text)
    msg['to'] = to
    msg['from'] = sender
    msg['subject'] = subject
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return {'raw': raw}

def send_email_via_gmail(to, subject, body):
    try:
        creds = authenticate_gmail()
        service = build('gmail', 'v1', credentials=creds)
        message = create_message("me", to, subject, body)
        sent = service.users().messages().send(userId="me", body=message).execute()
        speak("Email sent successfully.")
        print(f"[Gmail] Message ID: {sent.get('id')}")
        return True
    except Exception as e:
        print(f"[Gmail Error] {e}")
        speak("I couldn't send the email.")
        return False

# -------------------- News / Weather / Summarize --------------------
NEWS_CACHE = {"time": 0, "articles": []}
NEWS_CACHE_TTL = 3600  # seconds

def fetch_news(category="technology", country="us"):
    global NEWS_CACHE
    now = time.time()
    if NEWS_CACHE["articles"] and (now - NEWS_CACHE["time"] < NEWS_CACHE_TTL):
        return NEWS_CACHE["articles"]
    if not NEWS_API_KEY:
        speak("News API key not configured.")
        return []
    try:
        url = f"https://newsapi.org/v2/top-headlines?category={category}&country={country}&apiKey={NEWS_API_KEY}"
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        data = r.json()
        articles = data.get("articles", [])
        NEWS_CACHE = {"time": now, "articles": articles}
        return articles
    except Exception as e:
        print(f"[News Error] {e}")
        return NEWS_CACHE.get("articles", [])

def get_weather(city):
    if not WEATHER_API_KEY:
        speak("Weather API key is not configured.")
        return
    try:
        url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={WEATHER_API_KEY}&units=metric"
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        data = r.json()
        if data.get("cod") != 200:
            speak(f"I couldn't find weather for {city}.")
            return
        temp = data["main"]["temp"]
        desc = data["weather"][0]["description"]
        speak(f"The temperature in {city} is {temp}°C with {desc}.")
    except Exception as e:
        print(f"[Weather Error] {e}")
        speak("I couldn't fetch the weather right now.")

def summarize_url_with_gemini(url):
    if not GENAI_AVAILABLE:
        speak("AI summarization is not available.")
        return
    # ensure scheme
    if not re.match(r"^https?://", url):
        url = "https://" + url
    try:
        speak("Fetching the page and summarizing. This may take a moment.")
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        content = r.text[:150000]  # truncate large pages
        prompt = f"Summarize the following webpage in 5 concise bullet points:\n\n{content}"
        model = genai.GenerativeModel("gemini-1.5-flash-latest")
        response = model.generate_content(prompt)
        summary = response.text
        speak("Here is the summary.")
        speak(summary)
    except Exception as e:
        print(f"[Summarizer Error] {e}")
        speak("I couldn't summarize that URL.")

# -------------------- Music & File utilities --------------------
def play_song_by_name(query):
    if not MUSIC_LIB_AVAILABLE:
        speak("Music library not available.")
        return
    keys = list(musiclibrary.music.keys())
    match = difflib.get_close_matches(query, keys, n=1, cutoff=0.5)
    if not match:
        speak("Song not found in your music library.")
        return
    name = match[0]
    link = musiclibrary.music[name]
    if link.startswith("http://") or link.startswith("https://"):
        webbrowser.open(link)
    else:
        open_file_cross_platform(link)
    speak(f"Playing {name}")

def open_file_cross_platform(path):
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(path)
        elif system == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
        speak(f"Opening {os.path.basename(path)}")
    except Exception as e:
        print(f"[Open file error] {e}")
        speak("I couldn't open that file.")

def search_and_open_file(filename, search_path=None):
    if search_path is None:
        search_path = os.path.expanduser("~")
    filename_lower = filename.lower().strip()
    for root, dirs, files in os.walk(search_path):
        for f in files:
            if filename_lower == f.lower() or filename_lower in f.lower():
                full = os.path.join(root, f)
                open_file_cross_platform(full)
                return True
    return False

# -------------------- Helpers for robust voice input --------------------
def listen_for_phrase(timeout=6, phrase_time_limit=6):
    try:
        with sr.Microphone() as source:
            recognizer.pause_threshold = 0.8
            recognizer.adjust_for_ambient_noise(source, duration=0.7)
            print("[Listening] ...")
            audio = recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
        text = recognizer.recognize_google(audio)
        print(f"[Heard] {text}")
        return text
    except sr.WaitTimeoutError:
        print("[Listen] timeout")
        return None
    except sr.UnknownValueError:
        print("[Listen] unknown")
        return None
    except sr.RequestError as e:
        print(f"[Listen] request error: {e}")
        speak("Speech recognition service is unavailable.")
        return None
    except Exception as e:
        print(f"[Listen] unexpected: {e}")
        return None

def ask_yes_no(prompt, timeout=5):
    """Ask a yes/no question and return True for yes, False for no, None for unknown."""
    speak(prompt)
    ans = listen_for_phrase(timeout=timeout, phrase_time_limit=3)
    if not ans:
        return None
    ans = ans.lower()
    if "yes" in ans or "yeah" in ans or "ya" in ans:
        return True
    if "no" in ans or "nah" in ans:
        return False
    return None

def sanitize_speech_email(text):
    """Convert common spoken email forms to actual email: ' at '->'@', ' dot '->'.'"""
    if not text:
        return text
    t = text.lower()
    t = t.replace(" at ", "@").replace(" dot ", ".").replace(" underscore ", "_").replace(" dash ", "-")
    t = t.replace(" space ", "")
    t = t.replace(" ", "")  # remove remaining spaces
    return t

def looks_like_email(address):
    if not address:
        return False
    # basic validation
    return re.match(r"[^@]+@[^@]+\.[^@]+", address) is not None

# -------------------- NEW: get_verified_email + GUI confirm --------------------
def gui_confirm(prompt_text):
    """If tkinter available, show a Yes/No dialog. Returns True/False or None if tkinter unavailable."""
    if not TK_AVAILABLE:
        return None
    try:
        root = tk.Tk()
        root.withdraw()  # hide main window
        answer = messagebox.askyesno("Jarvis confirmation", prompt_text)
        root.destroy()
        return answer
    except Exception as e:
        print(f"[GUI confirm error] {e}")
        return None

def get_verified_email():
    """
    Ask the user to spell the recipient email letter-by-letter.
    Say 'dot' for '.', 'at' for '@', and 'done' when finished.
    Show live transcription in terminal and offer GUI confirmation if available.
    Returns validated email string or None if cancelled.
    """
    while True:
        speak("Please spell the recipient's email address, letter by letter. Say 'dot' for dot, 'at' for at, and say 'done' when finished.")
        spelled = []
        print("\n[Spelling session started] Type 'done' verbally when finished.")
        while True:
            speak("Say the next character, or say 'done' if finished.")
            part = listen_for_phrase(timeout=8, phrase_time_limit=5)
            if not part:
                speak("I didn't catch that character. Please repeat the character or say 'done'.")
                continue
            p = part.lower().strip()
            print(f"[Spell heard] '{p}'")

            if p in ("done", "finish", "that's all", "that is all"):
                break

            # handle common mappings
            if p in ("dot", "period", "point"):
                spelled.append(".")
                speak("Added dot.")
            elif p in ("at", "atto", "add at"):
                spelled.append("@")
                speak("Added at.")
            elif p in ("underscore", "under score", "under-score"):
                spelled.append("_")
                speak("Added underscore.")
            elif p in ("dash", "hyphen", "minus"):
                spelled.append("-")
                speak("Added dash.")
            elif len(p) == 1:
                # single letter or number recognized
                spelled.append(p)
                speak(f"Added {p}")
            else:
                # multi-letter chunk like 'gmail' or 'com'
                spelled.append(p)
                speak(f"Added {p}")

            # live transcription in terminal
            current = "".join(spelled).replace(" ", "")
            print(f"[Current spelled email] {current}")

        email = "".join(spelled).replace(" ", "")
        if not email:
            speak("You didn't spell any characters. Canceling.")
            return None

        # Ask for confirmation (GUI first, then voice)
        confirm_text = f"You spelled: {email}. Is this correct?"
        gui_ans = gui_confirm(confirm_text)
        if gui_ans is True:
            if looks_like_email(email):
                return email
            else:
                speak("That doesn't look like a valid email address. Do you want to spell again? Say yes to retry or no to cancel.")
                retry = listen_for_phrase(timeout=6, phrase_time_limit=3)
                if retry and "yes" in retry.lower():
                    continue
                else:
                    return None
        elif gui_ans is False:
            speak("Okay, let's try again.")
            continue
        else:
            # no GUI available — fallback to voice confirmation
            speak(confirm_text + " Please say yes or no.")
            confirm = listen_for_phrase(timeout=6, phrase_time_limit=3)
            if confirm and "yes" in confirm.lower():
                if looks_like_email(email):
                    return email
                else:
                    speak("That doesn't look like a valid email address. Do you want to spell again? Say yes to retry or no to cancel.")
                    retry = listen_for_phrase(timeout=6, phrase_time_limit=3)
                    if retry and "yes" in retry.lower():
                        continue
                    else:
                        return None
            else:
                speak("Okay, let's try again.")
                continue

def send_email_with_confirmation():
    """
    Uses get_verified_email() to get recipient, then asks for subject and body, confirms,
    then calls send_email_via_gmail(). This function can be used in place of the older email flow.
    """
    recipient = get_verified_email()
    if not recipient:
        speak("Recipient not confirmed. Cancelling email.")
        return

    speak("What is the subject?")
    subject = listen_for_phrase(timeout=10, phrase_time_limit=8) or "No subject"

    speak("What should I say in the email?")
    body = listen_for_phrase(timeout=30, phrase_time_limit=25) or ""

    # Final confirmation (GUI preferred)
    final_text = f"Send email to: {recipient}\nSubject: {subject}\nMessage: {body}\n\nSend now?"
    gui_ans = gui_confirm(final_text)
    if gui_ans is True:
        send_ok = True
    elif gui_ans is False:
        send_ok = False
    else:
        # fallback to voice
        speak("Are you sure you want to send this email? Please say yes or no.")
        final = listen_for_phrase(timeout=8, phrase_time_limit=4)
        send_ok = (final and "yes" in final.lower())

    if send_ok:
        success = send_email_via_gmail(recipient, subject, body)
        if success:
            speak("Email sent.")
    else:
        speak("Okay, I will not send the email.")

# -------------------- Command processing --------------------
def process_command(c):
    if not c:
        return
    c_lower = c.lower()
    print(f"[Command] {c_lower}")

    # Websites
    sites = {
        "google": "https://google.com",
        "spotify": "https://open.spotify.com",
        "youtube": "https://youtube.com",
        "facebook": "https://facebook.com",
        "instagram": "https://instagram.com",
        "linkedin": "https://linkedin.com"
    }
    for name, url in sites.items():
        if f"open {name}" in c_lower:
            webbrowser.open(url)
            speak(f"Opening {name}")
            return

    # Play music
    if c_lower.startswith("play "):
        song = c_lower[len("play "):].strip()
        if song:
            play_song_by_name(song)
        else:
            speak("Please tell me the song name.")
        return

    # News
    if "news" in c_lower:
        articles = fetch_news()
        if not articles:
            speak("No news available right now.")
            return
        speak("Here are the top headlines.")
        for i, a in enumerate(articles[:5], 1):
            title = a.get("title", "No title")
            print(f"{i}. {title}")
            speak(title)
            time.sleep(0.4)
        return

    # Weather with confirmation
    if "weather" in c_lower:
        speak("Which city do you want the weather for?")
        city = listen_for_phrase(timeout=6, phrase_time_limit=6)
        if not city:
            speak("I didn't get the city name.")
            return
        # confirm
        ok = ask_yes_no(f"Did you say {city}? Please say yes or no.")
        if ok is False:
            # retry once
            speak("Okay, please say the city again.")
            city = listen_for_phrase(timeout=6, phrase_time_limit=6)
            if not city:
                speak("I still didn't catch it. Try using the weather command again later.")
                return
            ok2 = ask_yes_no(f"Did you say {city}?")
            if ok2 is not True:
                speak("Okay I won't fetch the weather.")
                return
        elif ok is None:
            # unclear, proceed but warn
            speak(f"I will try with {city}.")
        get_weather(city)
        return

    # Send email with robust confirmation and validation
    if "send email" in c_lower or "send an email" in c_lower:
        # Use the spelling + confirm flow
        send_email_with_confirmation()
        return

    # File open
    if "open file" in c_lower or "search file" in c_lower:
        speak("Tell me the filename with extension.")
        filename = listen_for_phrase(timeout=8, phrase_time_limit=8)
        if not filename:
            speak("I didn't get the filename.")
            return
        found = search_and_open_file(filename)
        if not found:
            speak("File not found in your home directory. Do you want me to search entire drive? Say yes or no.")
            ans = listen_for_phrase(timeout=5, phrase_time_limit=3)
            if ans and "yes" in ans.lower():
                search_and_open_file(filename, search_path="C:/")
        return

    # Summarize
    if "summarize" in c_lower or "summarise" in c_lower:
        speak("Please tell me the URL to summarize.")
        url = listen_for_phrase(timeout=10, phrase_time_limit=12)
        if url:
            summarize_url_with_gemini(url)
        else:
            speak("I didn't get the URL.")
        return

    # Fallback Gemini
    if GENAI_AVAILABLE:
        try:
            speak("Let me think...")
            model = genai.GenerativeModel("gemini-1.5-flash-latest")
            response = model.generate_content(c)
            reply = response.text
            speak(reply)
        except Exception as e:
            print(f"[Gemini Error] {e}")
            speak("I couldn't reach the AI right now.")
    else:
        speak("I did not understand that. I can open websites, play music, fetch news, give weather, send emails, open files, and summarize URLs.")

# -------------------- Wake loop (fallback mode) --------------------
def fallback_wake_loop():
    speak("Listening for wake word. Say 'Jarvis' to activate.")
    while True:
        try:
            with sr.Microphone() as source:
                recognizer.adjust_for_ambient_noise(source, duration=0.7)
                print("[Fallback] listening for wake word...")
                audio = recognizer.listen(source, timeout=6, phrase_time_limit=3)
            try:
                word = recognizer.recognize_google(audio)
                print(f"[Fallback heard] {word}")
                if "jarvis" in word.lower():
                    speak("Yes?")
                    command = listen_for_phrase(timeout=6, phrase_time_limit=10)
                    if command:
                        process_command(command)
                    else:
                        speak("I didn't hear a command.")
            except sr.UnknownValueError:
                continue
            except sr.RequestError as e:
                print(f"[Fallback request error] {e}")
                time.sleep(1)
        except sr.WaitTimeoutError:
            continue
        except KeyboardInterrupt:
            break
        except Exception as e:
            print(f"[Fallback unexpected] {e}")
            time.sleep(1)

# -------------------- Main --------------------
def main():
    speak("Initializing Jarvis...")
    # For now we will use fallback wake loop (speechrecognition). If you want Porcupine,
    # set PICOVOICE_AVAILABLE and PICOVOICE_ACCESS_KEY appropriately.
    if PICOVOICE_AVAILABLE and PICOVOICE_ACCESS_KEY:
        try:
            # If you want to use Porcupine, you can call porcupine_wake_loop here (not provided in this file)
            # For safety and simplicity, we fallback to speech-recognition based wake loop.
            fallback_wake_loop()
        except Exception:
            print("[Main] Porcupine unavailable, switching to fallback.")
            fallback_wake_loop()
    else:
        fallback_wake_loop()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("Shutting down Jarvis.")
    except Exception as e:
        print(f"[Main Error] {e}")
