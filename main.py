import os
import re
import sys
import time
import queue
import datetime
import threading
import subprocess
import urllib.parse
import webbrowser
from typing import Optional, List, Dict, Any

# Windows COM initialization for multi-threading
try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None

import speech_recognition as sr
import pyttsx3
from dotenv import load_dotenv

# Google GenAI official SDK
try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None
    types = None

# CustomTkinter GUI framework
try:
    import customtkinter as ctk
except ImportError:
    ctk = None

# Load environment variables
load_dotenv()

# Configuration from environment variables
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
AI_MODEL_ENV = os.getenv("AI_MODEL")
ASSISTANT_NAME = os.getenv("ASSISTANT_NAME", "Jarvis")

# Prioritized ultra-fast Gemini models
MODELS_TO_TRY = [
    AI_MODEL_ENV if AI_MODEL_ENV else "gemini-3.5-flash-lite",
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
    "gemini-flash-lite-latest",
    "gemini-3.5-flash"
]
# Remove duplicates preserving order
MODELS_TO_TRY = list(dict.fromkeys([m for m in MODELS_TO_TRY if m]))

SYSTEM_INSTRUCTION = (
    f"You are {ASSISTANT_NAME}, an intelligent, friendly and efficient personal AI voice assistant. "
    "Respond in the same language the user speaks. "
    "When answering in Hindi or Indian context, write your response using Latin/Roman script "
    "(Hinglish, e.g. 'Main badhiya hoon, aap bataiye kaise hain') "
    "because the computer's speech engine uses English voices (Microsoft Zira / David) "
    "which can vocalize Hinglish cleanly but cannot pronounce Devanagari Hindi characters. "
    "If the user asks in English, reply in natural English. "
    "Keep responses short, clear, and conversational (1 to 2 sentences maximum) so they can be spoken quickly via text-to-speech. "
    "Never use markdown formatting like asterisks, bullet points, headers, tables, or emojis."
)


# ============================================================================
# HELPER: DEVANAGARI TO PHONETIC HINGLISH TRANSLITERATOR
# ============================================================================
def transliterate_devanagari_if_needed(text: str) -> str:
    """
    If text contains Devanagari Hindi characters and the active SAPI voice is an
    English-only voice (David / Zira), transliterate phonetically to Latin script
    so the speech engine does not silently skip Devanagari characters.
    """
    if not text:
        return ""
    # Check if text contains Devanagari Unicode range (0900 to 097F)
    has_devanagari = any('\u0900' <= char <= '\u097f' for char in text)
    if not has_devanagari:
        return text

    mapping = {
        'अ': 'a', 'आ': 'aa', 'इ': 'i', 'ई': 'ee', 'उ': 'u', 'ऊ': 'oo', 'ऋ': 'ri',
        'ए': 'e', 'ऐ': 'ai', 'ओ': 'o', 'औ': 'au', 'क': 'k', 'ख': 'kh', 'ग': 'g',
        'घ': 'gh', 'ङ': 'ng', 'च': 'ch', 'छ': 'chh', 'ज': 'j', 'झ': 'jh', 'ञ': 'ny',
        'ट': 't', 'ठ': 'th', 'ड': 'd', 'ढ': 'dh', 'ण': 'n', 'त': 't', 'थ': 'th',
        'द': 'd', 'ध': 'dh', 'न': 'n', 'प': 'p', 'फ': 'ph', 'ब': 'b', 'भ': 'bh',
        'म': 'm', 'य': 'y', 'र': 'r', 'ल': 'l', 'व': 'v', 'श': 'sh', 'ष': 'sh',
        'स': 's', 'ह': 'h', 'ा': 'aa', 'ि': 'i', 'ी': 'ee', 'ु': 'u', 'ू': 'oo',
        'ृ': 'ri', 'े': 'e', 'ै': 'ai', 'ो': 'o', 'ौ': 'au', 'ं': 'n', 'ँ': 'n',
        'ः': 'h', '्': '', '।': '.'
    }
    consonants = 'कखगघङचछजझञटठडढणतथदधनपफबभमयरलवशषसह'
    matras = 'ािीुूृेैोौ्'

    res = []
    i = 0
    while i < len(text):
        char = text[i]
        if char in mapping:
            trans = mapping[char]
            if char in consonants:
                next_char = text[i+1] if i + 1 < len(text) else ''
                if next_char not in matras and next_char != '्':
                    trans += 'a'
            res.append(trans)
        else:
            res.append(char)
        i += 1
    return ''.join(res)


# ============================================================================
# 1. TEXT-TO-SPEECH (TTS) MANAGER (Thread-Safe, Windows SAPI5 & pyttsx3)
# ============================================================================
class TTSManager:
    """
    Production-grade text-to-speech manager for Windows.
    Features:
    - Primary output via pyttsx3 with SAPI5 engine.
    - Automatic native Windows SAPI fallback to guarantee audible output.
    - Non-blocking background worker thread (GUI never freezes).
    - Sentence chunking for responsive interruption via stop().
    - Mute/unmute, speech rate, volume, and voice selection controls.
    - Markdown symbol stripping and Devanagari-to-Hinglish phonetics.
    - Comprehensive diagnostic logs for tracing speech lifecycle.
    """

    def __init__(self, on_status_change=None):
        self.queue = queue.Queue()
        self.stop_event = threading.Event()
        self.ready_event = threading.Event()
        self.muted = False
        self.is_speaking = False
        self.current_voice_id = None
        self.current_voice_name = "Default"
        self.cached_voices: List[Dict[str, str]] = []
        self.has_hindi_voice = False
        self.rate = 175
        self.volume = 1.0
        self.on_status_change = on_status_change
        self._engine = None

        # Start background speech worker thread
        self.worker_thread = threading.Thread(target=self._speech_worker, daemon=True)
        self.worker_thread.start()
        # Wait for initial voice enumeration
        self.ready_event.wait(timeout=2.0)

    def _init_engine(self):
        """Initialize pyttsx3 engine using Windows SAPI5 driver."""
        try:
            engine = pyttsx3.init("sapi5")
            engine.setProperty("rate", self.rate)
            engine.setProperty("volume", self.volume)
            if self.current_voice_id:
                try:
                    engine.setProperty("voice", self.current_voice_id)
                except Exception:
                    pass
            return engine
        except Exception as e:
            print(f"[TTS Init Warning]: pyttsx3.init('sapi5') failed: {e}")
            return None

    def _speech_worker(self):
        """Dedicated worker thread handling SAPI5 COM lifetime."""
        if pythoncom:
            pythoncom.CoInitialize()

        try:
            self._engine = self._init_engine()
            if self._engine:
                try:
                    voices = self._engine.getProperty("voices")
                    for v in voices:
                        self.cached_voices.append({"id": v.id, "name": v.name})
                        # Check if any Hindi SAPI voice is present (language ID 0439 / 1081)
                        if "hindi" in v.name.lower() or "kalpana" in v.name.lower() or "hemant" in v.name.lower():
                            self.has_hindi_voice = True
                        if "zira" in v.name.lower() and not self.current_voice_id:
                            self.current_voice_id = v.id
                            self.current_voice_name = v.name

                    if not self.current_voice_id and voices:
                        self.current_voice_id = voices[0].id
                        self.current_voice_name = voices[0].name
                except Exception as e:
                    print(f"[TTS Voices Warning]: {e}")

            # Native SAPI fallback enumeration if needed
            if not self.cached_voices and win32com:
                try:
                    sp = win32com.client.Dispatch("SAPI.SpVoice")
                    for token in sp.GetVoices():
                        desc = token.GetDescription()
                        self.cached_voices.append({"id": token.Id, "name": desc})
                        if "zira" in desc.lower() and not self.current_voice_id:
                            self.current_voice_id = token.Id
                            self.current_voice_name = desc
                    if not self.current_voice_id and self.cached_voices:
                        self.current_voice_id = self.cached_voices[0]["id"]
                        self.current_voice_name = self.cached_voices[0]["name"]
                except Exception as e:
                    print(f"[TTS Native SAPI Warning]: {e}")

        finally:
            self.ready_event.set()

        print(f"[TTS Engine Ready] Loaded {len(self.cached_voices)} Windows SAPI voice(s). Selected Voice: '{self.current_voice_name}'")
        if not self.has_hindi_voice:
            print("[TTS Language Note] No native Windows Hindi SAPI voice found (David and Zira are English US). Devanagari Hindi text will be voiced phonetically in Hinglish.")

        while True:
            item = self.queue.get()
            if item is None:
                break

            text, voice_id = item

            if self.muted:
                print(f"[TTS Skipped]: Voice is OFF (Muted). Displaying text only.")
                self.queue.task_done()
                continue

            if not text or not str(text).strip():
                self.queue.task_done()
                continue

            # 1. Strip markdown symbols, code blocks, URLs
            clean_text = self._clean_for_speech(text)
            if not clean_text:
                self.queue.task_done()
                continue

            # 2. If active voice is English and text contains Devanagari, transliterate to Hinglish
            if not self.has_hindi_voice and any('\u0900' <= c <= '\u097f' for c in clean_text):
                devanagari_sample = clean_text[:30]
                clean_text = transliterate_devanagari_if_needed(clean_text)
                print(f"[TTS Transliteration]: Converted '{devanagari_sample}...' -> '{clean_text[:35]}...' for {self.current_voice_name}")

            self.stop_event.clear()
            self.is_speaking = True
            if self.on_status_change:
                self.on_status_change(True)

            print(f"[TTS Speaking Start]: \"{clean_text[:65]}{'...' if len(clean_text) > 65 else ''}\"")
            spoken_successfully = False

            # Primary: pyttsx3 SAPI5 playback
            try:
                if not self._engine:
                    self._engine = self._init_engine()

                if self._engine:
                    # Reset internal loop state if interrupted previously
                    if getattr(self._engine, "_inLoop", False):
                        self._engine._inLoop = False

                    if voice_id:
                        try:
                            self._engine.setProperty("voice", voice_id)
                        except Exception:
                            pass
                    self._engine.setProperty("rate", self.rate)
                    self._engine.setProperty("volume", self.volume)

                    # Sentence chunking allows responsive stop()
                    sentences = re.split(r'(?<=[.!?\n]) +', clean_text)
                    for sentence in sentences:
                        if self.stop_event.is_set():
                            print("[TTS Stop]: Speech interrupted by user.")
                            break
                        chunk = sentence.strip()
                        if chunk:
                            self._engine.say(chunk)
                            self._engine.runAndWait()

                    spoken_successfully = not self.stop_event.is_set()

            except Exception as pyttsx_err:
                print(f"[TTS Notice]: pyttsx3 error '{pyttsx_err}'. Triggering native Windows SAPI fallback...")
                self._engine = None

            # Secondary: Native Windows SAPI.SpVoice fallback
            if not spoken_successfully and not self.stop_event.is_set() and win32com:
                try:
                    sp = win32com.client.Dispatch("SAPI.SpVoice")
                    sp.Volume = int(self.volume * 100)
                    sp.Rate = int(max(-10, min(10, (self.rate - 175) / 15)))
                    if voice_id:
                        for token in sp.GetVoices():
                            if token.Id == voice_id:
                                sp.Voice = token
                                break

                    sp.Speak(clean_text, 0)  # Synchronous playback on worker thread
                    spoken_successfully = True
                    print("[TTS Fallback]: Audio playback completed via native Windows SAPI.")
                except Exception as fallback_err:
                    print(f"[TTS Error]: Native SAPI playback failed: {fallback_err}")

            if spoken_successfully:
                print("[TTS Speaking Complete]: Audio playback finished successfully.")

            self.is_speaking = False
            if self.on_status_change:
                self.on_status_change(False)
            self.queue.task_done()

        if pythoncom:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    @staticmethod
    def _clean_for_speech(text: str) -> str:
        """Strip markdown symbols, URLs, asterisks, and code formatting for speech."""
        if not text:
            return ""
        # Remove markdown links [text](url) -> text
        text = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', text)
        # Remove URLs
        text = re.sub(r'https?://\S+', '', text)
        # Remove markdown format symbols (* _ ` # ~ >)
        text = re.sub(r'[*_`#~>]', '', text)
        # Remove list dashes/bullets
        text = re.sub(r'^[ \t]*[-•*][ \t]+', '', text, flags=re.MULTILINE)
        return re.sub(r'\s+', ' ', text).strip()

    def get_available_voices(self) -> List[Dict[str, str]]:
        """Return list of available system voices."""
        self.ready_event.wait(timeout=2.0)
        return self.cached_voices

    def set_voice(self, voice_id: str):
        """Set the active SAPI voice ID."""
        self.current_voice_id = voice_id
        for v in self.cached_voices:
            if v["id"] == voice_id:
                self.current_voice_name = v["name"]
                break
        print(f"[TTS Config]: Active voice changed to '{self.current_voice_name}'")

    def set_rate(self, rate: int):
        """Set speaking rate in words per minute (100 to 300)."""
        self.rate = max(100, min(300, int(rate)))
        if self._engine:
            try:
                self._engine.setProperty("rate", self.rate)
            except Exception:
                pass
        print(f"[TTS Config]: Speech speed set to {self.rate} WPM")

    def set_volume(self, volume: float):
        """Set volume level from 0.0 to 1.0."""
        self.volume = max(0.0, min(1.0, float(volume)))
        if self._engine:
            try:
                self._engine.setProperty("volume", self.volume)
            except Exception:
                pass
        print(f"[TTS Config]: Volume set to {int(self.volume * 100)}%")

    def speak_text(self, text: str):
        """Convert text into speech and play it through speakers asynchronously."""
        if not text or not str(text).strip():
            return
        if self.muted:
            print(f"[TTS Request]: Speech requested but Voice is OFF (Muted). Outputting text only.")
            return

        print(f"[TTS Request Enqueued]: \"{str(text)[:50]}{'...' if len(str(text)) > 50 else ''}\"")
        self.queue.put((str(text), self.current_voice_id))

    def speak(self, text: str):
        """Backward-compatible alias for speak_text."""
        self.speak_text(text)

    def stop(self):
        """Immediately stop currently playing audio and clear pending speech."""
        print("[TTS Stop]: Clearing speech queue and silencing audio.")
        self.stop_event.set()
        while not self.queue.empty():
            try:
                self.queue.get_nowait()
                self.queue.task_done()
            except queue.Empty:
                break

        if self._engine:
            try:
                self._engine.stop()
            except Exception:
                pass

        if win32com:
            try:
                sp = win32com.client.Dispatch("SAPI.SpVoice")
                sp.Speak("", 2)  # SVSFPurgeBeforeSpeak
            except Exception:
                pass

    def toggle_mute(self) -> bool:
        """Toggle mute state and return new state."""
        self.muted = not self.muted
        if self.muted:
            print("[TTS Voice State]: Voice output turned OFF.")
            self.stop()
        else:
            print("[TTS Voice State]: Voice output turned ON.")
        return self.muted

    def set_muted(self, muted: bool):
        """Set mute state explicitly."""
        self.muted = bool(muted)
        if self.muted:
            print("[TTS Voice State]: Voice output turned OFF.")
            self.stop()
        else:
            print("[TTS Voice State]: Voice output turned ON.")

    def shutdown(self):
        """Cleanly terminate TTS worker thread."""
        self.stop()
        self.queue.put(None)


# Global TTS manager instance for module-level speak_text
_global_tts_manager: Optional[TTSManager] = None

def speak_text(text: str) -> None:
    """Module-level speak_text function to convert any text to speech."""
    global _global_tts_manager
    if _global_tts_manager is None:
        _global_tts_manager = TTSManager()
    _global_tts_manager.speak_text(text)


# ============================================================================
# 2. AI INTELLIGENCE ENGINE (Official Google GenAI SDK & Multi-turn Chat)
# ============================================================================
class AIEngine:
    """Conversational AI engine with multi-turn memory and automatic model fallback."""

    def __init__(self):
        self.client: Optional[genai.Client] = None
        self.chat_session = None
        self.active_model: Optional[str] = None
        self.is_connected = False
        self.last_error: Optional[str] = None
        self._init_connection()

    def _init_connection(self) -> bool:
        """Connect to Google GenAI and create active chat session."""
        if not GEMINI_API_KEY:
            self.last_error = (
                "GEMINI_API_KEY is not configured in .env file. "
                "Please add: GEMINI_API_KEY=your_gemini_api_key"
            )
            self.is_connected = False
            return False

        if not genai:
            self.last_error = "google-genai SDK is not installed."
            self.is_connected = False
            return False

        try:
            self.client = genai.Client(api_key=GEMINI_API_KEY)
        except Exception as e:
            self.last_error = f"Failed to initialize GenAI client: {e}"
            self.is_connected = False
            return False

        for model in MODELS_TO_TRY:
            try:
                chat = self.client.chats.create(
                    model=model,
                    config=types.GenerateContentConfig(
                        system_instruction=SYSTEM_INSTRUCTION,
                        max_output_tokens=200,
                        temperature=0.7,
                    )
                )
                # Test connectivity with ping
                ping = chat.send_message("hi")
                if ping and ping.text:
                    self.chat_session = chat
                    self.active_model = model
                    self.is_connected = True
                    self.last_error = None
                    return True
            except Exception as e:
                print(f"[AI Notice] Model {model} probe failed: {e}")
                continue

        self.last_error = "Could not establish connection with any Gemini model. Check network & API quota."
        self.is_connected = False
        return False

    def reset_chat(self) -> bool:
        """Reset conversation session history."""
        return self._init_connection()

    def generate_response(self, user_prompt: str) -> str:
        """Send message to conversational session and return AI response."""
        if not self.is_connected or not self.chat_session:
            # Attempt re-connection
            if not self._init_connection():
                return self.last_error or "AI service is offline. Please check your API key."

        try:
            response = self.chat_session.send_message(user_prompt)
            if response and response.text:
                return response.text.strip()
            return "I received an empty response. Please ask again."

        except Exception as err:
            err_str = str(err)
            print(f"[AI Error with {self.active_model}]: {err_str}")

            # Fallback to alternate model if rate-limited or busy
            for model in MODELS_TO_TRY:
                if model == self.active_model:
                    continue
                try:
                    new_chat = self.client.chats.create(
                        model=model,
                        config=types.GenerateContentConfig(
                            system_instruction=SYSTEM_INSTRUCTION,
                            max_output_tokens=200,
                            temperature=0.7,
                        )
                    )
                    resp = new_chat.send_message(user_prompt)
                    if resp and resp.text:
                        self.chat_session = new_chat
                        self.active_model = model
                        return resp.text.strip()
                except Exception:
                    continue

            if "429" in err_str or "quota" in err_str.lower():
                return "API quota limit reached. Please check your Gemini API plan."
            if "503" in err_str or "unavailable" in err_str.lower():
                return "The AI server is experiencing high traffic. Please try again in a moment."

            return f"Sorry, I encountered an issue: {err_str[:90]}"


# ============================================================================
# 3. SAFE COMMAND PROCESSOR (Instant Desktop & Web Automation)
# ============================================================================
class CommandProcessor:
    """Safely executes whitelisted desktop and system actions."""

    @staticmethod
    def process(text: str, tts_manager: TTSManager) -> Optional[str]:
        raw = text.strip()
        cmd = raw.lower()

        # Stop speech command
        if cmd in ("stop speaking", "stop", "chup", "ruko", "quiet", "silence", "shant ho jao"):
            tts_manager.stop()
            return "Stopped speaking."

        # Identity
        if any(q in cmd for q in [
            "what is your name", "kya naam hai tumhara", "tumhara naam kya hai",
            "who are you", "aap kaun ho", "tell me about yourself"
        ]):
            return f"I am {ASSISTANT_NAME}, your personal AI voice assistant."

        # Current Time
        if any(q in cmd for q in [
            "what is the time", "current time", "what time is it",
            "time kya hai", "kitne baje", "kya time ho raha hai", "time batao"
        ]):
            now = datetime.datetime.now()
            return f"The current time is {now.strftime('%I:%M %p')}."

        # Current Date
        if any(q in cmd for q in [
            "what is today's date", "today's date", "what is the date",
            "aaj ki date", "aaj kya tarikh hai", "date batao"
        ]):
            now = datetime.datetime.now()
            return f"Today is {now.strftime('%A, %B %d, %Y')}."

        # Open YouTube / Play song
        if "youtube" in cmd and any(w in cmd for w in ["play", "chalao", "bajao", "search", "dhoondo"]):
            query = re.sub(r'^(?:play|search|youtube par|youtube pe|kholo|chalao|bajao)\s*', '', cmd)
            query = re.sub(r'(?:on youtube|pe chalao|par chalao|chalao|bajao|search karo)$', '', query).strip()
            query = query.replace("youtube", "").strip()
            if query:
                url = f"https://www.youtube.com/results?search_query={urllib.parse.quote(query)}"
                webbrowser.open(url)
                return f"Playing {query} on YouTube."
            webbrowser.open("https://www.youtube.com")
            return "Opening YouTube."

        # Open Websites
        if "open youtube" in cmd or "youtube kholo" in cmd:
            webbrowser.open("https://www.youtube.com")
            return "Opening YouTube."

        if "open google" in cmd or "google kholo" in cmd:
            webbrowser.open("https://www.google.com")
            return "Opening Google."

        if "open whatsapp" in cmd or "whatsapp kholo" in cmd:
            webbrowser.open("https://web.whatsapp.com")
            return "Opening WhatsApp Web."

        if "open github" in cmd or "github kholo" in cmd:
            webbrowser.open("https://github.com")
            return "Opening GitHub."

        # Google Search
        search_triggers = [
            "search google for", "google search", "search on google",
            "google par search karo", "google pe dhoondo", "search for"
        ]
        for trigger in search_triggers:
            if trigger in cmd:
                query = cmd.split(trigger, 1)[1].strip()
                if query:
                    url = f"https://www.google.com/search?q={urllib.parse.quote(query)}"
                    webbrowser.open(url)
                    return f"Searching Google for {query}."

        # Open Chrome safely
        if "open chrome" in cmd or "chrome kholo" in cmd:
            chrome_paths = [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                os.path.expanduser(r"~\AppData\Local\Google\Chrome\Application\chrome.exe")
            ]
            launched = False
            for path in chrome_paths:
                if os.path.exists(path):
                    subprocess.Popen([path])
                    launched = True
                    break
            if not launched:
                webbrowser.open("https://www.google.com")
            return "Opening Google Chrome."

        # Open Notepad safely
        if "open notepad" in cmd or "notepad kholo" in cmd:
            subprocess.Popen(["notepad.exe"])
            return "Opening Notepad."

        # Open Calculator safely
        if "open calculator" in cmd or "calculator kholo" in cmd:
            subprocess.Popen(["calc.exe"])
            return "Opening Calculator."

        # Greetings
        if cmd in ("hello", "hi", "hey", "namaste", "suno", "hello jarvis"):
            return f"Hello! How can I assist you today?"

        return None


# ============================================================================
# 4. VOICE LISTENER (Microphone & Speech Recognition)
# ============================================================================
class VoiceListener:
    """Manages microphone audio capture and speech-to-text recognition."""

    def __init__(self, on_status_change=None, on_speech_recognized=None):
        self.recognizer = sr.Recognizer()
        self.recognizer.dynamic_energy_threshold = True
        self.on_status_change = on_status_change
        self.on_speech_recognized = on_speech_recognized
        self.is_listening = False
        self.stop_listening_flag = threading.Event()
        self.language = "en-IN"  # en-IN or hi-IN

    def start_listening(self):
        """Start listening in a non-blocking daemon thread."""
        if self.is_listening:
            return
        self.is_listening = True
        self.stop_listening_flag.clear()
        threading.Thread(target=self._listen_loop, daemon=True).start()

    def stop_listening(self):
        """Signal listener loop to stop."""
        self.stop_listening_flag.set()
        self.is_listening = False
        if self.on_status_change:
            self.on_status_change("idle")

    def _listen_loop(self):
        """Listen loop on background thread."""
        try:
            mic = sr.Microphone()
        except Exception as e:
            if self.on_status_change:
                self.on_status_change(f"Mic Error: {e}")
            self.is_listening = False
            return

        try:
            if self.on_status_change:
                self.on_status_change("Calibrating background noise...")
            with mic as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=0.8)

            while not self.stop_listening_flag.is_set():
                if self.on_status_change:
                    self.on_status_change("listening")

                try:
                    with mic as source:
                        audio = self.recognizer.listen(
                            source,
                            timeout=6,
                            phrase_time_limit=10
                        )

                    if self.stop_listening_flag.is_set():
                        break

                    if self.on_status_change:
                        self.on_status_change("processing")

                    text = self.recognizer.recognize_google(audio, language=self.language)
                    if text and text.strip():
                        if self.on_speech_recognized:
                            self.on_speech_recognized(text.strip())
                        # After recognizing a phrase, stop one-shot listening
                        break

                except sr.WaitTimeoutError:
                    # Timeout with no speech detected; check if still listening
                    continue
                except sr.UnknownValueError:
                    if self.on_status_change:
                        self.on_status_change("Could not understand audio")
                    time.sleep(1)
                    continue
                except sr.RequestError as e:
                    if self.on_status_change:
                        self.on_status_change(f"Speech recognition network error: {e}")
                    break
                except OSError as e:
                    if self.on_status_change:
                        self.on_status_change(f"Microphone device error: {e}")
                    break

        except Exception as e:
            if self.on_status_change:
                self.on_status_change(f"Audio error: {e}")
        finally:
            self.is_listening = False
            if self.on_status_change:
                self.on_status_change("idle")


# ============================================================================
# 5. MODERN DESKTOP GUI (CustomTkinter)
# ============================================================================
class AssistantGUI:
    """Modern dark desktop GUI for AI Voice Assistant using CustomTkinter."""

    def __init__(self, root):
        self.root = root
        self.root.title(f"{ASSISTANT_NAME} - AI Voice Assistant")
        self.root.geometry("880x740")
        self.root.minsize(720, 620)

        # Set CustomTkinter theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Initialize Engines
        self.tts = TTSManager(on_status_change=self._on_tts_status)
        self.ai = AIEngine()
        self.listener = VoiceListener(
            on_status_change=self._on_listener_status,
            on_speech_recognized=self._on_speech_recognized
        )

        self._build_ui()

        # Handle window closing protocol
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Welcome message
        self.root.after(500, self._show_welcome)

    def _build_ui(self):
        """Construct the modern GUI layout."""
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

        # --- 1. HEADER FRAME ---
        header_frame = ctk.CTkFrame(self.root, corner_radius=0, fg_color="#181a20", height=65)
        header_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        header_frame.grid_columnconfigure(1, weight=1)

        # App Title & Icon
        title_label = ctk.CTkLabel(
            header_frame,
            text=f"⚡ {ASSISTANT_NAME.upper()} AI",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="#00d2ff"
        )
        title_label.grid(row=0, column=0, padx=20, pady=12, sticky="w")

        # Status Pill Indicator
        self.status_pill = ctk.CTkLabel(
            header_frame,
            text="🟢 Ready",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#242731",
            text_color="#2ed573",
            corner_radius=12,
            padx=14,
            pady=5
        )
        self.status_pill.grid(row=0, column=1, padx=10, pady=12, sticky="w")

        # Header Action Buttons
        btn_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        btn_frame.grid(row=0, column=2, padx=15, pady=10, sticky="e")

        # Clear, unambiguous Voice On / Voice Off button
        self.mute_btn = ctk.CTkButton(
            btn_frame,
            text="🔊 Voice: ON",
            width=105,
            height=32,
            fg_color="#1b9a59",
            hover_color="#168049",
            command=self.toggle_mute
        )
        self.mute_btn.pack(side="left", padx=5)

        self.stop_btn = ctk.CTkButton(
            btn_frame,
            text="⏹ Stop Speech",
            width=100,
            height=32,
            fg_color="#e63946",
            hover_color="#d62828",
            command=self.stop_speech
        )
        self.stop_btn.pack(side="left", padx=5)

        # --- 2. CHAT HISTORY FRAME (SCROLLABLE) ---
        self.chat_frame = ctk.CTkScrollableFrame(
            self.root,
            corner_radius=10,
            fg_color="#101216",
            scrollbar_button_color="#2b2d42",
            scrollbar_button_hover_color="#457b9d"
        )
        self.chat_frame.grid(row=1, column=0, sticky="nsew", padx=15, pady=(10, 5))
        self.chat_frame.grid_columnconfigure(0, weight=1)

        # --- 3. SETTINGS & VOICE CONTROLS BAR ---
        controls_frame = ctk.CTkFrame(self.root, fg_color="#181a20", corner_radius=8, height=48)
        controls_frame.grid(row=2, column=0, sticky="ew", padx=15, pady=(0, 8))

        # Language Selector
        lang_lbl = ctk.CTkLabel(controls_frame, text="Lang:", font=ctk.CTkFont(size=12))
        lang_lbl.pack(side="left", padx=(10, 3), pady=8)

        self.lang_menu = ctk.CTkOptionMenu(
            controls_frame,
            values=["English (en-IN)", "Hindi (hi-IN)"],
            width=135,
            height=28,
            command=self._on_lang_changed
        )
        self.lang_menu.set("English (en-IN)")
        self.lang_menu.pack(side="left", padx=3, pady=8)

        # Voice Selector
        voice_lbl = ctk.CTkLabel(controls_frame, text="Voice:", font=ctk.CTkFont(size=12))
        voice_lbl.pack(side="left", padx=(10, 3), pady=8)

        voices = self.tts.get_available_voices()
        voice_names = [v["name"] for v in voices] if voices else ["System Default"]
        self.voice_map = {v["name"]: v["id"] for v in voices}

        self.voice_menu = ctk.CTkOptionMenu(
            controls_frame,
            values=voice_names[:4],
            width=180,
            height=28,
            command=self._on_voice_changed
        )
        for v in voices:
            if v["id"] == self.tts.current_voice_id:
                self.voice_menu.set(v["name"])
                break
        self.voice_menu.pack(side="left", padx=3, pady=8)

        # Speech Rate Selector
        rate_lbl = ctk.CTkLabel(controls_frame, text="Speed:", font=ctk.CTkFont(size=12))
        rate_lbl.pack(side="left", padx=(10, 3), pady=8)

        self.rate_menu = ctk.CTkOptionMenu(
            controls_frame,
            values=["140 WPM (Slow)", "175 WPM (Normal)", "210 WPM (Fast)"],
            width=135,
            height=28,
            command=self._on_rate_changed
        )
        self.rate_menu.set("175 WPM (Normal)")
        self.rate_menu.pack(side="left", padx=3, pady=8)

        # Clear Chat Button
        self.clear_btn = ctk.CTkButton(
            controls_frame,
            text="🗑 Clear Chat",
            width=90,
            height=28,
            fg_color="#343a40",
            hover_color="#495057",
            command=self.clear_chat
        )
        self.clear_btn.pack(side="right", padx=10, pady=8)

        # --- 4. BOTTOM INPUT BAR ---
        input_frame = ctk.CTkFrame(self.root, fg_color="#181a20", corner_radius=10, height=65)
        input_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=(0, 15))
        input_frame.grid_columnconfigure(1, weight=1)

        # Microphone Toggle Button
        self.mic_btn = ctk.CTkButton(
            input_frame,
            text="🎤 Listen",
            width=110,
            height=42,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#00adb5",
            hover_color="#009688",
            command=self.toggle_listening
        )
        self.mic_btn.grid(row=0, column=0, padx=(12, 8), pady=12)

        # Text Input Box
        self.entry_box = ctk.CTkEntry(
            input_frame,
            placeholder_text="Type your question or command here... (Press Enter to Send)",
            height=42,
            font=ctk.CTkFont(size=14),
            border_color="#2b2d42",
            fg_color="#0f1115"
        )
        self.entry_box.grid(row=0, column=1, sticky="ew", padx=5, pady=12)
        self.entry_box.bind("<Return>", lambda event: self.send_text_message())

        # Send Button
        self.send_btn = ctk.CTkButton(
            input_frame,
            text="➤ Send",
            width=85,
            height=42,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#3a86ff",
            hover_color="#2563eb",
            command=self.send_text_message
        )
        self.send_btn.grid(row=0, column=2, padx=(8, 12), pady=12)

    def speak_text(self, text: str):
        """Play response through speakers automatically when unmuted."""
        if not text or not str(text).strip():
            return
        self.tts.speak_text(text)

    def _show_welcome(self):
        """Display initial greeting message and speak welcome audio."""
        status_text = (
            f"Hello! I am {ASSISTANT_NAME}, your AI Voice Assistant.\n"
            f"• AI Model: {'Google Gemini (' + str(self.ai.active_model) + ')' if self.ai.is_connected else 'Offline (Check API Key)'}\n"
            f"• TTS Status: SAPI5 Active ({len(self.tts.cached_voices)} voice(s) available)\n"
            "• You can type messages below or click '🎤 Listen' to speak.\n"
            "• Try saying: 'What time is it?', 'Open YouTube', 'Open Notepad', or any question."
        )
        self.add_message("Assistant", status_text)
        self.speak_text(f"Hello! I am {ASSISTANT_NAME}. Your assistant is ready.")

    def add_message(self, sender: str, text: str):
        """Append formatted message bubble to chat scrollable frame."""
        is_user = (sender.lower() == "user")

        msg_row = ctk.CTkFrame(self.chat_frame, fg_color="transparent")
        msg_row.pack(fill="x", padx=10, pady=6)

        bubble_color = "#1e3a5f" if is_user else "#21252b"
        border_color = "#3a86ff" if is_user else "#3e4451"
        align = "e" if is_user else "w"
        badge_color = "#00d2ff" if is_user else "#2ed573"

        bubble = ctk.CTkFrame(
            msg_row,
            fg_color=bubble_color,
            border_width=1,
            border_color=border_color,
            corner_radius=12
        )
        bubble.pack(anchor=align, padx=5, pady=2)

        header_lbl = ctk.CTkLabel(
            bubble,
            text=f"{'👤 You' if is_user else '⚡ ' + ASSISTANT_NAME}",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=badge_color
        )
        header_lbl.pack(anchor="w", padx=12, pady=(8, 2))

        content_lbl = ctk.CTkLabel(
            bubble,
            text=text,
            font=ctk.CTkFont(size=13),
            wraplength=600,
            justify="left",
            text_color="#f1f2f6"
        )
        content_lbl.pack(anchor="w", padx=12, pady=(2, 10))

        self.root.after(50, lambda: self.chat_frame._parent_canvas.yview_moveto(1.0))

    def _update_status(self, text: str, color: str = "#2ed573"):
        """Update top status pill safely on main thread."""
        self.status_pill.configure(text=text, text_color=color)

    def _on_listener_status(self, status: str):
        """Callback for speech listener status updates."""
        def update():
            if status == "listening":
                self._update_status("🟠 Listening... Speak now", "#ffa502")
                self.mic_btn.configure(text="⏹ Stop", fg_color="#ff4757", hover_color="#ee5253")
            elif status == "processing":
                self._update_status("🔵 Processing speech...", "#70a1ff")
            elif status == "idle":
                if not self.tts.is_speaking:
                    self._update_status("🟢 Ready", "#2ed573")
                self.mic_btn.configure(text="🎤 Listen", fg_color="#00adb5", hover_color="#009688")
            else:
                self._update_status(f"ℹ️ {status[:25]}", "#eccc68")
        self.root.after(0, update)

    def _on_tts_status(self, is_speaking: bool):
        """Callback for TTS speech engine state."""
        def update():
            if is_speaking:
                self._update_status("🟣 Speaking...", "#a55eea")
            elif not self.listener.is_listening:
                self._update_status("🟢 Ready", "#2ed573")
        self.root.after(0, update)

    def _on_speech_recognized(self, text: str):
        """Handles text received from microphone."""
        self.root.after(0, lambda: self.process_user_input(text))

    def _on_lang_changed(self, choice: str):
        """Switch speech recognition language."""
        if "Hindi" in choice:
            self.listener.language = "hi-IN"
        else:
            self.listener.language = "en-IN"

    def _on_voice_changed(self, choice: str):
        """Switch active TTS voice."""
        voice_id = self.voice_map.get(choice)
        if voice_id:
            self.tts.set_voice(voice_id)

    def _on_rate_changed(self, choice: str):
        """Adjust speech rate."""
        if "Slow" in choice:
            self.tts.set_rate(140)
        elif "Fast" in choice:
            self.tts.set_rate(210)
        else:
            self.tts.set_rate(175)

    def toggle_listening(self):
        """Toggle microphone listening state."""
        if self.listener.is_listening:
            self.listener.stop_listening()
        else:
            self.tts.stop()
            self.listener.start_listening()

    def toggle_mute(self):
        """Toggle TTS audio mute."""
        is_muted = self.tts.toggle_mute()
        if is_muted:
            self.mute_btn.configure(text="🔇 Voice: OFF", fg_color="#c0392b", hover_color="#962d22")
            print("[Voice Control]: Voice output is now OFF (Text-only mode).")
        else:
            self.mute_btn.configure(text="🔊 Voice: ON", fg_color="#1b9a59", hover_color="#168049")
            print("[Voice Control]: Voice output is now ON (Audible mode).")

    def stop_speech(self):
        """Stop TTS audio immediately."""
        self.tts.stop()
        self._update_status("🟢 Ready", "#2ed573")

    def clear_chat(self):
        """Clear message history and reset AI session."""
        for widget in self.chat_frame.winfo_children():
            widget.destroy()
        self.ai.reset_chat()
        self.add_message("Assistant", "Chat history and session memory have been cleared.")

    def send_text_message(self):
        """Handle send button click or Enter key press."""
        text = self.entry_box.get().strip()
        if not text:
            return
        self.entry_box.delete(0, "end")
        self.process_user_input(text)

    def process_user_input(self, text: str):
        """Coordinate user input processing across local commands and AI."""
        self.add_message("User", text)

        # Stop listening if active
        if self.listener.is_listening:
            self.listener.stop_listening()

        # Check for exit commands
        cmd_lower = text.lower().strip()
        if cmd_lower in ("exit", "quit", "close assistant", "bye", "band karo"):
            self.add_message("Assistant", "Goodbye! Have a wonderful day.")
            self.speak_text("Goodbye! Have a wonderful day.")
            self.root.after(1500, self.on_closing)
            return

        # Check instant local commands first
        local_reply = CommandProcessor.process(text, self.tts)
        if local_reply:
            self.add_message("Assistant", local_reply)
            self.speak_text(local_reply)
            return

        # Process with AI on background thread so UI never freezes
        self._update_status("🟣 AI Thinking...", "#3a86ff")

        def ai_worker():
            ai_response = self.ai.generate_response(text)
            self.root.after(0, lambda: self._on_ai_response_ready(ai_response))

        threading.Thread(target=ai_worker, daemon=True).start()

    def _on_ai_response_ready(self, response_text: str):
        """Display AI response in chat and AUTOMATICALLY speak it through speakers."""
        self.add_message("Assistant", response_text)
        self.speak_text(response_text)

    def on_closing(self):
        """Clean shutdown protocol."""
        try:
            self.listener.stop_listening()
            self.tts.shutdown()
        except Exception:
            pass
        self.root.destroy()
        sys.exit(0)


# ============================================================================
# 6. INTERACTIVE CLI MODE (Fallback for headless / terminal execution)
# ============================================================================
def run_cli_mode():
    """Terminal interactive mode for AI voice assistant."""
    print("=" * 60)
    print(f"      {ASSISTANT_NAME.upper()} AI VOICE ASSISTANT - CLI MODE")
    print("=" * 60)

    tts = TTSManager()
    ai = AIEngine()
    listener = VoiceListener()

    print(f"[*] AI Status: {'Connected (' + str(ai.active_model) + ')' if ai.is_connected else 'Offline: ' + str(ai.last_error)}")
    tts.speak_text(f"Hello! {ASSISTANT_NAME} is ready in command line mode.")

    print("\nInstructions:")
    print(" • Type your message and press Enter.")
    print(" • Type 'voice' or 'listen' to trigger microphone speech input.")
    print(" • Type 'mute' to toggle voice output.")
    print(" • Type 'clear' to reset chat memory.")
    print(" • Type 'exit' or 'quit' to close.\n")

    while True:
        try:
            user_input = input("\nYou > ").strip()
            if not user_input:
                continue

            if user_input.lower() in ("exit", "quit", "band karo"):
                print(f"{ASSISTANT_NAME} > Goodbye!")
                tts.speak_text("Goodbye!")
                break

            if user_input.lower() in ("clear", "reset"):
                ai.reset_chat()
                print(f"{ASSISTANT_NAME} > Session memory cleared.")
                continue

            if user_input.lower() in ("mute", "unmute"):
                is_muted = tts.toggle_mute()
                print(f"{ASSISTANT_NAME} > Voice {'OFF (Muted)' if is_muted else 'ON (Audible)'}.")
                continue

            if user_input.lower() in ("voice", "listen", "mic"):
                print("[*] Calibrating microphone & listening...")
                recognizer = sr.Recognizer()
                try:
                    with sr.Microphone() as source:
                        recognizer.adjust_for_ambient_noise(source, duration=0.8)
                        print("[*] Speak now...")
                        audio = recognizer.listen(source, timeout=6, phrase_time_limit=10)
                    user_input = recognizer.recognize_google(audio, language="en-IN")
                    print(f"You (Voice) > {user_input}")
                except Exception as e:
                    print(f"[Mic Error]: {e}")
                    continue

            # Check local command
            cmd_reply = CommandProcessor.process(user_input, tts)
            if cmd_reply:
                print(f"{ASSISTANT_NAME} > {cmd_reply}")
                tts.speak_text(cmd_reply)
                continue

            # Query AI
            print(f"{ASSISTANT_NAME} > Thinking...")
            reply = ai.generate_response(user_input)
            print(f"{ASSISTANT_NAME} > {reply}")
            tts.speak_text(reply)

        except KeyboardInterrupt:
            print("\nExiting...")
            break

    tts.shutdown()


# ============================================================================
# 7. MAIN ENTRY POINT
# ============================================================================
def main():
    """Main application entry point."""
    if "--cli" in sys.argv:
        run_cli_mode()
        return

    if ctk is None:
        print("[Notice] CustomTkinter is not installed. Launching in CLI mode...")
        run_cli_mode()
        return

    try:
        root = ctk.CTk()
        app = AssistantGUI(root)
        root.mainloop()
    except Exception as e:
        print(f"[GUI Launch Error]: {e}. Falling back to CLI mode...")
        run_cli_mode()


if __name__ == "__main__":
    main()
