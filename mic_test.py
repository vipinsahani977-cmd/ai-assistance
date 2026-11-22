import os
import sys
import speech_recognition as sr
import pyttsx3
import pythoncom

print("=" * 60)
print("       AI VOICE ASSISTANT - HARDWARE & AUDIO DIAGNOSTIC")
print("=" * 60)

# 1. Text-To-Speech Diagnostic
print("\n[Step 1/3] Testing Text-To-Speech Engine (pyttsx3 / SAPI5)...")
try:
    pythoncom.CoInitialize()
    engine = pyttsx3.init()
    engine.setProperty("rate", 175)
    engine.setProperty("volume", 1.0)
    voices = engine.getProperty("voices")
    print(f" -> Found {len(voices)} SAPI voice(s):")
    for i, v in enumerate(voices):
        print(f"    [{i}] {v.name} (ID: {v.id})")
    
    # Try speaking short phrase
    print(" -> Playing audio test through speakers...")
    engine.say("Audio test in progress. Text to speech is working correctly.")
    engine.runAndWait()
    print(" [PASS] Text-to-Speech output verified successfully.")
except Exception as e:
    print(f" [FAIL] TTS initialization/playback error: {e}")
finally:
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass

# 2. Microphone Enumeration
print("\n[Step 2/3] Detecting Input Audio Devices...")
try:
    mics = sr.Microphone.list_microphone_names()
    if not mics:
        print(" [WARNING] No microphones detected by PyAudio.")
    else:
        print(f" -> Detected {len(mics)} audio device(s):")
        for idx, name in enumerate(mics):
            print(f"    Device [{idx}]: {name}")
except Exception as e:
    print(f" [FAIL] Could not query audio devices: {e}")

# 3. Live Microphone Capture & Speech Recognition
print("\n[Step 3/3] Testing Live Audio Capture & Speech Recognition...")
recognizer = sr.Recognizer()
recognizer.dynamic_energy_threshold = True

try:
    with sr.Microphone() as source:
        print(" -> Calibrating background noise (1 second)...")
        recognizer.adjust_for_ambient_noise(source, duration=1.0)
        print(f" -> Current ambient energy threshold: {recognizer.energy_threshold:.1f}")
        print(" -> [MIC ACTIVE] Please speak clearly now (e.g., 'Hello Assistant' or 'Kya haal hai')...")
        audio = recognizer.listen(source, timeout=8, phrase_time_limit=6)

    print(" -> Audio captured! Sending to speech recognition service...")
    try:
        text = recognizer.recognize_google(audio, language="en-IN")
        print(f"\n [PASS] Recognized Speech: \"{text}\"")
    except sr.UnknownValueError:
        print("\n [NOTICE] Voice was recorded, but speech could not be recognized clearly.")
    except sr.RequestError as e:
        print(f"\n [NOTICE] Speech recognition service error (check internet): {e}")

except sr.WaitTimeoutError:
    print("\n [NOTICE] No audio detected within 8 seconds. Please check your mic volume / mute switch.")
except OSError as e:
    print(f"\n [FAIL] Microphone hardware access error: {e}")
except Exception as e:
    print(f"\n [FAIL] Unexpected audio test error: {e}")

print("\n" + "=" * 60)
print("Diagnostic Complete.")
print("=" * 60)
