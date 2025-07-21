import os
import json
import datetime
import webbrowser
import pyttsx3
import requests
import sounddevice as sd
import tkinter as tk
from threading import Thread
from vosk import Model, KaldiRecognizer
import screen_brightness_control as sbc
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import warnings
import pythoncom
from tkinter import PhotoImage, scrolledtext

warnings.filterwarnings("ignore")

vosk_model = Model("vosk-model-small-en-us-0.15")

stop_signal = False
OPENROUTER_API_KEY = "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"

def speak(text):
    pythoncom.CoInitialize()
    engine = pyttsx3.init(driverName='sapi5')
    engine.setProperty('rate', 170)
    voices = engine.getProperty('voices')
    for voice in voices:
        if "female" in voice.name.lower() or "zira" in voice.id.lower():
            engine.setProperty('voice', voice.id)
            break
    print("Tany:", text)
    chat_box.insert(tk.END, f"Tany: {text}\n")
    engine.say(text)
    engine.runAndWait()
    pythoncom.CoUninitialize()

def update_status(message, color="gray"):
    status_label.config(text=message, fg=color)
    root.update_idletasks()

def get_audio():
    update_status("Listening...", "orange")
    rec = KaldiRecognizer(vosk_model, 16000)
    duration = 4
    def callback(indata, frames, time, status):
        rec.AcceptWaveform(bytes(indata))
    with sd.RawInputStream(samplerate=16000, blocksize=8000, dtype='int16',
                           channels=1, callback=callback):
        sd.sleep(duration * 1000)
    result = json.loads(rec.FinalResult())
    update_status("Processing...", "blue")
    return result.get("text", "").lower()

def change_volume(amount):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    current = volume.GetMasterVolumeLevelScalar()
    volume.SetMasterVolumeLevelScalar(min(max(current + amount, 0.0), 1.0), None)

def mute_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    volume.SetMute(1, None)

def ask_chatgpt(prompt):
    update_status("Thinking...", "purple")
    try:
        headers = {
            "Authorization": f"Bearer {OPENROUTER_API_KEY}",
            "Content-Type": "application/json"
        }
        data = {
            "model": "openai/gpt-3.5-turbo",
            "messages": [
                {"role": "system", "content": "You are a helpful voice assistant."},
                {"role": "user", "content": prompt}
            ]
        }
        res = requests.post("https://openrouter.ai/api/v1/chat/completions", headers=headers, json=data)
        if res.status_code == 200:
            return res.json()["choices"][0]["message"]["content"].strip()
        else:
            return "Sorry, I couldn't connect to AI services."
    except Exception as e:
        print("GPT Error:", e)
        return "Something went wrong while contacting the AI."

def handle(command):
    if "time" in command:
        return f"The time is {datetime.datetime.now().strftime('%I:%M %p')}"
    elif "date" in command:
        return f"Today is {datetime.datetime.now().strftime('%A, %B %d, %Y')}"
    elif "youtube" in command:
        webbrowser.open("https://youtube.com")
        return "Opening YouTube."
    elif "google" in command:
        webbrowser.open("https://google.com")
        return "Opening Google."
    elif "notepad" in command:
        os.system("notepad")
        return "Opening Notepad."
    elif "shutdown" in command:
        speak("Are you sure you want to shut down?")
        if "yes" in get_audio():
            os.system("shutdown /s /t 1")
            return "Shutting down."
        return "Shutdown canceled."
    elif "restart" in command:
        speak("Should I restart your PC?")
        if "yes" in get_audio():
            os.system("shutdown /r /t 1")
            return "Restarting now."
        return "Restart canceled."
    elif "brightness up" in command:
        try:
            current = sbc.get_brightness()[0]
            new_level = min(current + 20, 100)
            sbc.set_brightness(new_level)
            return f"Increasing brightness to {new_level}%"
        except:
            return "Couldn't change brightness."
    elif "brightness down" in command:
        try:
            current = sbc.get_brightness()[0]
            new_level = max(current - 20, 0)
            sbc.set_brightness(new_level)
            return f"Decreasing brightness to {new_level}%"
        except:
            return "Couldn't change brightness."
    elif "volume up" in command:
        change_volume(0.1)
        return "Increasing volume."
    elif "volume down" in command:
        change_volume(-0.1)
        return "Decreasing volume."
    elif "mute" in command:
        mute_volume()
        return "Volume muted."
    elif "battery" in command:
        try:
            from psutil import sensors_battery
            battery = sensors_battery()
            return f"Battery is at {battery.percent}% and {'charging' if battery.power_plugged else 'not charging'}."
        except:
            return "Battery info not available."
    elif "stop" in command or "goodbye" in command:
        return "Goodbye!"
    return None

def main():
    global stop_signal
    speak("Hello, I am Tany. You can start speaking now.")
    update_status("Listening...", "#666666")
    while not stop_signal:
        command = get_audio()
        if not command:
            continue
        chat_box.insert(tk.END, f"You: {command}\n")
        response = handle(command)
        if not response:
            response = ask_chatgpt(command)
        speak(response)
        if "goodbye" in response or "stop" in command:
            break

def run_tany():
    main()

def start_tany():
    global stop_signal
    stop_signal = False
    t = Thread(target=run_tany)
    t.daemon = True
    t.start()

def stop_tany():
    global stop_signal
    stop_signal = True
    update_status("Stopped Listening", "red")

def quit_app():
    global stop_signal
    stop_signal = True
    update_status("Exiting...", "red")
    root.after(1000, root.destroy)

root = tk.Tk()
root.title("Tany AI Assistant")
root.geometry("500x500")
root.configure(bg="#F5F7FA")

header = tk.Label(root, text="Ask Tany. \nShe’s already listening.", font=("Segoe UI", 18, "bold"), fg="#5F79D5", bg="#F5F7FA")
header.pack(pady=(30, 10))

waveform = tk.Canvas(root, width=440, height=70, bg="#F5F7FA", bd=0, highlightthickness=0)
waveform.pack()
waveform.create_line(10, 35, 430, 35, fill="#9C5EFF", width=4, smooth=True)

chat_box = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=60, height=10, font=("Segoe UI", 11))
chat_box.pack(pady=10, padx=20)
chat_box.configure(state='normal', bg="white", fg="black")

entry_frame = tk.Frame(root, bg="#E8EDF2", bd=0)
entry_frame.pack(pady=10, fill="x", padx=30)
entry = tk.Entry(entry_frame, font=("Segoe UI", 13), bd=0, bg="#FFFFFF", fg="#000000")
entry.pack(side="left", fill="x", expand=True, padx=(10, 0), pady=10)

mic_btn = tk.Button(entry_frame, text="🎙️", font=("Segoe UI", 14), bg="#7F3FBF", fg="white", bd=0, command=start_tany)
mic_btn.pack(side="right", padx=10)

control_frame = tk.Frame(root, bg="#F5F7FA")
control_frame.pack(pady=10)


status_label = tk.Label(root, text="Waiting to start...", font=("Segoe UI", 10), fg="#6C6F73", bg="#F5F7FA")
status_label.pack(pady=(5, 15))

root.mainloop()