import os
import speech_recognition
import win32com.client
import webbrowser

def say(text):
    # Use the SAPI.SpVoice COM object to speak on Windows
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

def takeCommand():
    r = speech_recognition.Recognizer()
    with speech_recognition.Microphone() as source:
        r.pause_threshold = 1
        print("Listening...")
        audio = r.listen(source)

        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query.lower()  # Convert to lowercase for case-insensitive comparisons
        except speech_recognition.UnknownValueError:
            print("Sorry, I did not understand. Please repeat.")
            return ""
        except speech_recognition.RequestError as e:
            print(f"Could not request results from Google Speech Recognition service; {e}")
            return ""

if __name__ == "__main__":
    print('Pycharm')
    say("Hello, I am Jarvis A.I.")

    while True:
        command = takeCommand()

        if "open youtube" in command:
            webbrowser.open("https://www.youtube.com")
        elif "search google" in command:
            say("What do you want to search on Google?")
            search_query = takeCommand()
            if search_query:
                webbrowser.open(f"https://www.google.com/search?q={search_query.replace(' ', '+')}")
        elif "search wikipedia" in command:
            say("What do you want to search on Wikipedia?")
            search_query = takeCommand()
            if search_query:
                webbrowser.open(f"https://en.wikipedia.org/wiki/{search_query.replace(' ', '_')}")
        elif "search spotify" in command:
            say("What song or artist do you want to search on Spotify?")
            search_query = takeCommand()
            if search_query:
                webbrowser.open(f"https://open.spotify.com/{search_query.replace(' ', '%20')}")
        elif "search jiosaavn" in command:
            say("What song or artist do you want to search on JioSaavn?")
            search_query = takeCommand()
            if search_query:
                webbrowser.open(f"https://www.jiosaavn.com/{search_query.replace(' ', '-')}")
        elif "exit" in command:
            say("Goodbye!")
            break
        else:
            say("I'm sorry, I didn't understand that command. Please repeat.")
