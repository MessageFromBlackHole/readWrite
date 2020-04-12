import win32com.client as wincl
speak = wincl.Dispatch("SAPI.SpVoice")
speak.Speak("I am a Black Hole. I need to tell you something important.")