import speech_recognition as sr
import win32com.client as wincl
import pyautogui
import time
import pyperclip

r = sr.Recognizer()
mic = sr.Microphone()

while True:
    print("Say something ...")

    with mic as source:
        r.adjust_for_ambient_noise(source) #雑音対策
        audio = r.listen(source)

    print ("Now to recognize it...")

    try:
        recogn = r.recognize_google(audio, language='ja-JP')
        print(recogn)

        # "さようなら" と言ったら音声認識を止める
        if r.recognize_google(audio, language='ja-JP') == "さようなら" :
            print("end")
            break

        pyperclip.copy(recogn)

        pyautogui.click(x=1050, y=680, interval=0.5, button="left")
        pyautogui.hotkey("ctrl", "v")
        pyautogui.hotkey("enter")
        time.sleep(5)

        pyautogui.click(x=1110, y=615, interval=0.5, button="right")
        pyautogui.click(x=1120, y=625, interval=0.5, button="left")
        time.sleep(1)
        clip_str = pyperclip.paste()
        print(clip_str)

        voice = wincl.Dispatch("SAPI.SpVoice")
        voice.Rate = 0  # [-10 to 10]
        voice.Speak(clip_str)


    # 以下は認識できなかったときに止まらないように。
    except sr.UnknownValueError:
        print("could not understand audio")
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; {0}".format(e))
