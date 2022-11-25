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

        # コメント入力欄の位置を指定
        pyautogui.click(x=840, y=730, interval=0.5, button="left")
        pyautogui.hotkey("ctrl", "v")
        pyautogui.hotkey("enter")
        time.sleep(5)

        # 最新の会話の位置を指定
        pyautogui.click(x=862, y=658, interval=0.5, button="right")
        # 右クリック後に出る「コピー」の位置を指定
        pyautogui.click(x=882, y=681, interval=0.5, button="left")
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
