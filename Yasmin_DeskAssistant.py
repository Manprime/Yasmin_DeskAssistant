import pyttsx3
import datetime
from duckduckgo_search import DDGS
import webbrowser
import speech_recognition as sr
import time
import os
import sys
import textwrap

class Desk():
    def __init__(self):
        self.talk=pyttsx3.init()
        self.analyse=sr.Recognizer()
        self.ip_type=None
    
    def speak(self,txt):
        out_voices = self.talk.getProperty('voices')
        self.talk.setProperty('voice', out_voices[1].id)
        self.talk.setProperty('rate', 200)
        self.talk.setProperty('volume',1.0) 
        self.talk.say(txt)
        self.talk.runAndWait()
        


    def analyseVoice(self):
        with sr.Microphone() as source:
            print("Listening...........")
            self.analyse.energy_threshold=300
            #acceptable input voice level
            audio=self.analyse.listen(source)
            
            try:
                print("Recognizing.....")
                word = self.analyse.recognize_google(audio , language='en-in').lower()
                print("User said:"+ word)
                return word
                
            except:
                tell="Sorry,not able to hear you."
                print(tell)
                self.speak(tell)
                time.sleep(1)
                #to wait for 1 secs 
                return "d"
                

                
            
        
    def input_mode(self):
        tell= ("Press 1 to activate 'voice mode' or continue in 'text mode' for inputs.")
        print(tell)
        self.speak(tell)
        self.i_type = input("Input: ")
        return self.ip_type  

    def wish(self):
        hr = int(datetime.datetime.now().hour)
        if hr >= 0 and hr <= 12:
            tell= "Good Morning!"
            print(tell)
            self.speak(tell)
        elif hr > 12 and hr < 16:
            tell= "Good Afternoon!"
            print(tell)
            self.speak(tell)
        else:
            tell = "Good Evening!"
            print(tell)
            self.speak(tell)
        tell= "Hi, I am Yasmin. I am here to help and assist you. Feel free to ask anything."
        print(tell)
        self.speak(tell)

    
    
    def ask_ques(self,type="0"):
        if type=="0":
            tell= "What may i help you with?"
            print(tell)
            self.speak(tell)
        else:
            pass
        if self.i_type == "1":
            ques = self.analyseVoice()
        else:
            ques = input("Input: ")
            print("Processing.....")
        return ques
    
    
    def ask_again(self):
        self.speak("To end assist type 'end' or 'stop' or 'end assist'.To continue give input as 'continue'")
        print("**To end assist type 'end' or 'stop' or 'end assist'.\nTo continue give input as 'continue'.")
        choice = input("Enter :")
        if choice == "end" or choice == "stop" or choice == 'end assist':
            say = "Thank you for letting me assist you."
            self.speak(say)
            print(say)
            if int(datetime.datetime.now().hour) < 18:
                say = "Have a nice day! Bye."
                print(say)
                self.speak(say)
            else:
                say = "Good Night! Have sweet dreams."
                print(say)
                self.speak(say)
            exit(1)
        else:
            say = "Please wait....Let's continue. "
            print(say)
            self.speak(say)
            self.basic_fun()
    
    def open_application(self, app):
        if "libre office" in app or "libreoffice" in app:
            self.speak("Opening libre office")
            os.startfile('"E:/Libre Office/program/soffice.exe"')

        elif "firefox" in app or "mozilla" in app:
            self.speak("Opening Mozilla Firefox")
            os.startfile("C:/Program Files/Mozilla Firefox/firefox.exe")

        elif "microsoft word" in app or "msword" in app:
            self.speak("Opening Microsoft Word")
            os.startfile("C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE")

        elif "wordpad" in app:
            self.speak("Opening wordpad")
            os.startfile( "C:/Program Files/Windows NT/Accessories/wordpad.exe")

        elif "writer" in app or "libre writer" in app or "libreoffice writer" in app:
            self.speak("Opening Microsoft Word")
            os.startfile("E:/Libre Office/program/swriter.exe")

        elif "impress" in app or "libreoffice impress" in app or "presentation" in app:
            self.speak("Opening Libre Office impress")
            os.startfile("E:/Libre Office/program/simpress.exe")

        elif "notepad" in app:
            self.speak("Opening notepad")
            os.startfile("C:/Program Files/Microsoft Office/root/Office16/ONENOTE.EXE")

        elif "libre calc" in app:
            self.speak("Opening Libre Office calc")
            os.startfile("E:/Libre Office/program/scalc.exe")

        elif "powerpoint" in app:
            self.speak("Opening Microsoft powerpoint")
            os.startfile("C:/Program Files/Microsoft Office/root/Office16/POWERPNT.EXE")

        else:
            self.speak("Application not available")
    
    def basic_fun(self):
        while True:
            q = self.ask_ques().lower()
            if "translate" in q:
                self.speak("sure")
                tell = "Translating text requires an internet connecion.\nPress '1' to continue. "
                self.speak(tell)
                print(tell)
                ch = input("\aInput: ")
                if ch == "1":
                    text = input("Input text to translate : ")
                    lang = input("Enter input language(e.g 'en' for english) :")
                    to_lang = input("Enter first two characters of lang to translate text (e.g 'or' for odia) \n: ")
                    try:
                        results = DDGS().translate(keywords=text, from_=lang, to=to_lang)
                        print(results)
                    except Exception as e:
                        tell = "Error...Either there is a problem while connecting internet or language not supported."
                        self.speak(tell)
                        print(tell)

            elif "who made you" in q or "who created you" in q:
                self.speak("My creator is Man-Prime. He is very friendly. Iam really happy that he created me. It is because of him I am able to interact with you.Well, Man-Prime is his Github Profile name . His real name is ....Sorry,I just got distracted.....")

            elif "open youtube" in q:
                say = 'Opening Youtube in browser......'
                print(say)
                self.speak(say)
                webbrowser.open_new_tab("https://youtube.com")

            elif "open browser" in q:
                self.speak("sure")
                webbrowser.open_new_tab("https://lite.duckduckgo.com/lite")
                say = 'opening browser....'
                print(say)
                self.speak(say)

            elif "end" in q or "exit" in q:
                self.ask_again()

            elif "open help" in q or "help menu" in q:
                say = "Welcome to help menu.\nMention your queries in simple terms and avoid long sentences.\nQuick Access terms:\nTranslate-> To Translate sentences\nPassword-> generate/check Password\nopen [item] -> to open specific tasks such as youtube,browser,help menu, news ,apps etc.Type 'open' before the task to get error free experience.\nAI Assist-> for generating text\nopen web->for web search\nAlways describe your queries in simple terms.\nInstant Mode-> Type 'instant' then enter your query to get quick results"
                print(say)
                self.speak(say)

            elif "who i am" in q:
                self.speak("If you talk then you are definitely a human or are you an AI bot who trying to mess around with me?\nIn any case, my creator has taught me to help others in need. So, I will assist you.As i donot have eyes,it is difficult to say but i can talk,so many believe me to be another human. Okay that seems like a long response so let's move on.")

            elif "how are you" in q or "are you good" in q or "how do you do" in q:
                say = "I am feeling lively. Thank you for asking. Your concern really means a lot to me. How are you doing? "
                print(say)
                self.speak(say)
                converse=self.ask_ques(1)
                if "bad" in converse or "not good" in converse or "miserable" in converse:
                    self.speak("hope, everything gets good")
                else:
                    tell="Nice"
                    self.speak(tell)
                self.speak("Let's move on.")

            elif 'open app' in q:
                say = "Which application do you want to open?"
                print(say)
                self.speak(say)
                app = self.ask_ques(1)
                self.open_application(app)

            elif 'news' in q:
                self.speak("Sure")
                say = "Which country specific news do you want to know?"
                print(say)
                self.speak(say)
                newsQuery = input("Enter category(e.g Global):")
                try:
                    news_results = DDGS().news(keywords=newsQuery, region="wt-wt", safesearch="off", timelimit="5", max_results=10)
                    for i in range(10):
                        print("URL:", news_results[i]['url'], "\nSource:", news_results[i]['source'])
                        print("Date:", news_results[i]['date'])
                        print("Title", news_results[i]['title'], "\nDescription:", news_results[i]['body'], "\n")
                except Exception as e:
                    say = "There was an unexpected exception."
                    print(say)
                    self.speak(say)

            elif 'instant mode' in q:
                info = "You have entered 'Instant Mode'. Ask questions and get rapid answers.\n*To exit instant mode give input as 'change mode'.\nLet's get started."
                print(info)
                self.speak(info)
                ip_type = self.input_mode()
                while (True):
                    try:
                        ques1 = self.ask_ques(0)
                        if (ques1 == "change mode"):
                            break
                        else:
                            results = DDGS().answers(keywords=ques1)
                            key_to_lookup = 'text'
                            value = results[0][key_to_lookup]
                            print(value)
                            self.speak(value)
                    except Exception as e:
                        say = "Existing instant mode due to some error Hope you understand."
                        self.speak(say)
                        print("Reason",e)
                        
                        break

            else:
                self.parameters(q)
    
    def parameters(self, ques):
        qf = 0
        device_query = ["device", "ip address", "model no.", "hardware details", "software details", "network details"]
        for i in device_query:
            if i in ques:
                say = "Sorry, I cannot answer device related queries in order to respect user privacy."
                print(say)
                self.speak(say)
                qf = 1
                break

        date_query = ["date", "today's date"]
        for i in date_query:
            if i in ques:
                say = "If you want to know today's date, it is "
                self.speak(say)
                self.speak(datetime.datetime.now().strftime("%m-%d-%Y"))
                print(say, "", datetime.datetime.now().strftime("%d-%m-%Y"), "\n")
                qf = 1
                break

        time_query = ["time", "time now"]
        for i in time_query:
            if i in ques:
                say = "If you want to know the time, it is "
                self.speak(say)
                self.speak(datetime.datetime.now().strftime("%H:%M:%S"))
                print(say, "", datetime.datetime.now().strftime("%H:%M:%S"), "\n")
                qf = 1
                break

        web_query = ["open web", "online search", "find website", "search website", "surf web", "web surf",'search',"web search"]
        for i in web_query:
            if i in ques:
                say = "Entering World Wide Web.....\nWelcome to the world of internet"
                print(say)
                self.speak(say)
                query = input("Search Query: ")
                results = DDGS().text(keywords=query, region='wt-wt', safesearch='moderate', timelimit=None, max_results=6)
                for i in range(5):
                    print("Title:",results[i]['title'],"\nSource:",results[i]['href'],"\nBody:",results[i]['body'])
                qf = 1
                break

        assistant_query = ["who are you", "tell about yourself", 'your name', "name", "you", "hi", "hey"]
        for i in assistant_query:
            if i in ques:
                say = "Hi,I am Yelena, your friendly desktop assistant. I am here to help always with a smiling face.\n I am new to the assisting job, so please bear me if i make any mistakes. Hope you would be kind with me so that we achieve great things together and assisting you would be easy for me."
                print(say)
                self.speak(say)
                qf = 1
                break

        ai_chat = ["ai assist", "chat online", "ai chat", "text generator", "write"]
        for i in ai_chat:
            if i in ques:
                self.speak("Sure.Initiating new chat session.")
                self.ai_assist()
                qf = 1
                break

        if qf == 1:
            pass
        else:
            say = "Sorry,unable to provide assist for it. Do you want to check it in online AI chat? "
            print(say)
            self.speak(say)
            decision = input("Enter '1' for online AI assist otherwise '0' to continue: ")
            if decision == '1':
                self.speak("Sure. Initiating new chat session.")
                self.ai_assist()
            else:
                self.ask_again()
    
    def ai_assist(self):
        say = "*To end online assist,give input as 'end' or 'close chat' or 'exit chat' "
        self.speak(say)
        print(say)
        while(True):
            tell="Please ask your query."
            print(tell)
            self.speak(tell)
            query=self.ask_ques(1)
            if query=='end' or query=='close chat' or query=='exit chat':
                break
            else:
                try:

                    result = DDGS().chat(keywords=query, model="mixtral-8x7b", timeout=20)
                    wrapper=textwrap.TextWrapper(width=100)
                    word_list = wrapper.wrap(text = result)
                    print('\n')
                    for word in word_list:
                        print(word)
                    s1 = "Do you want to save it as a test file or save as a audio file or read it aloud?"
                    print("\n", s1)
                    self.speak(s1)
                    s1_ip = input("'A' for save as audio file or 'R' for read aloud or 'T' to save as text file\nInput:")
                    if s1_ip == 'A':
                        baseName= "GenAudio"
                        suffix = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
                        fname = "_".join([baseName, suffix]) # e.g. 'GenAudio_240624_171442'
                        file_format = ".mp3"  # e.g. .mp3 for audio
                        full_filename = fname + file_format 
                        self.talk.save_to_file(result, full_filename)
                        self.speak("Audio file saved.")
                    elif s1_ip == 'R':
                        self.speak(result)
                    elif s1_ip == 'T':
                        baseFileName = "GenText"
                        suffix = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
                        fname = "_".join([baseFileName, suffix])
                        file_format = ".txt"
                        full_filename = fname + file_format
                        with open(full_filename, 'w') as file:
                            strTime = datetime.datetime.now().strftime("%H:%M:%S")
                            file.write(strTime + "\n")
                            for word in word_list:
                                file.write(result + '\n')
                            self.speak("File saved.")              
                    else:
                        pass
                except Exception as e:
                    k = "No response...There was some error while connecting to the internet.\nPlease check your connection and try again."
                    print(k)
                    self.speak(k)
                    print("Exception arising due to")
                    print(e)
                    self.ask_again()

if __name__ == "__main__":
    assistant = Desk()
    assistant.wish()
    assistant.input_mode()
    assistant.basic_fun()


                
        
            




