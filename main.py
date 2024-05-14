import datetime
import os
import speech_recognition as sr
import win32com.client
import webbrowser
import datetime
import openai
import apikey

speaker = win32com.client.Dispatch("SAPI.SpVoice")

chatStr = ""

def chat(query):
    global chatStr
    apiKey = apikey.apikey
    openai.api_key = apiKey
    chatStr += f"Dhruv: {query}\n Oasis: "
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt= chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    # todo: Wrap this inside of a  try catch block
    print(chatStr)
    chatStr += f"{response['choices'][0]['text']}\n"
    speaker.Speak(response["choices"][0]["text"])
    return response["choices"][0]["text"]

def ai(prompt):
    apiKey = "Enter Your API Key"
    openai.api_key = apiKey

    text = f"OpenAI response for the Prompt: {prompt}\n **************************************************\n\n"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.9,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    print(response["choices"][0]["text"])
    text += response["choices"][0]["text"]

    if not os.path.exists("OpenAI"):
        os.mkdir("OpenAI")

    with open(f"OpenAI/prompt-{prompt[0:30]}.txt","w") as f:
        f.write(text)

def dirSearch(fileToSearch,query):
    rootdirC="C:/"
    rootdirD="D:/"
    rootdirE="E:/"

    if "D drive".lower() in query.lower():
        for relpath, dirs, files in os.walk(rootdirD):
            if (fileToSearch in files):
                fullpath = os.path.join(rootdirD, relpath, fileToSearch)
                os.startfile(fullpath)

    elif "E drive".lower() in query.lower():
        for relpath, dirs, files in os.walk(rootdirE):
            if (fileToSearch in files):
                fullpath = os.path.join(rootdirE, relpath, fileToSearch)
                os.startfile(fullpath)

    elif "C drive".lower() in query.lower():
        for relpath, dirs, files in os.walk(rootdirC):
            if (fileToSearch in files):
                fullpath = os.path.join(rootdirC, relpath, fileToSearch)
                os.startfile(fullpath)


def TakeCommand():
    r = sr.Recognizer()
    mic = sr.Microphone()
    with mic as source:
        audio = r.listen(mic)
        try:
            query = r.recognize_google(audio, language="en-in")
            print("Recognising...")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Working for further improvement."

if __name__ == '__main__':
    print("OasisAI.")
    speaker.Speak("Hello I am Oasis")
    followCommand = True
    while followCommand:
        print("Listening...")
        speaker.Speak("Listening...")
        query = TakeCommand()
        #todo:add more websites
        sites = [["youtube","https://youtube.com"],["google","https://google.com"],["instagram","https://instagram.com"],["anix","https://anix.to/home"],["apna college","https://www.apnacollege.in"],["wikipedia","https://wikipedia.com"],["gmail","https://gmail.com"],["amazon","https://amazon.com"],["whatsapp","https://web.whatsapp.com"],["telegram","https://web.telegram.org/k/"],["codechef","https://codechef.com"],["github","https://github.com"],["flipkart","https://flipkart.com"],["aniwatch","https://aniwatch.to/home"],["manga reader","https://mangareader.to"],["movies","https://streamm4u.com"],["codedex","https://codedex.io"]]

        #todo:add more files
        files = [["one piece 1071","onepiece1071.mkv"],["one piece 1072","onepiece1072.mkv"],["one piece 1069","onepiece1069.mkv"],["one piece 1070","onepiece1070.mkv"],["extraction 2","extraction2.mkv"],["spider-man into the spider verse","spider man into spider verse.mkv"],["Spider-Man across the spider verse","Spider.Man.Across.The.Spider.Verse.2023.mkv"],["Operation fortune","Operation Fortune.mkv"],]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speaker.Speak(f"opening {site[0]} sir...")
                webbrowser.open(site[1])

        for file in files:
            if f"Open {file[0]}".lower() in query.lower():
                speaker.Speak(f"opening {file[0]} sir...")
                dirSearch(file[1], query)

        if f"play song".lower() in query.lower():
            musicpath = "C:/Users/Lenovo/Downloads/song.mp3"
            os.startfile(musicpath)

        elif f"play one piece".lower() in query.lower():
            videopath = "D:\Anime\One Piece\one piece 1071.mkv"
            os.startfile(videopath)

        # todo:add more applications
        elif f"open word".lower() in query.lower():
            wordpath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk"
            os.startfile(wordpath)
        elif f"the time" in query.lower():
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            speaker.Speak(f"The time is {strfTime}")

        elif f"the date" in query.lower():
            strfDate = datetime.datetime.now().strftime("%d:%m:%Y")
            speaker.Speak(f"The date is {strfDate}")

        elif f"ai".lower() in query.lower():
            ai(prompt=query)

        elif f"Close".lower() in query.lower():
            followCommand = False

        elif f"reset chat".lower() in query.lower():
            chatStr = ""

        else:
            print("Chatting...")
            chat(query)

print("Closing OasisAI.")
speaker.Speak("Closing OasisAI.")
