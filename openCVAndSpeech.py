import speech_recognition as sr
import cv2
import pyodbc
import playsound  # to play saved mp3 file
from gtts import gTTS  # google text to speech
import os  # to save/open files
import wolframalpha  # to calculate strings into formula
from selenium import webdriver  # to control browser operations




num  = 1
def assistant_speaks(output):
    global num

    # num to rename every audio file
    # with different name to remove ambiguity
    num += 1
    print("PerSon : ", output)

    toSpeak = gTTS(text=output, lang='en', slow=False)
    # saving the audio file given by google text to speech
    file = str(num) + ".mp3 "
    toSpeak.save(file)

    # playsound package is used to play the same file.
    playsound.playsound(file, True)
    os.remove(file)


def get_audio():
    r = sr.Recognizer()
    audio = ''

    with sr.Microphone() as source:
        print("Speak...")

        # recording the audio using speech recognition
        audio = r.listen(source, phrase_time_limit=10)
    print("Stop.")  # limit 10 secs

    try:

        text = r.recognize_google(audio, language='en-US')
        print("You : ", text)
        return text

    except:

        assistant_speaks("Could not understand your audio, PLease try again !")
        return 0


def process_text(input):
    try:
        if 'search' in input or 'play' in input:
            # a basic web crawler using selenium
            search_web(input)
            return

        elif "who are you" in input or "define yourself" in input:
            speak = '''Hello, I am Person. Your personal Assistant. 
            I am here to make your life easier. You can command me to perform 
            various tasks such as calculating sums or opening applications etcetra'''
            assistant_speaks(speak)
            return

        elif "who made you" in input or "created you" in input:
            speak = "I have been created by program ."
            assistant_speaks(speak)
            return

        elif "calculate" in input.lower():

            # write your wolframalpha app_id here
            app_id = "WOLFRAMALPHA_APP_ID"
            client = wolframalpha.Client(app_id)

            indx = input.lower().split().index('calculate')
            query = input.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            assistant_speaks("The answer is " + answer)
            return

        elif 'open' in input:

            # another function to open
            # different application availaible
            open_application(input.lower())
            return

        else:

            assistant_speaks("I can search the web for you, Do you want to continue?")
            ans = get_audio()
            if 'yes' in str(ans) or 'yeah' in str(ans):
                search_web(input)
            else:
                return
    except:

        assistant_speaks("I don't understand, I can search the web for you, Do you want to continue?")
        ans = get_audio()
        if 'yes' in str(ans) or 'yeah' in str(ans):
            search_web(input)


def search_web(input):
    driver = webdriver.Chrome(executable_path=r"D:\Project 1\Opencv\chromedriver.exe");

    driver.implicitly_wait(1)
    driver.maximize_window()

    if 'youtube' in input.lower():

        assistant_speaks("Opening in youtube")
        indx = input.lower().split().index('youtube')
        query = input.split()[indx + 1:]
        driver.get("http://www.youtube.com/results?search_query =" + '+'.join(query))
        return

    elif 'wikipedia' in input.lower():

        assistant_speaks("Opening Wikipedia")
        indx = input.lower().split().index('wikipedia')
        query = input.split()[indx + 1:]
        driver.get("https://en.wikipedia.org/wiki/" + '_'.join(query))
        return

    else:

        if 'google' in input:

            indx = input.lower().split().index('google')
            query = input.split()[indx + 1:]
            driver.get("https://www.google.com/search?q =" + '+'.join(query))

        elif 'search' in input:

            indx = input.lower().split().index('google')
            query = input.split()[indx + 1:]
            driver.get("https://www.google.com/search?q =" + '+'.join(query))

        else:

            driver.get("https://www.google.com/search?q =" + '+'.join(input.split()))

        return


# function used to open application
# present inside the system.
def open_application(input):
    if "chrome" in input:
        assistant_speaks("Google Chrome")
        os.startfile('C:\Program Files (x86)\Google\Chrome\Application\chrome.exe')
        return

    elif "word" in input:
        assistant_speaks("Opening Microsoft Word")
        os.startfile(r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE')
        return

    elif "excel" or "Excel" in input:
        assistant_speaks("Opening Microsoft Excel")
        os.startfile('C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE')
        return

    else:

        assistant_speaks("Application not available")
        return




#connect database
def getProfile(id):
    conn =  pyodbc.connect('Driver={SQL Server};'
                      'Server=DELL-PC\SQLEXPRESS;'
                      'Database=User;'
                      'Trusted_Connection=yes;')


    cmd = "SELECT * FROM UserData WHERE PersonID ="+str(id)
    cursor = conn.execute(cmd)
    profile = None
    for row in cursor:
        profile = row
    conn.close()
    return profile


# load data to detect
net = cv2.dnn.readNet('weights/yolov3.weights', 'cfg/yolov3.cfg')
classes = []
with open("coco.names", "r") as f:
    classes = [line.strip() for line in f.readlines()]
layer_names = net.getLayerNames()
output_layers = [layer_names[i[0] - 1] for i in net.getUnconnectedOutLayers()]


rec =  cv2.face.LBPHFaceRecognizer_create()
rec.read("recognizer/trainningData.yml")

#call back function
def callback(recognizer, audio):
    try:
        outp = get_audio()
      
        global val

        if (outp == 'can you hear me'):
            print('yes')
        if (outp == 'finish'):
            val = 1
            print('exiting...')
        if (outp == 'put text'):
            val = 2
        if (outp == 'refresh'):
            val = 0
        if (outp == 'can you see me'):
            val = 8
            print('yes')
        if (outp == "who am i" or outp == "who am I"):
            val = 9

        if (outp == 'good' or outp == 'excellent'):
            print('thank you')
    except sr.UnknownValueError:
        print('could not unsertand audio')
    except sr.RequestError as e:
        print('could not request from Google Speech Recognition service; {0}".format(e)')


r = sr.Recognizer()
m = sr.Microphone()



cap = cv2.VideoCapture(0)
val = 0
# imgsq = cv2.imread("img1.jpg")


font = cv2.FONT_HERSHEY_SIMPLEX
face_cascade = cv2.CascadeClassifier('D:\\PROGRAM LANGUAGE\\python\\speechrecoginzer\\haarcascade_frontalface_default.xml')


with m as source:
    r.adjust_for_ambient_noise(source)  # we only need to calibrate once, before we start listening
stop_listening = r.listen_in_background(m, callback)


while True:
    ret, img = cap.read()
    gray = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)

    if (val == 1):
        cv2.putText(img, "I can hear you", (30, 80), cv2.FONT_HERSHEY_SIMPLEX, 1.5, (0, 255, 255), 2)
    if (val == 8):
        img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        faces = face_cascade.detectMultiScale(img, 1.3, 5)
        for (x, y, w, h) in faces:
            cv2.rectangle(img, (x, y), (x+w, y+h), (255, 0, 0), 2)
            roi_color = img[y:y + h, x:x + w]
            rect = img[y + h // 2 - 100:y + h // 2+ 100, x + w // 2- 100:x + w // 2 + 100]

    if (val == 9):
        faces = face_cascade.detectMultiScale(img, 1.3, 5)
        for (x, y, w, h) in faces:
            cv2.rectangle(img, (x, y), (x+w, y+h), (255, 0, 0), 2)
            roi_color = img[y:y + h, x:x + w]
            id, conf = rec.predict(gray[y:y + h, x:x + w])
            profile = getProfile(id)
            if (profile != None):
                cv2.putText(img, "Name : " + str(profile[1]), (x, y + h + 20), font, 1, (0, 0, 255))
                cv2.putText(img, "Gender : " + str(profile[2]), (x, y + h + 45), font, 1, (0, 255, 0))
                cv2.putText(img, "Age : " + str(profile[3]), (x, y + h + 70), font, 1, (255, 255, 0))


    if (cv2.waitKey(10) & 0xFF == ord('b')):
        break
    if (val == 1):
        break
    cv2.imshow('Video', img)
cap.release()
cv2.destroyAllWindows()