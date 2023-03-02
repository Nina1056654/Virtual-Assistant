import pyttsx3
from datetime import datetime
import wikipedia
import pywhatkit as kit
import webbrowser
import smtplib
import imaplib
import email
from email.message import EmailMessage
from decouple import config

EMAIL = config("EMAIL")
PASSWORD = config("PASSWORD")

def speak(audio):
	
	engine = pyttsx3.init()
	# getter method(gets the current value
	# of engine property)
	voices = engine.getProperty('voices')
	
	# setter method .[0]=male voice and
	# [1]=female voice in set Property.
	engine.setProperty('voice', voices[0].id)
	
	# Method for the speaking of the assistant
	engine.say(audio)
	
	# Blocks while processing all the currently
	# queued commands
	engine.runAndWait()

def Hello():
	
	# This function is for when the assistant
	# is called it will say hello and then
	# take query
	speak("Hello sir, I am your virtual assistant. Tell me how may I help you?")

def tellDay():
	
	# This function is for telling the
	# day of the week
	day = datetime.today().weekday() + 1
	
	#this line tells us about the number
	# that will help us in telling the day
	Day_dict = {1: 'Monday', 2: 'Tuesday',
				3: 'Wednesday', 4: 'Thursday',
				5: 'Friday', 6: 'Saturday',
				7: 'Sunday'}
	
	if day in Day_dict.keys():
		day_of_the_week = Day_dict[day]
		print(day_of_the_week)
		speak("The day is " + day_of_the_week)


def tellTime():
	
	# This method will give the time
	time = datetime.now()
	
	# the time will be displayed like
	# this "2020-06-05 17:50:14.582630"
	# and then after slicing we can get time
	print(time)
	hour, minute = time.strftime("%H"), time.strftime("%M")
	speak("Sir, the time is" + hour + "Hours and" + minute + "Minutes")

def open_gmail():
    webbrowser.open_new_tab("https://mail.google.com")
    speak("Google Mail open now.")

def search_on_wikipedia(query):
    # it will give the summary of 3 lines from
    # wikipedia we can increase and decrease
    # it also.
    result = wikipedia.summary(query, sentences=3)
    speak("According to wikipedia")
    print(result)
    speak(result)


def play_on_youtube(video):
    kit.playonyt(video)


def search_on_google(query):
    kit.search(query)

# def sendEmail(to, content):
#     server = smtplib.SMTP('smtp.gmail.com', 587)
#     server.ehlo()
#     server.starttls()
     
#     # Enable low security in gmail
#     server.login(EMAIL, PASSWORD)
#     server.sendmail(EMAIL, to, content)
#     server.close()

'''This feature is no longer supported as of May 30th, 2022. Need to use an app password to access google account.'''
def sendEmail(to, subject, message):
    try:
        email = EmailMessage()
        email['To'] = to
        email["Subject"] = subject
        email['From'] = EMAIL
        email.set_content(message)
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.ehlo()
        server.starttls()
        server.login(EMAIL, PASSWORD)
        server.send_message(email)
        server.close()
        return True
    except Exception as e:
        print(e)
        return False


def checkMail():
    server = imaplib.IMAP4_SSL("imap.gmail.com")
    try:
        server.login(EMAIL, PASSWORD)
    except Exception as e:
        if "[AUTHENTICATIONFAILED]" in str(e):
            return None
        else:
            raise e
    server.select(readonly=True)
    messages = []
    typ, data = server.search(None, '(UNSEEN)')
    for num in data[0].split():
        typ, data = server.fetch(num, '(RFC822)')
        for response_part in data:
            if isinstance(response_part, tuple):
                email_parser = email.parser.BytesFeedParser()
                email_parser.feed(response_part[1])
                msg = email_parser.close()
                messages.append({"HEADERS":{header.upper(): msg[header] for header in ['to', 'from', "subject"]}})
                body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        ctype = part.get_content_type()
                        cdispo = str(part.get('Content-Disposition'))
                        if ctype == 'text/plain' and 'attachment' not in cdispo:
                            body = part.get_payload(decode=True)
                            break
                else:
                    body = msg.get_payload(decode=True)
                messages[-1].update({"BODY": body.decode()})

    server.close()
    server.logout()
    return messages