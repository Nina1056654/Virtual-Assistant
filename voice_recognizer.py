import speech_recognition as sr
from datetime import datetime
from random import choice
from utils import opening_text
from functions import speak


# this method is for taking the commands
# and recognizing the command from the
# speech_Recognition module we will use
# the recongizer method for recognizing
def takeCommand():

	r = sr.Recognizer()

	# from the speech_Recognition module
	# we will use the Microphone module
	# for listening the command
	with sr.Microphone() as source:
		print('Listening...')
		
		# seconds of non-speaking audio before
		# a phrase is considered complete
		r.pause_threshold = 0.7
		audio = r.listen(source)
		
		# Now we will be using the try and catch
		# method so that if sound is recognized
		# it is good else we will have exception
		# handling
		try:
			print("Recognizing...")
			
			# for Listening the command in indian
			# english we can also use 'hi-In'
			# for hindi recognizing
			Query = r.recognize_google(audio, language='en-in')
			print(f"You just said: {Query}\n")

			if not ('exit' in Query or 'stop' in Query or 'bye' in Query):
				speak(choice(opening_text))
			else:
				# this will exit and terminate the program
				hour = datetime.now().hour
				if hour >= 21 and hour < 6:
					speak("Good night sir, take care!")
				else:
					speak('Have a nice day sir!')
				exit()
		except Exception as e:
			print(e)
			print("Pardon me, please say that again.")
			return "None"
		
		return Query

def YN_check(question):
    speak(question)
    ans = takeCommand()
    if 'yes' in ans:
        return True
    elif 'no' in ans:
        return False
    else:
        speak("Answer is not clear, please say again.")
        YN_check(question)