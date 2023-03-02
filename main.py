from datetime import datetime
from voice_recognizer import *
from functions import *
from excel_automation import *
import re
import os

# main method for executing
	# the functions
if __name__ == '__main__':

	# calling the Hello function for
	# making it more interactive
    Hello()

    # This loop is infinite as it will take
	# our queries continuously until and unless
	# we do not say bye to exit or terminate
	# the program
    while(True):
		
		# taking the query and making it into
		# lower case so that most of the times
		# query matches and we get the perfect
		# output
        query = takeCommand().lower()

        if 'open youtube' in query:
            speak('What do you want to play on Youtube, sir?')
            video = takeCommand().lower()
            play_on_youtube(video)

        elif 'open google' in query:
            speak('What do you want to search on Google, sir?')
            query = takeCommand().lower()
            search_on_google(query)

        elif 'open gmail' in query:
            open_gmail()
            
        elif "day" in query:
            tellDay()

        elif "time" in query:
            tellTime()

        elif "wikipedia" in query:
            
            # if any one wants to have a information
            # from wikipedia
            speak("Checking the wikipedia now.")
            query = query.replace("wikipedia", "")
            search_on_wikipedia(query)

        elif "note" in query:
            speak("Start recording.")
            filename = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
            
            # Write the transcribed text to a .txt file
            with open(f"{filename}.txt", "w") as file:
                while True:
                    text = takeCommand().lower()

                    # # check the quality of the transcribed audio
                    if not YN_check(f"Did you just say '{text}'?"):
                        speak('Try again then.')
                        continue
                    if text == "stop recording":
                        print("Recording stopped.")
                        break
                    file.write(text + '\n')
            print(f"Transcription saved to {filename}.txt")

        elif all([re.search(w, query) for w in ['send', 'email']]):
            wehi = False
            if YN_check("Do you want to send a guy in WEHI?"):
                speak("whome should I send?")
                username = input("Enter username: ")  
                to = f"{username}@wehi.edu.au"
                print(to)
                if YN_check("Is this the correct address?"):
                    wehi = True
            if not wehi:
                speak("what is the email address then?")
                to = input("Enter email address: ")   
            speak("What should be the subject?")
            subject = takeCommand().capitalize()
            speak("What is the message?")
            content = takeCommand().capitalize()
            if sendEmail(to, subject, content):
                speak("I've sent the email sir.")
            else:
                speak("Something went wrong while I was sending the mail. Please check the error logs sir.")

        elif all([re.search(w, query) for w in ['check', 'email']]):
            messages = checkMail()
            if messages is not None:
                if messages:
                    print("UNREAD MESSAGES (%s)" % len(messages))
                    i = 0
                    for m in messages:
                        print("---------------")
                        for l in m["HEADERS"]:
                            print("%s: %s" % (l.title(), m["HEADERS"][l]))
                        print()
                        body = m["BODY"]
                        body = re.sub(r"[\n\r]+","\n  ",body)
                        if len(body) > 1000:
                            body = body[:1000]+"..."
                        print("  "+body)
                        print()
                        i += 1
                        if i % 5 == 0:
                            print("---------------")
                            if not YN_check("%s messages remaining. See more? " % (len(messages)-i)):
                                print("Exited mail")
                else:
                    speak("No unread mail in your inbox.")
            else:
                speak("Invalid credentials, please try again sir.")

        elif all([re.search(w, query) for w in ['to', 'spreadsheet']]) or all([re.search(w, query) for w in ['to', 'excel']]):
            speak("Which audio transscript would you want to import into a spreadsheet?")
            filename = input("Enter the timestamp of the audio (yyyy_mm_dd_hh_mm_ss): ")  
            spreadsheet_name = input("Enter the name of the spreadsheet: ")
            if spreadsheet_name == '':
                spreadsheet_name = filename
            sheet_name = input("Enter the name of the sheet: ")
            if sheet_name == '':
                sheet_name = "Sheet1"
            # new = YN_check("Do you want to create a new spreadsheet?")
            header = ''
            if not os.path.isfile(f"{spreadsheet_name}.xlsx"):
                speak("What metadata do you want to record?")
                header = input("Enter the header of the sheet, split by ',': ")
            script_to_spreadsheet(header=header, transcript_name=filename, spreadsheet_name=spreadsheet_name, sheet_name=sheet_name)
            speak("Import finished, please check.")

        elif all([re.search(w, query) for w in ['open', 'spreadsheet']]) or all([re.search(w, query) for w in ['open', 'excel']]):
            speak("Which spreadsheet do you want to open?")
            spreadsheet_name = input("Enter the name of the spreadsheet: ")  
            speak("Which sheet do you want to edit?")
            sheet_name = input("Enter the name of the specific sheet: ") 
            if sheet_name == '':
                sheet_name = "Sheet1" 
            os.system(f"open -a 'Microsoft Excel.app' {spreadsheet_name}.xlsx")
            while(True):
                query = takeCommand().lower()
                if all([re.search(w, query) for w in ['edit', 'cell']]):
                    if 'heading' in query:
                        by_location = True
                        speak('Which cell do you want to edit, sir? Please give me the row and column headings.')
                        cell_loc = takeCommand().upper().replace(" ", "")
                    else:
                        by_location = False
                        speak('Which cell do you want to edit, sir? Please give me the information about the cell location. Note that all the unique keys as well as the column name of the new value are required.')
                        cell_pos = takeCommand()
                    speak('What is the new value?')
                    new_value = takeCommand().lower()
                    editable = edit_spreadsheet_cell(spreadsheet_name, sheet_name, by_location, cell_loc, new_value)
                    os.system('osascript -e \'quit app "Microsoft Excel"\'')
                    if editable:
                        speak("Cell edited, please check.")
                        os.system(f"open -a 'Microsoft Excel.app' {spreadsheet_name}.xlsx")
                    else:
                        speak("I cannot locate the cell, sir. Please check if the information is complete or try usinng the row and column headings instead.")
            os.system('osascript -e \'quit app "Microsoft Excel"\'')

        elif query != "None":
            speak('Sorry sir, I do not have this feature at the moment. Please upgrade me.')

        