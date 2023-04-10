# Virtual-Assistant
Genomics Metadata Multiplexing technical work

## Python Files
- `voice_recognizer`: Setup microphone and take commands.
- `excel_automation`: Contains functions about excel automation.
- `functions`: Contains common commands (e.g., reporting current datetime, doing online searches, watching YouTube videos, sending and checking emails, etc.)

### The implemented Voice assistant can perform the following tasks:


1. Opens a wepage : Youtube , G-Mail , Google Chrome 
	
	
		Human : Open [Youtube/Google/gmail]
		
		
2. Predicts date and time 
	
	
		Human : What [time/day] is it
		
	
3. Abstarct necessary information from wikipedia
	
   		
		Human: What is single cell sequencing according to Wikipedia
		
		
   The voice assistant abstarcts first 3 lines of wikipedia and gives the information to the user.
   
4. Record audio and trannscribe speech to text

		Human : I want to take notes
    
    The voice assistant will gennerate a {yyyy-mm-dd-hh-mm-ss}.txt file.
    
5. Send and check email

		Human : I want to [send/check] emails
    
6. Import text file to spreadsheet

		Human : Import transcript to [excel/spreadsheet]
    
7. Edit cells in spreadsheet

		Human : I want to [Open/Edit] [excel/spreadsheet]
		

Note: Talon scripts folder contains Python scripts and Talon files that define the voice commands for different applications. Talon is a machine learning tool which can help us interact with the computer using voice.
