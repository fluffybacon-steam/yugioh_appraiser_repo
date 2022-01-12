# YuGiOh Card Appraiser

####Language:Python
####Required libraries: selenium, pandas
####Required: chrome_webdriver.exe that supports your version of Chrome (https://chromedriver.chromium.org/downloads)
####Required: Microsoft Excel
####Required: A boat load of YuGiOh Cards

##Background 
A few months ago I uncovered my old but vast YuGiOh card collection from my youth. While I've totally forgotten how to play, I figured it would be cool to see if they had appreciated in valuve any more than my aunt's beanie babies. So I plugged all my cards into an excel spreadsheet and then went to https://yugiohprices.com/ to see what I have got. This process was incredibly tedious, so I wrote this script.

###The Script
cardprice.py is a simple Python webdriver that iterates through an excel file (.xls) of YuGiOh cards and returns their current market price. The excel file must be formated like so:

|    A      |      B       |       C      | ...
   Name	     Card Edition	   Set Number
DELINQUENT DUO	Unlimited Edition	MRL-039
HARPIE'S FEATHER DUSTER	Unlimited Edition	SDD-003
DARK MAGICIAN	Unlimited Edition	BPT-007
BLUE-EYES WHITE DRAGON	Unlimited Edition	BPT-009
BLUE-EYES WHITE DRAGON	Unlimited Edition	SDK-001
SNATCH STEAL	Unlimited Edition	DB1-EN021
![image](https://user-images.githubusercontent.com/39141161/149043432-ddc8a3c7-6695-4b40-82ea-b3bf27ab0a39.png)

######Disclaimer: this is highly dependent on https://yugiohprices.com/'s html so any changes on their end will break this.

###How to
The script will first ask you for a path to the excel file (e.g  \Users\baily\Scripts\Python\YuGiOh 401k\myexcelfile.xls)
Next, it will ask you if you have placed the chrome_webdriver.exe into the same directory above.
Then it will run. Likely for a very long time depending on your collection size. Printouts to the terminal will assure you it is still indeed running.
At completion the results will be written to a new excel file "results.xls". There you can copy the prices and then paste it alongside the cards in your original excel. I would have appended the results directly with ExcelWriter but support for that mode got rugged.

See you aall in the shadow realm.
