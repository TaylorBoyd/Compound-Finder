# Compound-Finder

Compound Finder version v1.2.0

-----------------------------------
Installation
-----------------------------------

	- Program comes ready to run just double click the .exe file

-----------------------------------
Usage
-----------------------------------

	- Program takes in a .txt file from the GCMS and outputs a .xls file with some peak guesses attached.
	- It makes these guesses by comparing the observed retention time values to the values in the Retention Time Library. If they are within a range of +/-0.1 it will print all the guesses in ascending order.
	- Guesses with a (??) placed in front of it were outside the range. The program just chooses the closest one and puts the (??) tag in front to signify that it could be a terrible guess.
	- Choose a Retention Library to use, then a GCMS file, and they type in what you would like the output file name to be.
	- For now all outputs will be placed into the same folder as the .exe file

-----------------------------------
Version 1.1.0 Changes
-----------------------------------
	
	- Many of the functions rewritten to be easier to scale in the future if we want more additions/info to be displayed
	- Instead of reading from temp.txt the entire time it now pulls all the data in as a 2-dimensional list and builds on that to add more/new info
	- Changed all functionality to now output as a .xls file
	- Pruned main data page, added deleted data to "More Info" tab
	- Added compound percentage feature
	- Fixed formatting of cells to accomodate long compound names

-----------------------------------
Version 1.2.0 Changes
-----------------------------------
	
	- Added functionality to pull info off the Rumplestilskin file and add it to a template CofA tab
	- When the box is checked it pulls info off Rumplestilskin based off the user specified lot number and places it in a template on the third tab.
	- IMPORTANT: Make sure the file name for Rumplestilskin is not changed
	- IMPORTANT: Feel free to add more infor to the Rumplestilskin file at any time. Just make sure the END tag at the end of the file is still there and the new info is place above it. The program will still work without the END but it will go through the entire 60k rows of the excel file. Any info placed AFTER the END will not be looked at.
