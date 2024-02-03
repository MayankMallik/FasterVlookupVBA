# FasterVlookupVBA
This repository contains a macro-enabled excel file (.xlsm file), and 2 text files which contain the code of userform and module respectively.

In the excel file, there is 1 sheet named Instructions which contains all the instructions and disclaimers to be taken into consideration before running vlookup

It uses concept of dictionaries. 2 dictionaries store the data from the two excel files that are input. 

Then it compares the 2 dictionaries and paste the values. Given the common column should be a primary key (It should not have duplicate values)

This file has been tested on an excel file with more than 900,000 rows. Normal vlookup sometimes hangs the system and other times takes more than 1 hr
This faster vlookup takes 2-6 mins depending on the configuration of your system. And it has 100% accuracy.
