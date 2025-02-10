# rdb-to-sheets
A very simple python script to convert the TXT output from RDB to an Excel Sheet

## How to use
1. On Meridian 1 do : LD 21; REQ PRT; TYPE RDB;
1. Copy the output from the first DN up to the last DN listed (see example below)
1. Save to somewhere as a raw txt
1. Use Bash script to set up Python3 VENV, or install requirements.txt things with pip on your system however you please.
1. Run : python3 ./rdb-to-sheets.py /your/rdb/file.txt 
1. Enjoy your now converted rdb excel sheet (made in same directory as rdb-to-sheets.py by default, can also be anywhere you desire python3 like this ./rdb-to-sheets.py /your/rdb/file.txt /your/output/rdb/sheet.xslx)

### Example of RDB.txt file contents
```
DN   0
TYPE NARS
NTBL AC1 

DN   1000
TYPE RDB 
ROUT 0 

DN   1001
TYPE RDB 
ROUT 1

.
.
Rest of it goes here
.
.

DN   #37
TYPE FFC 
FEAT RGAD 

DN   #51
TYPE FFC 
FEAT SPCC 

DN   #54
TYPE FFC 
FEAT RDST 
```
## Consideration
1. I know little coding and even less python, I've tested it on my Meridian 1 Option 11e only as thats the only system I own.
1. If you have suggestions on how to improve this feel free to submit those here.
1. Have fun !
