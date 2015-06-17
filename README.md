
This Ruby script is meant to reduce the effort of reading from a result pdf and then manually typing in the numbers of students.


-------------------------------------------------------------------------------------------------------------------------------------------------------
Requirements:

> This script may run on any OS with ruby >= 1.9 installed.
> The script needs the gem 'pdf-reader' installed. You can install simply by typing "gem install pdf-reader" on a system where Ruby is already installed.
> The script needs the gem 'axslx' installed to make the excel spreadsheet. 
  You can install it simply by typing "gem install axlsx" on a system where Ruby is already installed.

-------------------------------------------------------------------------------------------------------------------------------------------------------
Running the script:
The syntax to run is :
user:~$ ruby automate.rb "filename.pdf"[important] "filename.xlsx"[optional]

The script will output some text. These are the entries which have long names or do not comply with the standards. 
In such records the names must be filled manually though the seat numbers and their marks will be filled in the spreadsheet.

In case of absentees the spreadsheet will have empty values.
The spreadsheet values have been tested for random records.

The spreadsheet column number one will give seat number, the second will give student's name and the third mother's name.
The fourth coulmn onwards the subjects marks will be set.
The subjects are numbered as they are in the pdf. 
Hence, if "DISCRETE STRUCTURES PP" is given 01. then the 4th column will be the marks for "DISCRETE STRUCTURES PP" and so on.

The pdf contains results of all branches, so there is a need to manually choose which records are needed which is a simple task and has been eliminated
from the script to avoid unnecessary complications.

-------------------------------------------------------------------------------------------------------------------------------------------------------
Flow of the program:
The script first creates an object of axlsx to create the excelf file.
Then the script opens the pdf goes through it page by page
Every page usually has three records.
Every pages is scanned line by line looking for the seat number which is handled by a regular expression.
Also, the end of a record is a line starting with "ORDN" hence that line marks end of record.
This record or an array of strings is sent to process.
In process, every line is checked and first line contains student details which are recorded in variables.
Then the subjects are checked for. Every line has information of two subjects so using regular expressions we can differentiate the subjects and store it
in array. That particular part of the string is converted into an integer.
Finally process returns an array containing the record as it will be shown in the spreadsheet and parse returns all the records found on the page in an
array. These records are then written to the spreadsheet.

-------------------------------------------------------------------------------------------------------------------------------------------------------
Contact:
I have made the flow of program and the running of script very clear. Still, if you have any doubts you can contact me.
The code is extremely unreadable, which is why I have provided the program flow.
You may make changes to the script as to your liking but make sure to have a backup.

E-Mail : anupsalve8@gmail.com 
