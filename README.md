This VBA Project basically automate a process of creating final analysis sheet which usually takes 30-35 minutes manually.
In this, I have created 4 modules each doing different process .
Step -1 >First Module create a blank Excel spreadsheet and prompt the user to open an .csv file which has been downloaded from application.
Step -2 >The file downloaded from application has columns which are not in proper format. Now this module will arrange columns in pre defined format and copy it to New created spreadsheet.
Step -3 > In main Macro, now its time to run another Macro to calculate Variance and status . Then those columns will also be copied to New Spreadsheet.
Step -4 > Clear Macro module - this will clear the second macro which we run inside main macro in Step3
Step -5 > Module 4 - will create 2 Pivot Tables based on the Data in second Tab and rename it .
Step -6> There is predefined Tables in which the whole analysis and end result need to be written. Those will gets copied above Pivot Tables.
Step -7 > Close all the files and save New Spreadsheet created at Step 1.

All these steps will be done just by click of one RUN Button.

This VBA Code will cut copy paste data in to new file , use Vlookup function , create PIVOT Tables, run another Macro to generate Variance and Status of accounts.
It helps in increasing the productivity as this task took approx half an hour of an associate to prepare it manually.


You can downlaod the Test_Format_Macro_New.xlsm file and use the raw data to run it . 
