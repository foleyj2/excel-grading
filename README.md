# excel-grading
Overview: Macros, scripts, and sundry for those using Excel to do grading in courses

As a teacher at Reykjavík University, I find myself using Microsoft Excel often to give detailed feedback to students about their assignments.
To keep myself sane and eliminate repetitive tasks, I developed VBA code and scripts to make this easier.
In particular, it is helpful if you setup a template grading sheet, then use the macros to auto-fill and create tabs for each student/team.
Then use the PDF macros to output individually named PDF files for uploading to a LMS such as CANVAS.

You will want to load grading-macros.xlsm in Excel and enable the macros when prompted.
I highly suggest that you disable the macros when you first open it, look through the code to make sure that nothing too strange is in there, then close it, and reopen it.

Excel stores the VBA code in binary format side .xlsm files, which makes examining the code outside of the application hard.
I have included a (hopefully) up-to-date export of the code in the grading-macros.cls file.

More and more people are using Excel on Office360, which does not have access to VBA.
To that end, I started collecting Office360 Automation code in the excel360-grading.ts file.
Sadly, there seems to be no way to output multiple PDF files.
You might be able to output the entire workbook and use a PDF editor to split it into multiple files.

Enjoy!