# excel-grading

- Overview: Macros, scripts, and sundry for those using Excel to do grading in courses
- Author:  Joseph Timothy Foley
- Email: foley AT RU dot IS or foley AT MIT dot EDU
  
As a teacher at Reykjavík University, I find myself using Microsoft Excel often to give detailed feedback to students about their assignments.
To keep myself sane and eliminate repetitive tasks, I developed VBA code and scripts to make this easier.
In particular, it is helpful if you setup a template grading sheet, then use the macros to auto-fill and create tabs for each student/team.
Then use the PDF macros to output individually named PDF files for uploading to a LMS such as CANVAS.

You will want to load `grading-macros.xlsm` in Excel.
Safety tip:  I highly suggest that you disable the macros when you first open it and look through the code in the Developer tab to make sure that nothing too strange is in there (such as someone adding a trojan or bitcoin miner somehow).
**Running VBA macros without checking what they do is very much light running around with scissors:  maybe you will be OK, but maybe you could get seriously injured.  You have been warned.**
Once you have done a quick sanity check, re-open the file and enable the macros.

Excel stores the VBA code in binary format side `.xlsm` files, which makes examining the code outside of the application hard.
It also makes it really tricky to see what elements of the code changed.
I have included a (hopefully) up-to-date export of the code in the `grading-macros.cls` file.
In order to use this in your own custom file, you will need to open your `.xslm` file and past it into the "ThisWorkbook" module.
If you just import it in the VBA editor, it puts it into a new module which I haven't figured out how to use.
*Note to self:  figure out how this works, someday.*

More and more people are using Excel on Office360, which does not have access to VBA.
To that end, I started collecting Office360 Automation code in the excel360-grading.ts file.
Sadly, there seems to be no way to output multiple PDF files.
You might be able to output the entire workbook and use a PDF editor to split it into multiple files.

Enjoy!
