***************************************************************

                     GRADEBOOK EMAILER

                    Written by: Jason Lee

                        Version 0.2

***************************************************************

Hi! This is a script that I write that automatically emails parents, specifically used with eSchoolPlus TAC.

To use this, you will need to do a few things:

1. Use this Google Form to gather student and parent information.

   Form: https://docs.google.com/forms/d/1eO8z_cl08UhgAARuNm2BMLg_3LZQZ95bRbWRmGEap90/copy

   You need to use this form, or at least a form exactly formatted like this.
   The program is NOT dynamic to read other forms. Your emails will not read correctly if you do not use this.

   When all of your students fill this out, go to "Responses", click the three dots on the right, and select
   "Download responses (.csv)". Place this CSV file in the same folder as autoEmailer.exe. Rename it to something
   easily typeable, as you will have to type the filename into the program.

2. Get student grade data by going onto TAC. Go into the gradebook for the class you want to email.
   At the top right, there is a drop-down menu that probably says "Actions/Reports". Select "Printable Gradebook".

   **** IMPORTANT ****
   You NEED to have ONLY "Show Grade" selected under Options. Otherwise the program will not work.

   After you have selected "Show Grade", click run. Make sure that the first four columns of the spreadsheet are
   "Student ID", "Student Name", "Grade", "Average". Download this .xlsx file, put it in the same folder as autoEmailer.exe,
   and then rename it to something easily typeable, as the program will ask for the filename.

Now you should be ready to use the program. The first time you use it, there is a short first time setup that involves
granting access to sending emails from your work email account. This information is obviously only locally stored
as it is written by little Jason. You'll get a warning that this program isn't approved by Google, because it's not; I just
wrote this by myself. The developer is should be "vrhsmath.no.reply@gmail.com" a Gmail account I made up just for this.
If it's not, then that's a problem and this has somehow gotten compromised, which is bad and you shouldn't use it.

To bypass this screen, click "Advanced" on the left and click "Go to VRHS Automatic Parent Emailer (unsafe)". Again,
this program is safe, but Google just doesn't know that for sure and doesn't want to get sued. Allow everything, then
follow the instructions on the console.

***************************************************************
                        More Information
***************************************************************

If you want to change emailtemplate.txt, you should understand how it works.
You can see an example sent email at exampleemail.txt.
Essentially, the program looks for these keywords:

    "PARENT_FIRST_NAME"
    "PARENT_LAST_NAME"
    "STUDENT_FIRST_NAME"
    "STUDENT_LAST_NAME"
    "COURSE_NAME"
    "FAILING_ASSIGNMENTS"
    "TEACHER_EMAIL"

The program replaces "COURSE_NAME" with your selected course in the beginning, "TEACHER_EMAIL" with your email address you
provided, "FAILING_ASSIGNMENTS" from a list of assignments with grades = "M" or < 70, and the rest from the
student data .csv file you provided from the Google Form above. These keywords must be in capital letters to be replaced.

------------------------------------------------------------------------

You can put in URLs to each assignment in "URLs.ini". If you open it in a text editor (Notepad), put the URL on the
right side of the equal sign for every assignment. Then, the URL will be included in the email.
You should update the URLs everytime you have changed the assignments in your gradebook. You don't need to change it
for each class though.

------------------------------------------------------------------------

If you messed up on the first time setup, you can go into "emailconfig.ini" and change "first_time" to True. Then it
will run the first time configuration again. Alternatively, you can type in the information yourself if you want to
do that instead.

------------------------------------------------------------------------

If there is anything here that I didn't cover, let me know at jlee4804@gmail.com (my personal email).

Changelog:

0.1 - added file and error checking.
0.2 - fixed compatability with X and bugs with blank grades. Made inputting URLs better.
