# jobmailer
a simple tool for job seekers to inform employers they have not yet answered

# what you need
any Google account

# how to set it up
create empty spreadsheet in Google Sheets and choose __Tools > Script editor__
and copy jobmailer.gs there. save, close spreadsheet and reopen it.
You may need to give script permissions to run and access data!
Choose __Jobmailer > Create parameters__ and new sheet appears.
Feel free to modify it. With every _letter of interest_ create new table row
containing _date, recipient's email and job_ and one empty cell.

# how to send mails
select four columns and up to 100 rows with _date, recipient's email, job and fourth cell_.
the script reads line by line and assembles and sends mail messages for unanswered
(fourth cell is empty) _letters of interest_ older than defined days of grace.

Or grab a copy at
https://docs.google.com/spreadsheets/d/1YEqC5Qj2LiESJ5tN1-5Ie-lUguso8Bti_1pMxr-e6wA/edit?usp=sharing

# screenshots
![sample](https://raw.githubusercontent.com/vlna/jobmailer/master/doc/sample.png)

----
![mail](https://raw.githubusercontent.com/vlna/jobmailer/master/doc/mail.png)
