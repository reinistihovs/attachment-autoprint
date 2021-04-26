# attachment-autoprint
This powershell script fetches picture and pdf attachments from IMAP account and silently prints them automatically to the default printer.
You can run print.ps1 with powershell manually or create a scheduled task, which executes C:\attachment-print\task.vbs.
I have tested many ways to hide the running task, and had best results with vbs.

Instructions:
1) Place files in C:\attachment-print

2) edit the following lines in C:\attachment-print\print.ps1
$Username = "example@example.com"
$Password = "ExamplePass"
$client.Host = "mail.example.com"

3) Create a recurring task in task scheduler, which executes C:\attachment-print\task.vbs

P.S.
For PDF printing I have tested only with adobe reader set as default PDF program.


