# Automatic email reminder using Python and Windows Task Scheduler!

Description:
- This simple but powerful code uses win32.client library for sending Outlook emails configured on your machine.

Pre-requsites:
- Have MS Outlook installed, configured and open.
- Python installed on your laptop
- Python path added to Environment variables.(Usually automatically added when installing python)

Steps:
- Clone the repository to your local machine.
- Install the requirements using pip install -r requirements.txt
- Navigate to the project folder on your terminal .\cd Outlook
- Open Timesheet_reminder_oultook.py and add desired email list in line number 17 of the script(list_of_emails).
- Run the CODE!!! using below command on your terminal
- python Timesheet_reminder_oultook.py

NOTE:
- This code sends out an email each time the code is run. However, if you want an automatic schdeuler to send the reminder email on a desired day and time, use the following steps.

*Using Windows Task Sceduler to automatically send the reminder email every Friday at 9:30 AM*
 - Follow the steps as illustated in the screenshots for Automatic scheduling(running the code every friday at 9:30 am)
 - Open Windows Task Scheduler App. (Press Windows button > Search for 'Task Sceduler App' and follow along with the screenshots

_**This is the output of the code**_

![Output mail.png](..%2F..%2FOneDrive%20-%20PQA%2FDesktop%2FOutput%20mail.png)

![1-Create Task](https://user-images.githubusercontent.com/122895165/217017218-e1b02cfe-8f2e-4eb7-b0a9-8f3a30041c6c.png)
![2-Set it for Weekly](https://user-images.githubusercontent.com/122895165/217017199-3104c865-0a98-49ae-a8a1-57139b04ebc1.png)
![3-Set Friday](https://user-images.githubusercontent.com/122895165/217017205-08522fe4-25ab-4bde-882e-c10f2ca1d765.png)
![4-Choose Start a Program](https://user-images.githubusercontent.com/122895165/217017210-6007cd0f-d5e3-4679-8478-7d0051271ec9.png)
![5-Enter the path of py file](https://user-images.githubusercontent.com/122895165/217017212-30993a92-ba9e-4119-9dce-e31acfd20b73.png)
![6-Finish](https://user-images.githubusercontent.com/122895165/217017215-51d6e067-7d9f-46ca-a9c8-8f436b2227b8.png)

**Make sure to edit the list of emails in the script for the magic to happen!**

**CAUTION:
Don't Spam the actual email receiver. Test it if you need to, using your own emails Ids :)**

You are welcome to contribute!!

Feel free to contact me if you need help setting up!
