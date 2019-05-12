<p align="center">
    <img src="https://i.imgur.com/RMuc1ki.png">
</p>
# CatSync LRMS

Just another dumb library software thingy made for college finals. And since everyone will do something related to this, use this as a template. Uses 3 separate databases to store a single entry. Worst back-end code a college student can write. But then again who reads the backend code as long as it works? Just hope that your college professor doesn't find this.

## Features
- Send email when book is late
- Write up late fee slips
- Multiple currencies
- LTT easter eggs
- Bad code (So less suspicion that this code was downloaded from the internet)

## Dependencies
EA Send Mail Component : https://download.cnet.com/EASendMail-SMTP-Component/3000-2070_4-10521758.html

## Bugs
- Writing late fee slips usually dont work
- Emails are sent as long as the program is running and a user is switching windows
  - Mainly because I didn't add a loop instead I added a check late function in the OnLoad() function of each window
  - also there is no background processes to send the email if the program is closed
- Some rented books dont showup in the borrowed list
