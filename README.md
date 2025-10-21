# bulk_emails_sender_python

## Needed dependencies Windows
- python
- openpyxl

## Needed dependencies linux
### Arch Linux 

`sudo pacman -S python python-openpyxl tk`

### Debian/Ubuntu

`sudo apt-get install tk8.6 tk8.6-dev`

`sudo pip3 install openpyxl `

### Fedora

`sudo dnf install tk tk-devel`

`sudo dnf install python3-openpyxl`


## Gmail Setup for App

1. Set up Gmail App Password (required for security):
a. Go to your Google Account settings
b. Enable 2-factor authentication (must have for app password)
c. Go to Security → App passwords (if you dont see it you can search in searchbox on top)
d. Click on create Apppassword and ccoose name for app (ex. Mail) and it will Generate password for you
e. Copy the password and keep it save 
f. Use this 16-character password in the app instead of your regular Gmail password

2. Open the App
a. open terminal (in windows as administrator, linux with sudo for elevated privileges) and run the script
b. when app opens click on choose excel to choose the excel document for receivers
c. click the dropdown to choose which sheet from the excell document you need and press load sheet o load the excel documnet
d. write your email and Apppassword (not the gmail password for mail)
e. write subject for the mail
f. write the mesage text for the mail ( {name} menas tha the app will append the name of the receiver where is needed in the tekst)
g. after done with everything above click send mails and than yes for confirming and sending the mails.
