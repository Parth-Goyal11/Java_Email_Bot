# Java_Email_Bot
An automated JAVA Email Bot I made that uses the Java Mail api to send mail through gmail servers, and the APACHE Poi api to parse spreadsheet data to make user input easier. The program takes a spreadsheet of given emails and names and sends an email to all of them. 

How to set this up on your computer:

1. Go to google account settings
2. Allow third party apps
3. Turn on multi-factor authentication
4. Generate a 16 character app passcod
5. Download the javamail jar files and import them into your project
6. Download the JavaBeans Activation jar files and import them into your project
7. Download the Apache POI api and import the jar files into your project. 

How to use:

1. Open bot.java in src.
2. Replace the password variable with the 16 character passcode you obtained
3. Replace sender with your email address
4. Open the TestFile.xlsx and put the email addresses you want to send and email to along with their names.(KEEP CURRENT FORMAT OF THE SPREADSHEET AS IT IS)
5. Change the message by changing the email.setSubject and email.setText lines.
6. With some JAVA knowledge you can add other variables to the emails.
7. In order to add a new variable create it as a column in the spreadsheet, and use the returnNames() and returnEmails() functions as references to add that new variable into your email.
