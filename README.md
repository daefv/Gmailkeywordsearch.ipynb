# Gmailkeywordsearch.ipynb
this code works in google colab for python. This code connects to a Gmail 
account via IMAP and allows the user to specify a keyword to search for within the body of the 150 latest email messages. 
The code then counts the number of occurrences of the keyword in each email message and prints the total count at the end.

what is an app password for gmail and how do i get it? (please note, the App password 'yhtkxorkkrsydbfv' used in the code was once the correct App password, used to demonstrate what the output of the code will look like. It is no longer the App password to email account '5291ccc@gmail,' so using them in combination will return an authenication error. Please use your own email adress and App password)


An app password is a unique password that allows you to access your Gmail account from third-party apps that do not support two-factor authentication. 
This is to ensure the security of your account while still allowing you to use those apps.
To get an app password for Gmail, you need to follow these steps:

1. Go to your Google Account settings by visiting https://myaccount.google.com/ and sign in if necessary.

2. Click on the "Security" tab on the left-hand side of the page.

3. Under the "Signing in to Google" section, click on "App passwords."

4. You may be asked to enter your Google account password again for security reasons. 
After you enter your password, you will be taken to the app passwords page.

5. Select the app for which you want to generate an app password. (or select 'other' and name the App 'pyscript' - or anything

6. Follow the on-screen instructions to generate the app password. 
You may need to provide the app name, device type, and other details to generate the password.

7. Once the app password is generated, copy it and use it to sign in to the app.

The program should run in less than a minute assuming you are not purposely trying to break it. 

I used this program to find out that The keyword "review" appears 231 time(s) in the body of the latest 150 email messages.
which is ludacris, annoying and i hate when vendors solict reviews from me.
