#Windows 10 Universal Windows Platform and Office 365 Calendar Sample

This is a sample integration of the office 365 calendar in a windows 10 app

#Broken Connect Dialog in VS2015 14.0.24720.00 Update 1
Since I am not sure if it is my Dogfood VS or the dialog in general here is the unofficial guide.

##1. Sign in to Azure Management Portal
http://manage.windowsazure.com

Browse to the Active Directory Section 

Add a new App

Just enter a return URI which is a valid URL

You should get a APP/Id which needs to be stored in the app.xaml

Next time you sign in the dialog tells you the retrun url of the app ms-app:/xxxxxxxxxxx is not valid.
Copy that string and paste it in the azure ad configuration.

Done!

Notice you should use the final ms-appx package guid which you get by associating the app to the store.

  