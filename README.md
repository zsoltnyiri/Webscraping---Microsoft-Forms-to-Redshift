# Webscraping---Microsoft-Team-to-Redshift

Since there isn't any offical API yet to access the results of a Microsoft Forms' questionaire, this is a simple script to scrape it off the platform, and then write the results to S3 and ultimately to Redshift.

First logs into to Microsoft Platform with your creds
![login](/login.jpg)

![pass](/pass.jpg)

Then clicks through the interface, and downloads the designated csv
![dl](/download.jpg)

Finally export the data to Redshift
![redshift](/res.JPG)

Work in progress.
