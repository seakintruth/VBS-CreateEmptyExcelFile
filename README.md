# VBS-CreateEmptyExcelFile
Creates an Empty Excel File without using office automation.

Microsoft .xlsx files are just a Zipped collection of xml files

That means that we can write out the xml files and folder structure as text files, the zip the folder and rename that zip file to .xlsx. Now we have a new, empty excel file. Here is my code to do just that without having to have MS Excel installed.

I maintain a copy of this script is stack exchange here:
https://stackoverflow.com/a/67015856/1146659
