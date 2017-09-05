# Emailizer
Goes through an Excel file and removes rows that have duplicate emails

This program will allow a user to select a file with a specific file extension from the file explorer window and then will go through two columns with titles "email" and "hasEmail". As in, make sure those titles are in the first row. 

It will then go through those columns and first check if hasEmail is set to True. If so, it checks the row and see's if that email is already in the excel file. If it is, it will delete that entire row since it's a duplicate.

You will need this JAR files to fun the program:

poi-3.14-20160307.jar
poi-ooxml-3.14-20160307.jar
poi-ooxml-schemas-3.14-20160307.jar
xmlbeans-2.6.0.jar
