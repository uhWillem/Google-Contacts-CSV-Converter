# Google-Contacts-CSV-Converter

Convert your list of names, phone numbers, and email addresses to the Google Contact CSV format for easy import. 
This project is designed to work with an Excel file and has the pre-defined keywords in Dutch (first name -> voornaam). However, the code can be modified to suit your needs.

## Requirements
* Node.js installed
* Label added in the Contact_Info/labelname.txt file
* .xlsx file containing first name, last name, email address, and phone number added in the Contact_Info/ folder
>Refer to the examples/ directory for sample .xlsx and resulting Google CSV files (dummy information provided).


### Usage
1. Add your .xlsx file in the Contact_Info/ folder
2. Add a label in the labelname.txt file to automatically create a label in Google Contacts
3. Open command prompt in the project's root directory and run `npm install` to install required dependencies
4. Once this is finished Run `node index.js` to start the conversion process
4. The output and log files will be generated in the output/ folder assuming no errors were thrown.
5. Check the log file for any errors or missing data 
> The csv file will generate regardless of missing data if you want to leave any of the information blank.
