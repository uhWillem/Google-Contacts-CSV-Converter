# Google-Contacts-CSV-Converter

This project is meant to convert any list of names, phone numbers and email to the *google contact CSV format* so they can be easily imported. 
Currently the keywords it checks in excel are in *dutch* so if you want to use this you'll have to modify the code a bit (or the Excel file). 

## Requirements
*Node.js installed
*Must add a label in \Contact_Info\labelname.txt.
*Must add a .xlsx file in \Contact_Info\ containg first name, last name, email address and phone number.
>Examples of the .xlsx and resulting google CSV can be found in \examples\ (dummy info)


### Usage
1. Begin by adding your .xlsx file in the Contact_Info 
2. Add a label in the labelname.txt folder, this will automatically create a label in google contacts with these users. 
3. Open cmd in the project's root directory and type `npm install` to install the dependencies required for this project
4. Once this is finished type `node index.js` to start the conversion.
4. The output and logs file will have generated and placed inside the output folder assuming there were no errors thrown.
5. You can analyze the logs file for any issues, if any data is missing it will tell you there. 
