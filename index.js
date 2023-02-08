const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const folderPath = path.resolve(__dirname, "contact_info");
const labelpath = path.resolve(__dirname, "contact_info/labelname.txt")
const outputFilePath = path.resolve(__dirname, "output/output.csv");
const logsFilePath = path.resolve(__dirname, "output/logs.txt")

let output = [];
let outputStr = "";
let logs = [];
var data = [];

// adds a new line to the logs array
function addNewLine() {
    logs.push("");
}

// gets the value from the labelname.txt file
function getLabelName() {
    return new Promise((resolve, reject) => {
        fs.readFile(labelpath, 'utf8', (err, labelname) => {
            if (err) {
                reject(new Error(`An error occurred while reading the labelname.txt file: ${err}`));
            };

            if (labelname.toString() == " ") {
                reject(new Error("Please fill in the labelname.txt file with a label name"));
            };

            if (labelname.toString().length > 80) {
                reject(new Error("Please set a labelname in contact_info/labelname.txt OR your label name is too long, please shorten it to 80 characters or less"));
            };
            resolve(labelname);
        });
    });
}


async function main() {
    const labelName = await getLabelName();
    logs.push(`Label name: ${labelName}`);

    // reading files in contact_info directory
    fs.readdir(folderPath, (err, files) => {

        if (err) {
            throw new Error(`An error occurred while reading the contact_info folder: ${err}`);
        }

        logs.push(`Found files: ${files} in contact_info directory`);
        addNewLine();

        // checks if there's 2 files in the directory and one must be a .txt file and the other a .xlsx file
        if (files.length >= 2 && (files[0].toString().endsWith(".txt") == true || files[1].toString().endsWith(".txt") == true) && (files[0].toString().endsWith(".xlsx") == true || files[1].toString().endsWith(".xlsx") == true)) {

            // getting the path of the xlsx file
            // if the first file is a .xlsx file, it will be used, if not the second file will be used.
            // We know one of them is .xslsx because of the check above
            const filePath = path.join(folderPath, files[0].endsWith('.xlsx') ? files[0] : files[1]);


            // extracting the data from the xlsx file
            const xlsxfile = XLSX.readFile(filePath);
            const sheetName = xlsxfile.SheetNames[0];
            const worksheet = xlsxfile.Sheets[sheetName];
            data = XLSX.utils.sheet_to_json(worksheet);

            // google contact csv format
            output.push("Name,Given Name,Additional Name,Family Name,Yomi Name,Given Name Yomi,Additional Name Yomi,Family Name Yomi,Name Prefix,Name Suffix,Initials,Nickname,Short Name,Maiden Name,Birthday,Gender,Location,Billing Information,Directory Server,Mileage,Occupation,Hobby,Sensitivity,Priority,Subject,Notes,Language,Photo,Group Membership,E-mail 1 - Type,E-mail 1 - Value,E-mail 2 - Type,E-mail 2 - Value,Phone 1 - Type,Phone 1 - Value,Website 1 - Type,Website 1 - Value")

            logs.push("FORMAT: {Index} - {Voornaam} {Familienaam} | {Email} | {GSM}")
            addNewLine();

            // creating the output array with the final google csv data
            data.forEach((row, index) => {
                console.log(index)
                let log = `${index} - ${row.Voornaam} ${row.Familienaam} | ${row.Privemailadres} | ${row['GSM-nummer']}`;
                let newline = false;
                logs.push(log)

                if (row.Voornaam == undefined) {
                    logs.push(`Missing {VOORNAAM} at index: ${index}`)
                    newline = true;
                }

                if (row.Familienaam == undefined) {
                    logs.push(`Missing {FAMILIENAAM} at index: ${index}`)
                    newline = true;
                }

                if (row.Privemailadres == undefined) {
                    logs.push(`Missing {EMAIL} at index: ${index}`)
                    newline = true;
                }

                if (row['GSM-nummer'] == undefined) {
                    logs.push(`Missing {GSM-NUMMER} at index: ${index} `)
                    newline = true;

                } else if (row['GSM-nummer'].length != 13) {
                    logs.push(`Invalid {GSM-NUMMER} at index: ${index} PLEASE CORRECT MANUALLY OR REMOVE`)
                }

                if (newline) {
                    addNewLine();
                }

                // pushing the csv row to the output array
                output.push(`${row.Voornaam + " " + row.Familienaam},${row.Voornaam},,${row.Familienaam},,,,,,,,,,,,,,,,,,,,,,,,,${labelName} ::: * myContacts,,${row.Privemailadres},,,Mobile,${row['GSM-nummer']},,`)
            });

            // joining the array to a string
            outputStr = output.join('\n')

            // calling the write functions
            writeLogs(logs);
            writeOutput(outputStr);
        } else {
            throw new Error("There is more than two files in the contact_info folder or they are invalid, please ensure there's 2 files, one being .xlsx and one .txt");
        }

    });
}


// logs file is created if it does not exist, if it does exist it is wiped clean
function writeLogs(logs) {
    fs.writeFile(logsFilePath, logs.join('\n'), {
        encoding: 'utf8',
        flag: 'w'
    }, (err) => {
        if (err) {
            throw new Error(`An error occurred while writing the logs file: ${err}`);
        }
        console.log(`The logs file was successfully written to ${logsFilePath}`);
    });
}


// creating and writing the output file to the output folder
function writeOutput(outputStr) {
    fs.writeFile(outputFilePath, outputStr, {
        encoding: 'utf8',
        flag: 'w'
    }, (err) => {
        if (err) {
            throw new Error(`An error occurred while writing the output file: ${err}`);
        } else {
            console.log(`The output file was successfully written to ${outputFilePath}`);
        }
    });
}

// executing the main function that converts the xlsx file to a google csv file
main();