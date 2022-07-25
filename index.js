const express = require('express');
const { google } = require('googleapis');
const app = express();
const { GoogleAuth } = require('google-auth-library');

//DATA THAT WILL COME FROM BUBBLE
var new_spreadsheet_name = 'Testfile kopija';


app.post("/sheetGenerator", async(req, res) => {
    var query = req.query;

    //authorization
    const auth = new google.auth.GoogleAuth({
        keyFile: "credentials.json",
        scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'],
    });

    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: "v4", auth: client });
    const drive = google.drive({ version: "v3", auth: client });


    var spreadsheetId = "18Dxbp27atU7oNaZG-6tUz5WW97mWQO7YZIyV_gEG0CM";

    var newSpreadsheet = await CopyFile(query.new_spreadsheet_name, auth, drive);

    var valueInputOption = 'USER_ENTERED';
    var range = [
        'A1',
        'A2',
        'A3',
        'B3'
    ]
    await updateValues(googleSheets, newSpreadsheet.data.id, range[0], valueInputOption, query.value1);
    await updateValues(googleSheets, newSpreadsheet.data.id, range[1], valueInputOption, query.value2);
    await updateValues(googleSheets, newSpreadsheet.data.id, range[2], valueInputOption, query.value3);
    await updateValues(googleSheets, newSpreadsheet.data.id, range[3], valueInputOption, query.value4);

    res.send(newSpreadsheet.data.id);
});





async function CopyFile(newTitle, auth, drive) {

    let resource = {
        name: newTitle,
        parents: ['11_nAGevISte6oNXdgj4G_dfweQ3fQeBX'] //folder to upload to on Shared Drive
    }

    var x = drive.files.copy({
        fileId: '18Dxbp27atU7oNaZG-6tUz5WW97mWQO7YZIyV_gEG0CM', //file to be copied
        driveId: 'root',
        includeItemsFromAllDrives: true,
        corpora: 'drive',
        supportsAllDrives: true,
        resource: resource
    });

    x.then((value) => {
        console.log(value.data.id);
    });

    return x;
}






async function updateValues(googleSheets, spreadsheetId, range, valueInputOption, value) {

    let values = [
        [
            value // Cell values ...
        ],
    ];
    const resource = {
        values,
    };
    const request = {
        spreadsheetId,
        range,
        valueInputOption,
        resource,
    };
    try {
        const result = await googleSheets.spreadsheets.values.update(request);
        console.log('%d cells updated.', result.data.updatedCells);
        return result;
    } catch (err) { throw err; }
}



//FUNCTIONS

async function metaDataFromID(googleSheets, auth, spreadsheetId) {
    return googleSheets.spreadsheets.get({
        auth,
        spreadsheetId,
    });
}

async function getRows(spreadsheetId, googleSheets, auth, sheetName) {
    return googleSheets.spreadsheets.values.get({
        auth,
        spreadsheetId,
        range: sheetName,
    });
}

async function createSheet(title, googleSheets, drive) {

    const resource = {
        properties: {
            title,

        },
    };
    var newSheet = await googleSheets.spreadsheets.create({
        resource,
        fields: 'spreadsheetId',
    });
    var fileId = await newSheet.data.spreadsheetId;
    givePermissions(drive, fileId);
    return fileId;
}

async function givePermissions(drive, fileId) {
    return drive.permissions.create({
        resource: {
            type: "user",
            role: "writer",
            emailAddress: "rokasbalt02@gmail.com",
        },
        fileId: fileId,
        fields: "id",
    });
}

async function getFilesFromDrive(drive) {
    return drive.files.list({
        "corpora": "user",
        "includeItemsFromAllDrives": true,
        "supportsAllDrives": true
    });
}