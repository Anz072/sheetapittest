const express = require('express');
const { google } = require('googleapis');
const { GoogleAuth } = require('google-auth-library');
const port = process.env.PORT || '3000';
const fs = require('fs');
const bodyParser = require('body-parser');
const jsonParser = bodyParser.json();
const app = express();
require('dotenv').config();




const valueInputOption = 'USER_ENTERED';
const range = [
    'A1',
    'A2',
    'A3'
]

app.use(function(req, res, next) {

    // Website you wish to allow to connect
    res.setHeader('Access-Control-Allow-Origin', '*');

    // Request methods you wish to allow
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, PUT, PATCH, DELETE');

    // Request headers you wish to allow
    res.setHeader('Access-Control-Allow-Headers', 'X-Requested-With,content-type');

    // Set to true if you need the website to include cookies in the requests sent
    // to the API (e.g. in case you use sessions)
    res.setHeader('Access-Control-Allow-Credentials', true);

    // Pass to next layer of middleware
    next();
});

app.get('/', (req, res) => {
    res.send('This is an API test.')
})

/*  --------------------------------------------------
            SEND FROM BUBBLE IN BODY
            sheetID - id of edit sheet
            SheetName - name for the new spreadsheet
    --------------------------------------------------*/

app.post("/sheet", jsonParser, async(req, res) => {


    //authorization
    const auth = new google.auth.GoogleAuth({
        keyFile: ".env",
        scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'],
    });


    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: "v4", auth: client });
    const drive = google.drive({ version: "v3", auth: client });
    var newSpreadsheet;

    if (req.body.sheetID == '0') {
        //create new file
        var newSpreadsheetMain = await CopyFile(req.body.SheetName, auth, drive);
        newSpreadsheet = newSpreadsheetMain.data.id;
    } else {
        newSpreadsheet = req.body.sheetID;
    }

    await updateValues(googleSheets, newSpreadsheet, range[0], valueInputOption, req.body.values[0]);
    await updateValues(googleSheets, newSpreadsheet, range[1], valueInputOption, req.body.values[1]);
    await updateValues(googleSheets, newSpreadsheet, range[2], valueInputOption, req.body.values[2]);

    //download a file
    var file = await downloadFile(newSpreadsheet, drive);
    var sheetMeta = await metaDataFromID(googleSheets, auth, newSpreadsheet);
    var downloadName = sheetMeta.data.properties.title;
    var x = file.data.pipe(fs.createWriteStream('SheetPDF/' + downloadName + '.pdf'));

    var returnData = {
        sheetID: newSpreadsheet,
        downloadLINK: `https://testforgooglesheets.herokuapp.com/SheetPDF/${downloadName}.pdf`
    };
    res.send(returnData);
});

app.listen(port, () => {
    console.log(`API is listening at http://localhost:${port}`)
})


//FUNCTIONS


//file download
async function downloadFile(realFileId, drive) {
    fileId = realFileId;

    var file = drive.files.export({
        fileId: realFileId,
        alt: 'media',
        mimeType: 'application/pdf'
    }, { responseType: "stream" });

    return file;
}

//makes a copy & puts it in specified folder
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





//Updates a row in specified spreadsheet
async function updateValues(googleSheets, spreadsheetId, range, valueInputOption, value) {

    let values = [
        [
            value
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



//OTHER FUNCTIONS

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