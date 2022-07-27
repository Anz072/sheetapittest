const express = require('express');
const { google } = require('googleapis');
const { GoogleAuth } = require('google-auth-library');
const port = process.env.PORT || '3000';
const fs = require('fs');
const bodyParser = require('body-parser');
const { response } = require('express');
const { Console } = require('console');
const jsonParser = bodyParser.json();
const app = express();
require('dotenv').config();
const pdf2base64 = require('pdf-to-base64');
const credPull = process.env.CREDENTIALS;
const creds = JSON.parse(credPull);


const valueInputOption = 'USER_ENTERED';
const range = [
    'A1',
    'A2',
    'A3'
]

app.use(function(req, res, next) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, PUT, PATCH, DELETE');
    res.setHeader('Access-Control-Allow-Headers', 'X-Requested-With,content-type');
    res.setHeader('Access-Control-Allow-Credentials', true);
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
        credentials: creds,
        scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'],
    });


    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: "v4", auth: client });
    const drive = google.drive({ version: "v3", auth: client });
    var spreadsheetId = '0';

    if (req.body.sheetID == '0') {
        //create new file
        var newSpreadsheetMain = await CopyFile(req.body.SheetName, auth, drive);
        spreadsheetId = newSpreadsheetMain.data.id;
    } else {
        spreadsheetId = req.body.sheetID;
    }


    await updateValues(googleSheets, spreadsheetId, range[0], valueInputOption, req.body.values[0]);
    await updateValues(googleSheets, spreadsheetId, range[1], valueInputOption, req.body.values[1]);
    await updateValues(googleSheets, spreadsheetId, range[2], valueInputOption, req.body.values[2]);

    //download a file
    var file = await downloadFile(spreadsheetId, drive);
    var sheetMeta = await metaDataFromID(googleSheets, auth, spreadsheetId);
    var downloadName = sheetMeta.data.properties.title;


    var readerFeed = await writer(file, downloadName);
    var reader = fs.readFileSync(downloadName + ".pdf", { encoding: 'base64' });

    console.log(reader);
    var what = 'dab';
    var se = await pdf2base64(downloadName + '.pdf')
        .then(
            (response) => {
                what = response;
            }
        )
        .catch(
            (error) => {
                console.log(error);
            }
        );


    res.send(what);
});

app.listen(port, () => {
    console.log(`API is listening at http://localhost:${port}`)
})


//FUNCTIONS
async function writer(file, downloadName) {

    var w = fs.createWriteStream(downloadName + '.pdf');
    var m = file.data.pipe(w);
    return m;
}

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
    console.log(spreadsheetId);
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
        return result;
    } catch (err) { throw err; }
}



async function metaDataFromID(googleSheets, auth, spreadsheetId) {
    return googleSheets.spreadsheets.get({
        auth,
        spreadsheetId,
    });
}