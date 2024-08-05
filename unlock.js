const fs = require('fs');
const os = require('os');
const path = require('path');
const AdmZip = require('adm-zip');
const xml2js = require('xml2js');

const UNLOCKED_FILE_SUFFIX = '_odemčeno';

// Check the filename, that should be passed as argument (e.g. "node unlock.js file.xlsx")
if (process.argv.length < 3) {
    console.error("Error: No filename provided");
    process.exit(1);
}

let file = process.argv[2];

if (path.extname(file) !== '.xlsx') {
    console.error("Error: File is not of type .xlsx");
    process.exit(1);
}

// Extract the zip file to a temporary directory
let zip = new AdmZip(file);
let tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'unlocked-'));
zip.extractAllTo(tempDir, true);

const parser = new xml2js.Parser();
const builder = new xml2js.Builder();

// Remove the sheetProtection tag from all sheets
fs.readdirSync(path.join(tempDir, 'xl/worksheets')).forEach(function(filename) {
    if (filename.startsWith('sheet') && path.extname(filename) === '.xml') {
        let xml = fs.readFileSync(path.join(tempDir, 'xl/worksheets', filename), 'utf8');
        parser.parseString(xml, (err, result) => {
            if (err) {
                console.error(`Error parsing XML: ${err}`);
                return;
            }
            if (result.worksheet.hasOwnProperty('sheetProtection')) {
                delete result.worksheet.sheetProtection;
                let newXml = builder.buildObject(result);
                fs.writeFileSync(path.join(tempDir, 'xl/worksheets', filename), newXml, 'utf8');
            }
        });
    }
});

// Write the modified files back to a new zip file
let newZip = new AdmZip();
newZip.addLocalFolder(tempDir);
let newFile = path.join(path.dirname(file), path.basename(file, '.xlsx') + `${UNLOCKED_FILE_SUFFIX}.xlsx`);
newZip.writeZip(newFile);

console.log(`Successfully written ${newFile}`);
process.exit(0);