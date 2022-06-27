let XLSX = require('xlsx')
let fs = require('fs')

var listJson = {
    table: []
}
let workbook = XLSX.readFile('./duty.xlsx')
let worksheet = workbook.Sheets[workbook.SheetNames[0]]
for (let i = 2; i < 7; i++) {
    if (i !== 1) {
        listJson.table.push({
            name: worksheet[`${String.fromCharCode(65)}${i}`].v,
            coundOfDuty: parseInt(worksheet[`${String.fromCharCode(66)}${i}`].v)
        })
    }
}
fs.writeFile("readedFile.json", JSON.stringify(listJson, null, 4), (err) => {
    if (err) {
        console.error(err);
        return;
    }
    console.log("File has been created");
});