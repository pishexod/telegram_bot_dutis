
const{
    Telegraf
} = require('telegraf')

const bot = new Telegraf('5502505923:AAFOqf5wABh0U5xLVKisS_WMF0bY6bpFvV8')
const fs = require('fs')
let filename = './duty.xlsx';
var Excel = require('exceljs');
var workbook_new = new Excel.Workbook();
let XLSX = require('xlsx')
let listJson = {
    table: []
}

bot.start( (ctx) => ctx.reply( 'Я можу сказати хто йде їбашити як чорт в наряді'))
bot.command('doc', (ctx)=>{
    ctx.replyWithDocument({source:"./duty.xlsx"})
})
bot.command('duty', (ctx) => {
    let workbook = XLSX.readFile(filename)
    let worksheet = workbook.Sheets[workbook.SheetNames[0]]
    for (let i = 2; i < 7; i++) {
        if (i !== 1) {
            listJson.table.push({
                name: worksheet[`${String.fromCharCode(65)}${i}`].v,
                coundOfDuty: parseInt(worksheet[`${String.fromCharCode(66)}${i}`].v)
            })
        }
    }
    let minCount = Math.min(listJson.table[0].coundOfDuty, listJson.table[1].coundOfDuty, listJson.table[2].coundOfDuty, listJson.table[3].coundOfDuty, listJson.table[4].coundOfDuty)
    for (let i = 0; i < listJson.table.length; i++) {
        if (listJson.table[i].coundOfDuty === minCount) {
            ctx.reply( 'Йде їбашити курсант ' + listJson.table[i].name + ', кількість нарядів : ' + listJson.table[i].coundOfDuty);
            listJson.table[i].coundOfDuty += 1
            workbook_new.xlsx.readFile('./duty.xlsx')
                .then(function () {
                    var worksheet = workbook_new.getWorksheet(1);
                    var row = worksheet.getRow(i + 2);
                    row.getCell(2).value = row.getCell(2).value + 1; // A5's value set to 5
                    row.commit();
                    return workbook_new.xlsx.writeFile('duty.xlsx');
                })
            return
        }
    }
})
bot.launch()