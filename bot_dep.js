const TelegramBot = require('node-telegram-bot-api')
const fs = require('fs')
const token = '5502505923:AAFOqf5wABh0U5xLVKisS_WMF0bY6bpFvV8'
const bot = new TelegramBot(token, {polling: true})


let XLSX = require('xlsx')
let listJson = {
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

bot.onText(/\/start/, msg => {
    bot.sendMessage(msg.chat.id,'Я можу сказати хто йде їбашити як чорт в наряді')
})
bot.onText(/\/doc1/, msg => {
    bot.sendDocument(msg.chat.id, './duty.xlsx')
})
bot.onText(/\/duty/, msg => {
    let listJson = JSON.parse(fs.readFileSync('readedFile.json'))
    let minCount = Math.min(listJson.table[0].coundOfDuty, listJson.table[1].coundOfDuty, listJson.table[2].coundOfDuty, listJson.table[3].coundOfDuty, listJson.table[4].coundOfDuty)
    for (let i = 0; i < listJson.table.length; i++) {
        if (listJson.table[i].coundOfDuty === minCount) {
            bot.sendMessage(msg.chat.id, 'Йде їбашити курсант '+ listJson.table[i].name + ', кількість нарядів : ' + listJson.table[i].coundOfDuty);
            listJson.table[i].coundOfDuty += 1
            fs.writeFile("readedFile.json", JSON.stringify(listJson, null, 4), (err) => {
                if (err) {
                    console.error(err);
                    return;
                }
                console.log("File has been created");
            });
            return
        }
    }
})
bot.onText(/\/doc2/, msg => {
})
