const {Telegraf, Markup} = require('telegraf');
require('dotenv').config()
const bot = new Telegraf(process.env.BOT_TOKEN)

let filename = './duty.xlsx';
let Excel = require('exceljs');
let workbook_new = new Excel.Workbook();
let XLSX = require('xlsx')
let listJson = {table: []}


let name = ''
let index
let chatId

bot.start((ctx) => ctx.reply('Я можу сказати хто йде їбашити як чорт в наряді'))
bot.command('doc', (ctx) => {
    ctx.replyWithDocument({source: "./duty.xlsx"})
})

let workbook = XLSX.readFile(filename)
let worksheet = workbook.Sheets[workbook.SheetNames[0]]
for (let i = 2; i < 7; i++) {
    if (i !== 1) {
        listJson.table.push({
            name: worksheet[`${String.fromCharCode(65)}${i}`].v,
            coundOfDuty: parseInt(worksheet[`${String.fromCharCode(66)}${i}`].v),
        })
    }
}

bot.command('duty', (ctx) => {
    obj.findMin()
    try {
        chatId = ctx.chat.id
        ctx.reply('Йде їбашити курсант ' + listJson.table[index - 2].name + ', кількість нарядів : ' + listJson.table[index - 2].coundOfDuty, Markup.inlineKeyboard(
            [Markup.button.callback('Помилувати', 'btn_1'), Markup.button.callback('Підтвердити', 'btn_2')]
        ));
    } catch (e) {
        console.error(e)
    }
})

let obj = {
    findMin() {
        let minCount = Math.min(listJson.table[0].coundOfDuty, listJson.table[1].coundOfDuty, listJson.table[2].coundOfDuty, listJson.table[3].coundOfDuty, listJson.table[4].coundOfDuty,)
        for (let i = 0; i < listJson.table.length; i++) {
            if (listJson.table[i].coundOfDuty === minCount) {
                name = listJson.table[i].name
                index = i + 2
                break
            }
        }

    },
    unable() {
        if (index < listJson.table.length + 1) {
            index += 1;
        } else if (index === listJson.table.length + 1) {
            index = 2;
        }
        bot.telegram.sendMessage(chatId, 'Йде їбашити курсант ' + listJson.table[index - 2].name + ', кількість нарядів : ' + listJson.table[index - 2].coundOfDuty, Markup.inlineKeyboard(
            [Markup.button.callback('Помилувати', 'btn_1'), Markup.button.callback('Підтвердити', 'btn_2')]
        ))
    },
    able(i) {

        bot.telegram.sendMessage(chatId, 'Йде їбашити курсант ' + listJson.table[i].name + ', було нарядів : ' + listJson.table[i].coundOfDuty + ', а стане : ' + (listJson.table[i].coundOfDuty + 1))
        listJson.table[i].coundOfDuty += 1
        workbook_new.xlsx.readFile('./duty.xlsx')
            .then(function () {
                var worksheet = workbook_new.getWorksheet(1);
                var row = worksheet.getRow(i + 2);
                row.getCell(2).value = listJson.table[i].coundOfDuty; // A5's value set to 5
                row.commit();
                return workbook_new.xlsx.writeFile('duty.xlsx');
            })
    }
}


bot.action('btn_1', (ctx) => {
    obj.unable()
})


bot.action('btn_2', (ctx) => {
    obj.able(index - 2)
})
bot.launch()


process.once('SIGINT', () => bot.stop('SIGINT'))
process.once('SIGTERM', () => bot.stop('SIGTERM'))