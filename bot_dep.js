const {Telegraf, Markup} = require('telegraf');
require('dotenv').config()
const bot = new Telegraf(process.env.BOT_TOKEN)

let filename = './duty.xlsx';
let Excel = require('exceljs');
let XLSX = require('xlsx')
let cron = require('node-cron')
const fs = require("fs");


// cron.schedule('1 1 3 * * *', () => {
//     if (listId.table.length !== 0) {
//         for (let i = 0; i < listId.table.length; i++) {
//             bot.telegram.sendMessage(listId.table[i].id, 'Доброго вечора!')
//         }
//     }
// })
// cron.schedule('1 1 18 * * *', () =>{
//     if (listId.table.length !== 0) {
//         for (let i = 0; i < listId.table.length; i++) {
//             bot.telegram.sendMessage(listId.table[i].id, 'Вимкніть світло!')
//         }
//     }
// })
// cron.schedule('1 1 19 * * *', () =>{
//     if (listId.table.length !== 0) {
//         for (let i = 0; i < listId.table.length; i++) {
//             bot.telegram.sendMessage(listId.table[i].id, 'На добраніч!')
//         }
//     }
// })

let workbook_new = new Excel.Workbook();
let listJson = {table: []}
let listId = {table: []}

let name = ''
let index
let chatId
let messageId
let workbook = XLSX.readFile(filename)
let worksheet = workbook.Sheets[workbook.SheetNames[0]]

bot.start(async (ctx) => {
    await ctx.reply('Я можу сказати, хто йде їбашити в наряд')
    await obj.onlyStart()
})
bot.command('doc', (ctx) => {
    ctx.replyWithDocument({source: "./duty.xlsx"})
})


bot.command('duty', async (ctx) => {
    obj.findMin()
    try {
        let messageI = await ctx.reply('Йде їбашити курсант ' + listJson.table[index - 2].name + ', кількість нарядів : ' + listJson.table[index - 2].coundOfDuty, Markup.inlineKeyboard(
            [Markup.button.callback('Помилувати', 'btn_1'), Markup.button.callback('Підтвердити', 'btn_2')],
        ));
        console.log(ctx.chat.id)
        chatId = ctx.chat.id
        messageId = messageI.message_id

    } catch (e) {
        console.error(e)
    }
})

let obj = {
    async onlyStart() {
        try {
            for (let i = 2; i < 7; i++) {
                if (i !== 1) {
                    listJson.table.push({
                        name: worksheet[`${String.fromCharCode(65)}${i}`].v,
                        coundOfDuty: parseInt(worksheet[`${String.fromCharCode(66)}${i}`].v),
                    })
                }
            }
        } catch (e) {
            console.error(e)
        }
        console.log(listJson.table)
    },
    findMin() {
        try {

            let minCount = Math.min(listJson.table[0].coundOfDuty, listJson.table[1].coundOfDuty, listJson.table[2].coundOfDuty, listJson.table[3].coundOfDuty, listJson.table[4].coundOfDuty,)
            for (let i = 0; i < listJson.table.length; i++) {
                if (listJson.table[i].coundOfDuty === minCount) {
                    name = listJson.table[i].name
                    index = i + 2
                    break
                }
            }
        } catch (e) {
            console.error(e)
        }

    },
    async unable() {
        try {
            bot.telegram.deleteMessage(chatId, messageId)
            if (index < listJson.table.length + 1) {
                index += 1;
            } else if (index === listJson.table.length + 1) {
                index = 2;
            }
            let message = await bot.telegram.sendMessage(chatId, 'Йде їбашити курсант ' + listJson.table[index - 2].name + ', кількість нарядів : ' + listJson.table[index - 2].coundOfDuty, Markup.inlineKeyboard(
                [Markup.button.callback('Помилувати', 'btn_1'), Markup.button.callback('Підтвердити', 'btn_2')]
            ))
            messageId = message.message_id
        } catch (e) {
            console.error(e)
        }
    },
    async able(i) {
        try {
            bot.telegram.deleteMessage(chatId, messageId)
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
        } catch (e) {
            console.error(e)
        }
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