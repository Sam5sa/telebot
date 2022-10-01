import telebot
from telebot import types
from test_xlsx import form_req, get_req, init_exls_file ,delete_blank_rows ,unike_test, add_client, show_all_clients, get_rate, write_all_txt, write_row_txt

bot = telebot.TeleBot('....................')
pas = 1103

workers = ['I','mihey','masha','irina']
workers_id = [1111111, 1111111, 1111111, 1111111]

new_client =['ID твоего билета:', 'тип', 'статус', 'твой телефон:', 'твоя электронная почта:', 'Твоё ФИО:', 'дата', 'промоутер', 'какой подарок хочешь?','город','цена твоего билета: (в рублях)']
new_client_mask =['ID твоего билета:', 'тип', 'статус', 'твой телефон:', 'твоя электронная почта:', 'Твоё ФИО:', 'дата', 'промоутер', 'какой подарок хочешь?','город','цена твоего билета: (в рублях)']
confurm = 99

excel_file,client_sheet,mich_sheet,manya_sheet,artem_sheet = init_exls_file('client.xlsx','all','Михей','Маня','Артём')
delete_blank_rows(sheet_object=client_sheet)



#bot.send_message(720311671, 'gogo power ranger')

@bot.message_handler(commands=['start'])
def start(message):
    mess = f'Привет, {message.from_user.first_name} {message.from_user.last_name}. Добро пожаловать в театр n\n\nМы находимся не только в z. Выбери город чтобы фильтровать список представлений.\n\n Также ты можешь узнать о нас подробнее на нашем сайте.'
    site_or_city = types.InlineKeyboardMarkup()
    button1 = types.InlineKeyboardButton("узнать о нас больше", callback_data='site')
    button2 = types.InlineKeyboardButton("выбрать город", callback_data='city')
    site_or_city.add(button1,button2)
    bot.send_message(message.chat.id, mess, parse_mode='html', reply_markup = site_or_city)

'''
@bot.message_handler(commands=['add'])
def add(message):
    add_client(new_client_mask,sheet_object=client_sheet,excel_file=excel_file)
    f = open(r"client.xlsx","rb")
    bot.send_document(message.chat.id, f)
    f.close()
'''

@bot.message_handler(commands=['help'])
def help(message):
    help_markup = types.InlineKeyboardMarkup()
    button1 = types.InlineKeyboardButton("не помню пароль от госуслуг", callback_data='not_rememb_pass')
    help_markup.add(button1)
    bot.send_message(message.chat.id, 'С кокой проблемой мне помочь?', reply_markup = help_markup)

@bot.message_handler(commands=['get_req'])
def rate(message):
    rate = types.InlineKeyboardMarkup()
    button1 = types.InlineKeyboardButton("эксель файл", callback_data='exls')
    button2 = types.InlineKeyboardButton("сообщением", callback_data='mess')
    rate.add(button1,button2)
    form_req(sheet_object=client_sheet,excel_file=excel_file,mich_sheet = mich_sheet,manya_sheet = manya_sheet, artem_sheet = artem_sheet)
    if (message.from_user.id in workers_id): bot.send_message(message.chat.id,f"{message.from_user.first_name} {message.from_user.last_name}, как вам их отправить?", reply_markup=rate)

@bot.message_handler(commands=['get_rate'])
def ratep(message):
    if (message.from_user.id in workers_id): bot.send_message(message.chat.id,get_rate(sheet_object=client_sheet))

@bot.message_handler()
def get_text(message):
    global confurm
    if (confurm == 0):
        new_client[confurm] = str(message.text)
        confurm = 3
        bot.send_message(message.chat.id, f'{new_client_mask[confurm]}')
        return
    if (confurm == 3):
        new_client[confurm] = str(message.text)
        confurm = 4
        bot.send_message(message.chat.id, f'{new_client_mask[confurm]}')
        return
    if (confurm == 4):
        new_client[confurm] = str(message.text)
        confurm = 5
        bot.send_message(message.chat.id, f'{new_client_mask[confurm]}')
        return
    if (confurm == 5):
        new_client[confurm] = str(message.text)
        confurm = 10
        bot.send_message(message.chat.id, f'{new_client_mask[confurm]}')
        return
    if (confurm == 10):
        new_client[confurm] = str(message.text)
        confurm = 8
        bot.send_message(message.chat.id, f'{new_client_mask[confurm]}')
        return
    if (confurm == 7):
        new_client[confurm] = str(message.text)
        confurm = 20
        bot.send_message(message.chat.id,"Для завершения регистрации необходимо ввести уникальный код. Тебе его скажет твой промоутер")
        return
    if (confurm == 8):
        new_client[confurm] = str(message.text)
        bot.send_message(message.chat.id, 'Подтверди свои данные')
        bot.send_message(message.chat.id, f'{new_client[0]} | {new_client[3]} | {new_client[4]} | {new_client[5]} | {new_client[10]} | {new_client[8]}')
        rate = types.InlineKeyboardMarkup()
        button1 = types.InlineKeyboardButton("всё верно", callback_data='check')
        button2 = types.InlineKeyboardButton("ввести данные заново", callback_data='confurm')
        rate.add(button1,button2)
        bot.send_message(message.chat.id,"Если ты ошибся при наборе данных выбери 'ввести данные заново'.", reply_markup=rate)
        return
    if (confurm == 20):
        if (str(message.text) == str(int(new_client[0]) + pas)) and (unike_test(new_client, sheet_object=client_sheet) == 0):
            confurm = 69
            add_client(new_client,sheet_object=client_sheet,excel_file=excel_file)
            #write_row_txt(sheet_object=client_sheet)
            bot.send_message(message.chat.id,"Поздравляем! Регистрация успешно завершена")
        else:
            markup = types.InlineKeyboardMarkup()
            markup.add(types.InlineKeyboardButton("перейти", url='url'))
            bot.send_message(message.chat.id, "Произоша ошибка регистрации( Для получения помощи обратитесь к нашему менеджеру)", reply_markup = markup)
        return



'''
@bot.message_handler(commands=['send_req1'])
def instrucrion(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=3, one_time_keyboard=True)
    website = types.KeyboardButton("эксель файл")
    start = types.KeyboardButton("сообщением")
    markup.add(website,start)
    bot.send_message(message.chat.id, f"{message.from_user.first_name} {message.from_user.last_name} как вам их отправить?", reply_markup=markup)
'''


@bot.callback_query_handler(func=lambda call: True)
def callback_inline(call):
    try:
        if call.message:
            #f = open("C:\Users\Пользователь\Desktop\бот\client.xlsx","rb")
            if call.data == "exls":
                f = open(r"client.xlsx","rb")
                bot.send_document(call.message.chat.id, f)
                f.close()
            if call.data == "mess":
                if (call.message.chat.id == workers_id[0]): 
                    bot.send_message(call.message.chat.id, get_req(1,mich_sheet = mich_sheet,manya_sheet = manya_sheet, artem_sheet = artem_sheet))
                    bot.send_message(call.message.chat.id, get_req(2,mich_sheet = mich_sheet,manya_sheet = manya_sheet, artem_sheet = artem_sheet))
                    bot.send_message(call.message.chat.id, get_req(3,mich_sheet = mich_sheet,manya_sheet = manya_sheet, artem_sheet = artem_sheet))

                if (call.message.from_user.id == workers_id[1]): bot.send_message(call.message.chat.id, get_req(1,mich_sheet = mich_sheet,manya_sheet = manya_sheet, artem_sheet = artem_sheet))
                if (call.message.from_user.id == workers_id[2]): bot.send_message(call.message.chat.id, get_req(2,mich_sheet = mich_sheet,manya_sheet = manya_sheet, artem_sheet = artem_sheet))
                if (call.message.from_user.id == workers_id[3]): bot.send_message(call.message.chat.id, get_req(3,mich_sheet = mich_sheet,manya_sheet = manya_sheet, artem_sheet = artem_sheet))

            if call.data == "site":
                markup = types.InlineKeyboardMarkup()
                markup.add(types.InlineKeyboardButton("перейти", url='url'))
                bot.send_message(call.message.chat.id, "чтобы попасть на наш сайт нажми 'перейти'", reply_markup=markup)
            if call.data == "city":
                city_name = types.InlineKeyboardMarkup()
                buttonS = types.InlineKeyboardButton("Самара", callback_data='samara')
                #buttonN = types.InlineKeyboardButton("Нижний Новгород", callback_data='nijniy')
                #buttonK = types.InlineKeyboardButton("Казань", callback_data='kazan')
                #city_name.add(buttonS,buttonN,buttonK)
                city_name.add(buttonS)
                bot.send_message(call.message.chat.id, "Твой город:", parse_mode='html', reply_markup = city_name)
            if call.data == "samara":
                new_client[9] = 'samara'
                pay_type = types.InlineKeyboardMarkup()
                button1 = types.InlineKeyboardButton("Банковской картой", callback_data='bank')
                button2 = types.InlineKeyboardButton("Пушкинской картой", callback_data='pushka')
                pay_type.add(button1,button2)
                bot.send_message(call.message.chat.id, "Выбери преподчительный способ оплаты.", parse_mode='html', reply_markup = pay_type)
            if call.data == "nijniy":
                #new_client[9] = 'nijniy'
                pay_type = types.InlineKeyboardMarkup()
                button1 = types.InlineKeyboardButton("Банковской картой", callback_data='bank')
                button2 = types.InlineKeyboardButton("Пушкинской картой", callback_data='pushka')
                pay_type.add(button1,button2)
                bot.send_message(call.message.chat.id, "Выбери преподчительный способ оплаты.", parse_mode='html', reply_markup = pay_type)
            if call.data == "kazan":
                #new_client[9] = 'kazan'
                pay_type = types.InlineKeyboardMarkup()
                button1 = types.InlineKeyboardButton("Банковской картой", callback_data='bank')
                button2 = types.InlineKeyboardButton("Пушкинской картой", callback_data='pushka')
                pay_type.add(button1,button2)
                bot.send_message(call.message.chat.id, "Выбери преподчительный способ оплаты.", parse_mode='html', reply_markup = pay_type)
            if call.data == 'bank':
                bot.send_message(call.message.chat.id, "Для покупки билета по банковской карте перейди на сайт по ссылке. \n Нажми 'Купить билеты' и следуй инструкции на сайте")
                markup = types.InlineKeyboardMarkup()
                markup.add(types.InlineKeyboardButton("перейти", url='url'))
                bot.send_message(call.message.chat.id, "Сайт покупки билета", reply_markup=markup)
            if call.data == 'pushka':
                pay_type = types.InlineKeyboardMarkup()
                button1 = types.InlineKeyboardButton("Да", callback_data='est_pus')
                button2 = types.InlineKeyboardButton("Нет", callback_data='net_pus')
                pay_type.add(button1,button2)
                bot.send_message(call.message.chat.id, "Установленно ли у вас приложение 'Госуслуги культура'.", parse_mode='html', reply_markup = pay_type)
            if call.data == 'est_pus':
                bot.send_message(call.message.chat.id, "Для покупки билета по пушкинской карте перейди на сайт по ссылке. \n Нажми 'Купить билеты' и следуй инструкции на сайте")
                bot.send_message(call.message.chat.id, "ВАЖНО!!! В способе оплаты выбери пушкинскую карту.\nДля этого на сайте нажми на Пушкина, как на картинке ниже")
                photo = open(r'puskin.png', 'rb')
                bot.send_photo(call.message.chat.id, photo)
                markup = types.InlineKeyboardMarkup()
                markup.add(types.InlineKeyboardButton("перейти", url='url'))
                bot.send_message(call.message.chat.id, "Сайт покупки билета", reply_markup=markup)
                markup = types.InlineKeyboardMarkup()
                markup.add(types.InlineKeyboardButton("подтвердить", callback_data='confurm'))
                bot.send_message(call.message.chat.id, "После приобритения билета нажми 'подтвердить'",reply_markup=markup)

            if call.data == 'net_pus':
                markup = types.InlineKeyboardMarkup()
                button1 = types.InlineKeyboardButton("App store", url = 'https://apps.apple.com/ru/app/%D0%B3%D0%BE%D1%81%D1%83%D1%81%D0%BB%D1%83%D0%B3%D0%B8-%D0%BA%D1%83%D0%BB%D1%8C%D1%82%D1%83%D1%80%D0%B0/id1581979387')
                button2 = types.InlineKeyboardButton("Google play", url = 'https://play.google.com/store/apps/details?id=ru.gosuslugi.culture')
                markup.add(button1,button2)
                bot.send_message(call.message.chat.id, "Для получения пушкинской карты необходимо установить приложение 'Госуслуги культра' \n сделать это можно по ссылкам", reply_markup=markup)
                bot.send_message(call.message.chat.id, "После установки и регистрации в приложении пушкинская карта оформится в течении пары минут")
                bot.send_message(call.message.chat.id, "Для покупки билета по пушкинской карте перейди на сайт по ссылке. \n Нажми 'Купить билеты' и следуй инструкции на сайте")
                bot.send_message(call.message.chat.id, "ВАЖНО!!! В способе оплаты выбери пушкинскую карту.\nДля этого на сайте нажми на Пушкина, как на картинке ниже")
                photo = open(r'puskin.png', 'rb')
                bot.send_photo(call.message.chat.id, photo)
                markup1 = types.InlineKeyboardMarkup()
                markup1.add(types.InlineKeyboardButton("перейти", url='url'))
                bot.send_message(call.message.chat.id, "Сайт покупки билета", reply_markup=markup1)
                markup = types.InlineKeyboardMarkup()
                markup.add(types.InlineKeyboardButton("подтвердить", callback_data='confurm'))
                bot.send_message(call.message.chat.id, "После приобритения билета нажми 'подтвердить'",reply_markup=markup)

            if call.data == 'confurm':
                global confurm
                bot.send_message(call.message.chat.id, "Великолепно! Осталось совсем немного. \n\n Осталось ввести данные, необходимые для выдачи тебе подарка)")
                confurm = 0
                bot.send_message(call.message.chat.id, f'{new_client_mask[confurm]}')
            if call.data == 'check':
                bot.send_message(call.message.chat.id, "Почти готово... Осталось всего пару вопросов)")
                pay_type = types.InlineKeyboardMarkup()
                button1 = types.InlineKeyboardButton("промоутер", callback_data='promouter')
                button2 = types.InlineKeyboardButton("от знакомых", callback_data='sarafan')
                button3 = types.InlineKeyboardButton("реклама на сайтах, мессенджерах", callback_data='reklama')
                button4 = types.InlineKeyboardButton("другое", callback_data='other')
                pay_type.add(button1,button2,button3,button4)
                bot.send_message(call.message.chat.id, "Расскажи пожалуйста, как ты узнал(а) о нас.)", reply_markup = pay_type)
            if call.data in ['reklama','other']:
                new_client[7] = call.data
                confurm = 20
                markup = types.InlineKeyboardMarkup()
                markup.add(types.InlineKeyboardButton("перейти", url='url'))
                bot.send_message(call.message.chat.id, "Для завершения регистрации необходимо ввести уникальный код. Для этого свяжись с менеджером по ссылке)", reply_markup = markup)
            if call.data == 'sarafan':
                confurm = 7
                bot.send_message(call.message.chat.id, "Введи ФИ знакомого, рассказавшего вам о тетаре")
            if call.data == 'promouter':
                confurm = 7
                bot.send_message(call.message.chat.id, "Введи ФИ своего промоутера")
            if call.data == 'not_rememb_pass': 
                msg1 = "Откройте форму восстановления доступа, нажав на кнопку Войти и указав, что вы не знаете пароль.\n\n"
                msg2 = "Укажите номер телефона или e-mail, привязанный к аккаунту, или выберите из списка любой документ.\n\n"
                msg3 = "В последнем случае система самостоятельно отыщет адрес почты, который вы указывали в момент регистрации.\n\n"
                msg4 = "Ознакомьтесь с условиями использования сервиса и подтвердите, что действие запрашивает реальный человек.\n\n"
                msg5 = "В случае восстановления личного кабинета по номеру телефона дождитесь СМС-сообщение с проверочным кодом, внесите его в поле подтверждения, а после переходите к созданию нового пароля.\n\n"
                msg6 = "При обнулении доступа с использованием электронного ящика система направит письмо со ссылкой на форму активации нового ключа.\n\n"
                msg7 = "Пользователей с подтвержденной учеткой дополнительно попросят ввести номер паспорта, ИНН или СНИЛС, чтобы подтвердить личность."
                bot.send_message(call.message.chat.id, msg1+msg2+msg3+msg4+msg5+msg6+msg7)

    except Exception as e:
        print(repr(e))


bot.polling(none_stop=True)