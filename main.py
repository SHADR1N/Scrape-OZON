import telebot
from telebot import types
import peewee
import time 
import configparser
import requests
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import xlsxwriter

import threading
import yadisk
import os


your = ''
me = ''

y = yadisk.YaDisk(token=your)

db = peewee.SqliteDatabase('ozon.db')
class BaseModel(peewee.Model):
    class Meta:
        database = db


class Users(BaseModel):
    USERID = peewee.TextField( unique = True)
    _select_ = peewee.TextField( default = '' )


    @classmethod
    def get_row(cls, USERID):
        return cls.get(USERID == USERID)

    @classmethod
    def row_exists(cls, USERID):
        query = cls().select().where(cls.USERID == USERID)
        return query.exists()

    @classmethod
    def creat_row(cls, USERID):
        user, created = cls.get_or_create(USERID=USERID)


class NewWaybill(BaseModel):
    USERID = peewee.IntegerField( )
    NameWaybill = peewee.TextField( default = '' )
    Status = peewee.TextField( default = 'Edit' )


    @classmethod
    def get_row(cls, USERID):
        return cls.get(USERID == USERID)

    @classmethod
    def row_exists(cls, USERID):
        query = cls().select().where(cls.USERID == USERID)
        return query.exists()

    @classmethod
    def creat_row(cls, USERID, NameWaybill, Status):
        user, created = cls.get_or_create(USERID=USERID, NameWaybill = NameWaybill, Status = Status)


class dictURL(BaseModel):
    USERID = peewee.IntegerField( )
    NameWaybill = peewee.TextField( default = '' )
    URL = peewee.TextField( default = '' )
    price_first = peewee.TextField( default = '' )
    price_second = peewee.TextField( default = '' )
    name = peewee.TextField( default = '' )
    number = peewee.IntegerField( default = 0 )
    code = peewee.IntegerField( default = 0 )


    @classmethod
    def get_row(cls, USERID):
        return cls.get(USERID == USERID)

    @classmethod
    def row_exists(cls, USERID, URL, NameWaybill):
        query = cls().select().where(cls.USERID == USERID, cls.URL == URL,cls.NameWaybill == NameWaybill)
        return query.exists()

    @classmethod
    def creat_row(cls, USERID, NameWaybill, URL):
        user, created = cls.get_or_create(USERID=USERID, URL = URL, NameWaybill = NameWaybill)





db.create_tables([dictURL])
db.create_tables([NewWaybill])
db.create_tables([Users])





config = configparser.ConfigParser()
config.read("config.ini")


token = config['config']['token']
bot = telebot.TeleBot(token)




@bot.message_handler(commands=["start"])
def start(message):

    USERID = message.chat.id

    if not Users.row_exists(USERID):
        Users.creat_row(USERID)

    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    keyboard.add(*[types.KeyboardButton(name) for name in ['Создать накладную']])
    
    if NewWaybill.select().where(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit'):
        keyboard.add(*[types.KeyboardButton(name) for name in ['Сохранить накладную']])


    bot.send_message(message.chat.id, text = f'''Используй кнопки для навигации ⤵️''',reply_markup=keyboard, parse_mode="Html")  


def NEW_await(message):

    USERID = message.chat.id
    NameWaybill = message.text
    Status = 'Edit'
    NewWaybill.creat_row(USERID, NameWaybill, Status)
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    keyboard.add(*[types.KeyboardButton(name) for name in ['Создать накладную']])
    
    if NewWaybill.select().where(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit'):
        keyboard.add(*[types.KeyboardButton(name) for name in ['Сохранить накладную']])

    bot.send_message(USERID, f'Создана накладаная: {NameWaybill}',reply_markup=keyboard, parse_mode="Html")


@bot.message_handler(content_types=["text"])
def key(message):
    USERID = message.chat.id

    if message.text == 'Сохранить накладную':

        if not NewWaybill.select().where(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit'):
            keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
            keyboard.add(*[types.KeyboardButton(name) for name in ['Создать накладную']])
            
            if NewWaybill.select().where(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit'):
                keyboard.add(*[types.KeyboardButton(name) for name in ['Сохранить накладную']])

            bot.send_message(USERID, 'У вас нет созданных накладных для сохранения. Создайте новую!',reply_markup=keyboard, parse_mode="Html")
            return


        NameWaybill = NewWaybill.get(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit').NameWaybill
        check = dictURL.select().where(dictURL.USERID == USERID, dictURL.NameWaybill == NameWaybill, dictURL.price_second == '')
        if check:
            bot.send_message(USERID, 'В некоторых товарах не проставлена цена. Накладная не сохранена. Ожидайте в скором времени вся инфомрация прогрузится.')
            return

        w = NewWaybill.get(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit').NameWaybill
        s = NewWaybill.get(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit')
        s.Status = 'Upload'
        s.save()

        workbook = xlsxwriter.Workbook(f'{w}.xlsx')
        worksheet = workbook.add_worksheet()

        
        f = dictURL.select().where(dictURL.NameWaybill == w, dictURL.USERID ==USERID)
        x = []
        for el in f:

            s = [ el.code, el.name, el.name[:100], el.price_first, el.price_second, el.number, el.URL ]
            x.append( s ) 


        row = 0
        col = 0

        red = workbook.add_format()
        red.set_bg_color('red')

        worksheet.write(row, col,     'Код')
        worksheet.write(row, col + 1, 'Название')
        worksheet.write(row, col + 2, 'Название (100)')
        worksheet.write(row, col + 3, 'Со скидкой')
        worksheet.write(row, col + 4, 'Цена')
        worksheet.write(row, col + 5, 'Количество')
        worksheet.write(row, col + 6, 'Ссылка')
        row += 1
        for item in x:
            if item[4] == '0':
                worksheet.write(row, col,     str(item[0]),red)
                worksheet.write(row, col + 1, str(item[1]),red)
                worksheet.write(row, col + 2, str(item[2]),red)
                worksheet.write(row, col + 3, str(item[3]),red)
                worksheet.write(row, col + 4, str(item[4]),red)
                worksheet.write(row, col + 5, str(item[5]),red)
                worksheet.write(row, col + 6, str(item[6]),red)  
            else:
                worksheet.write(row, col,     str(item[0]))
                worksheet.write(row, col + 1, str(item[1]))
                worksheet.write(row, col + 2, str(item[2]))
                worksheet.write(row, col + 3, str(item[3]))
                worksheet.write(row, col + 4, str(item[4]))
                worksheet.write(row, col + 5, str(item[5]))
                worksheet.write(row, col + 6, str(item[6]))
            row += 1

        workbook.close()

        try:
            y.upload(f'{w}.xlsx', f'/{w}.xlsx')
            bot.send_message(USERID, f'Накладная создана и отправлена!\nИмя файла: {w}.xlsx')
        except:
            bot.send_message(USERID, f'Файл с таким названием уже есть. Накладная не может быть сохранена.')
        

        path = os.path.join(os.path.abspath(os.path.dirname(__file__)), f'{w}.xlsx')
        os.remove(path)

    if message.text == 'Создать накладную':
        if not NewWaybill.select().where(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit'):
            sent = bot.send_message(USERID, 'Введите название накладной:')
            bot.register_next_step_handler(sent, NEW_await)

        else:
            name = NewWaybill.get(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit').NameWaybill
            keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
            
            if NewWaybill.select().where(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit'):
                keyboard.add(*[types.KeyboardButton(name) for name in ['Сохранить накладную']])

            keyboard.add(*[types.KeyboardButton(name) for name in [f'Удалить накладную']])

            bot.send_message(USERID, f'У вас есть не сохраненная накладная, для создания новой сохраните предыдущую!\nИмя накладной: {name}',reply_markup=keyboard, parse_mode="Html")

    if message.text == 'Удалить накладную':
        for one in NewWaybill.select().where(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit'):
            one.delete_instance()
            name = NewWaybill.get(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit').NameWaybill
            for two in dictURL.select().where(dictURL.USERID == USERID, dictURL.NameWaybill == name):
                two.delete_instance()


        keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
        keyboard.add(*[types.KeyboardButton(name) for name in ['Создать накладную']])
        
        if NewWaybill.select().where(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit'):
            keyboard.add(*[types.KeyboardButton(name) for name in ['Сохранить накладную']])

        bot.send_message(USERID, 'Накладаная удалена. Теперь можете создать новую!',reply_markup=keyboard, parse_mode="Html")



    if 'ozon.ru' in message.text:
        URL = message.text
        URL = 'http'+str(URL.split('http')[1])
        if NewWaybill.select().where(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit'):
            NameWaybill = NewWaybill.get(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit').NameWaybill


            if not dictURL.row_exists(USERID, URL, NameWaybill):
                dictURL.creat_row(USERID, NameWaybill, URL)

            U = Users.get(Users.USERID == USERID)
            U._select_ = URL
            U.save()

            new = dictURL.get(dictURL.USERID == USERID, dictURL.URL == URL, dictURL.NameWaybill == NameWaybill)
            new.URL = URL
            new.price_first = ''
            new.price_second = ''
            new.name = ''
            new.number = 0
            new.save()
            sent = bot.send_message(USERID, f'Добавлен новый товар...\nВведите количество:')
            bot.register_next_step_handler(sent, TotalItem)
            return


        else:


            keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
            keyboard.add(*[types.KeyboardButton(name) for name in ['Создать накладную']])
            
            if NewWaybill.select().where(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit'):
                keyboard.add(*[types.KeyboardButton(name) for name in ['Сохранить накладную']])
        
            bot.send_message(USERID, 'У вас нет активных накладных. Пожалуйста создайте новую накладную.',reply_markup=keyboard, parse_mode="Html")

def TotalItem(message):
    total = message.text
    USERID = message.chat.id

    if total.isdigit() == True:

        new = dictURL.get(dictURL.USERID == USERID, dictURL.URL == Users.get(Users.USERID == USERID)._select_, dictURL.NameWaybill == NewWaybill.get(NewWaybill.USERID == USERID, NewWaybill.Status == 'Edit').NameWaybill)
        new.number = total
        new.save()

        bot.send_message(USERID, f'Добавлено количество: {total}')

    else:
        sent = bot.send_message(USERID, f'количество введено не правильно!\nВведите количество:')
        bot.register_next_step_handler(sent, TotalItem)


def get_info():
    chrome_options = Options()  
    chrome_options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.90 Safari/537.36')
    chrome_options.add_argument("--disable-javascript")
    chrome_options.add_argument("user-data-dir=selenium")
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(executable_path='./chromedriver.exe', chrome_options=chrome_options)
    driver.get('https://www.ozon.ru/')
    while True:
        if driver.find_elements_by_xpath('//header[@itemscope="itemscope"]'):
            break

    zxc = 0
    while True:
        if not driver.find_elements_by_xpath('//header[@itemscope="itemscope"]'):
            if zxc == 0:    
                bot.send_message('1529375723', 'Решите каптчу')
                zxc += 1
            continue


        zxc = 0
        for one in NewWaybill.select().where(NewWaybill.Status == 'Edit'):
            NameWaybill = one.NameWaybill


            for URL in dictURL.select().where(dictURL.price_second == '', dictURL.NameWaybill == NameWaybill):

                driver.get(URL.URL)
                time.sleep(4)
                if driver.find_elements_by_xpath('//h1[@class="b3a8"]'):
                    name = driver.find_element_by_xpath('//h1[@class="b3a8"]').text

                else:
                    name = 'ERROR'

                if driver.find_elements_by_xpath('//span[@class="c8q7 c8q8"]'):
                    price_first = driver.find_element_by_xpath('//span[@class="c8q7 c8q8"]').text
                    price_second = driver.find_element_by_xpath('//span[@class="c8r"]').text

                    price_first = price_first.replace('₽', '').strip()
                    price_second = price_second.replace('₽', '').strip()


                elif driver.find_elements_by_xpath('//span[@class="c8q7 c8q9"]'):
                    price_second = driver.find_element_by_xpath('//span[@class="c8q7 c8q9"]').text
                    
                    price_second = price_second.replace('₽', '').strip()
                    price_first = price_second

                elif driver.find_elements_by_xpath('//span[@class="c8q7"]'):
                    price_second = driver.find_element_by_xpath('//span[@class="c8q7"]').text
                    price_second = price_second.replace('₽', '').strip()
                    price_first = price_second

                elif driver.find_elements_by_xpath('//span[@class="c2h5"]'):
                    price_second = driver.find_element_by_xpath('//span[@class="c2h5"]').text
                   

                    price_second = price_second.replace('₽', '').strip()
                    price_first = price_second
                    
                elif driver.find_elements_by_xpath('//span[@class="c2h5 c2h7"]'):
                    price_second = driver.find_element_by_xpath('//span[@class="c2h5 c2h7"]').text

                    price_first = price_second.replace('₽', '').strip()
                    price_second = price_second.replace('₽', '').strip()


                elif driver.find_elements_by_xpath('//span[@class="c2h5 c2h6"]'):
                    price_second = driver.find_element_by_xpath('//span[@class="c2h5 c2h6"]').text
                    price_first = driver.find_element_by_xpath('//span[@class="c2h8"]').text
                    
                    price_first = price_first.replace('₽', '').strip()
                    price_second = price_second.replace('₽', '').strip()

                else:
                    price_first = '0'
                    price_second = '0'

                if driver.find_elements_by_xpath('//span[@class="b2d7 b2d9"]'):
                    code = driver.find_element_by_xpath('//span[@class="b2d7 b2d9"]').text
                    code = code.split('Код товара:')[1].strip()

                else:
                    code = 'ERROR'

                if name == 'ERROR':
                    URL.delete_instance()
                    continue


                w = dictURL.get( dictURL.price_first == '', dictURL.URL == URL.URL, dictURL.NameWaybill == NameWaybill )
                w.price_first = price_first
                w.code = code
                w.price_second = price_second
                w.name = name
                w.save()

            time.sleep(2)



t = threading.Thread(target=get_info)
t.start()


if __name__ == '__main__':
    bot.polling(none_stop=True)
