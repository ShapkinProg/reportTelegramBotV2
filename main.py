from __future__ import print_function

import mimetypes
from email import encoders
from email.contentmanager import maintype, subtype
from email.mime.base import MIMEBase
from calendar import monthrange
from datetime import datetime
from dateutil.parser import parse
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from tokenBot import token
import telebot
import datetime
from aiogram import Bot, types
from aiogram.dispatcher import Dispatcher
from aiogram.utils import executor
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.dispatcher import FSMContext
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import schedule
import smtplib
from a1range import A1Range
import os.path
from google.oauth2 import service_account
from googleapiclient.discovery import build
# import pandas as pd
import xlsxwriter

mode = False
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'credentials.json')
credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
SAMPLE_SPREADSHEET_ID = '1_Xmafcx0hV-MdOfkyHBF_LLGI7c3Nz3zd3SIuKsuops'
SAMPLE_RANGE_NAME = 'ОС'
service = build('sheets', 'v4', credentials=credentials).spreadsheets().values()

bot1 = telebot.TeleBot(token=token)
tb = telebot.TeleBot(token)
bot = Bot(token)
dp = Dispatcher(bot, storage=MemoryStorage())


class que(StatesGroup):
    Q = State()
    Q1 = State()
    Q2 = State()
    Q3 = State()


class queD(StatesGroup):
    Q = State()
    Q1 = State()

def parce_date(s):
    l = len(s)
    integ = []
    i = 0
    while i < l:
        s_int = ''
        a = s[i]
        while '0' <= a <= '9':
            s_int += a
            i += 1
            if i < l:
                a = s[i]
            else:
                break
        i += 1
        if s_int != '':
            integ.append(int(s_int))

    return integ


def msg(id1):
    offset = datetime.timedelta(hours=3)
    tz = datetime.timezone(offset, name='МСК')
    dt = datetime.datetime.now(tz=tz)
    resultF = service.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                         range='ОС').execute()
    resultQ = service.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                         range='Вопросы').execute()
    resultF = [x for x in resultF.get('values') if x]
    resultQ = [x for x in resultQ.get('values') if x]
    workbook = xlsxwriter.Workbook('data.xlsx')
    worksheet = workbook.add_worksheet('ОС')
    bold = workbook.add_format({'bold': True})
    count = 0
    isEmpty = True
    if len(resultF) > 1:
        for i in resultF:
            if count != 0:
                bold = None
                msg_data = parce_date(i[3])
            else:
                worksheet.write_row(col=0, row=count, data=i, cell_format=bold)
                count += 1
                continue
            if msg_data[1] == dt.month and msg_data[2] == dt.year and dt.day -msg_data[0] <=7:
                worksheet.write_row(col=0, row=count, data=i, cell_format=bold)
                count += 1
            elif msg_data[2] == dt.year and msg_data[1] == dt.month-1:
                if dt.day + monthrange(dt.year, dt.month-1)[1] - msg_data[0] <= 7:
                    worksheet.write_row(col=0, row=count, data=i, cell_format=bold)
                    count += 1

    worksheet = workbook.add_worksheet('Вопросы')
    if bold is None:
        bold = workbook.add_format({'bold': True})
    if count > 1:
        isEmpty = False
    count = 0
    if len(resultQ) > 1:
        for i in resultQ:
            if count != 0:
                bold = None
                msg_data = parce_date(i[2])
            else:
                worksheet.write_row(col=0, row=count, data=i, cell_format=bold)
                count += 1
                continue
            if msg_data[1] == dt.month and msg_data[2] == dt.year and dt.day -msg_data[0] <=7:
                worksheet.write_row(col=0, row=count, data=i, cell_format=bold)
                count += 1
            elif msg_data[2] == dt.year and msg_data[1] == dt.month-1:
                if dt.day + monthrange(dt.year, dt.month-1)[1] - msg_data[0] <= 7:
                    worksheet.write_row(col=0, row=count, data=i, cell_format=bold)
                    count += 1
    if count > 1:
        isEmpty = False
    workbook.close()
    if not isEmpty:
        file = open('data.xlsx', 'rb')
        tb.send_document(id1, file)
        tb.send_message(id1, "Также отправил его вам на почту")
        tb.send_message(id1, "Для того чтобы продолжить пользоваться ботом нажми /start")
        msg = MIMEMultipart()
        msg['From'] = "ReportBotDenga@yandex.ru"  # "ReportBotDenga@outlook.com"
        msg['To'] = "eshliapnikova@dengabank.ru"  # "eshliapnikova@dengabank.ru"
        msg[
            'Subject'] = "Автоматический отчёт данных за неделю из бота"  # tema + "   Автор сообщения: " + str(name_chel)
        filepath = "data.xlsx"  # Имя файла в абсолютном или относительном формате
        filename = os.path.basename(filepath)
        if os.path.isfile(filepath):  # Если файл существует
            ctype, encoding = mimetypes.guess_type(filepath)  # Определяем тип файла на основе его расширения
            if ctype is None or encoding is not None:  # Если тип файла не определяется
                ctype = 'application/octet-stream'  # Будем использовать общий тип
            maintype, subtype = ctype.split('/', 1)  # Получаем тип и подтип
            with open(filepath, 'rb') as fp:
                file = MIMEBase(maintype, subtype)  # Используем общий MIME-тип
                file.set_payload(fp.read())  # Добавляем содержимое общего типа (полезную нагрузку)
                fp.close()
            encoders.encode_base64(file)  # Содержимое должно кодироваться как Base64
            file.add_header('Content-Disposition', 'attachment', filename=filename)  # Добавляем заголовки
            msg.attach(file)  # Присоединяем файл к сообщению
        server = smtplib.SMTP(host='smtp.yandex.ru', port=587)
        server.starttls()
        server.login(msg['From'], "yrngmlzkoswfyneg")
        server.send_message(msg)
        server.quit()
        print("mail send, ura!")
    else:
        tb.send_message(id1, "Изменений за последнюю неделю не наблюдаю")
        tb.send_message(id1, "Для того чтобы продолжить пользоваться ботом нажми /start")


def telegram_bot(token):
    b1 = KeyboardButton('Жалоба')
    b2 = KeyboardButton('Предложение')
    kb_client = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb_client.row(b1, b2)

    @dp.message_handler(commands='start')
    async def start_message(message: types.message):
        await bot.send_message(message.from_user.id, "Если вам нужно заполнить форму нажмите /feedback \n\n"
                                                     "Eсли вы хотите поулчить файл нажмите  /docs\n\n"
                                                     "Eсли же вы хотите задать вопрос нажмите /question\n\n")

    @dp.message_handler(commands='info')
    async def info(message: types.message):
        await bot.send_message(message.from_user.id, "Идёт формирвоание отчёта...")
        msg(message.from_user.id)

    @dp.message_handler(commands='docs')
    async def start_message(message: types.message):
        root = BASE_DIR + '/files'
        files = os.listdir(root)
        kb_docs = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        for i in files:
            kb_docs.add(KeyboardButton(i))
        await bot.send_message(message.from_user.id, "Выберите юр.лицо:",
                               reply_markup=kb_docs)
        await queD.Q.set()

    @dp.message_handler(commands='feedback')
    async def startAnswer(message: types.message):
        await bot.send_message(message.from_user.id, "Для того, чтобы отправить обратную связь, выбери тему сообщения ",
                               reply_markup=kb_client)
        global mode
        mode = False
        await que.Q.set()

    @dp.message_handler(commands='question')
    async def startQuestion(message: types.message):
        await bot.send_message(message.from_user.id, "Для того, чтобы задать вопрос сначала представьтесь, "
                                                     "введи своё имя и фамилию:")
        global mode
        mode = True
        await que.Q1.set()

    @dp.message_handler(state=queD.Q)
    async def answerD0_q1(message: types.Message, state: FSMContext):
        root = BASE_DIR + '/files'
        files = os.listdir(root)
        try:
            files = os.listdir(root + '/' + message.text)
        except:
            await bot.send_message(message.from_user.id, 'Нет такого файла(', reply_markup=ReplyKeyboardRemove())
            await state.finish()
            await bot.send_message(message.from_user.id, "Для того чтобы продолжить пользоваться ботом нажми /start")
            return
        await state.update_data(answer0=message.text)
        kb_docs = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        for i in files:
            kb_docs.add(KeyboardButton(i))
        await bot.send_message(message.from_user.id, "Выберите нужный файл:",
                               reply_markup=kb_docs)
        await queD.next()

    @dp.message_handler(state=queD.Q1)
    async def answerD1_q1(message: types.Message, state: FSMContext):
        data = await state.get_data()
        root = BASE_DIR + '/files' + '/' + data.get("answer0") + '/' + message.text
        try:
            file = open(root, 'rb')
        except:
            await bot.send_message(message.from_user.id, 'Нет такого файла(', reply_markup=ReplyKeyboardRemove())
            await state.finish()
            await bot.send_message(message.from_user.id, "Для того чтобы продолжить пользоваться ботом нажми /start")
            return
        await bot.send_message(message.from_user.id, 'Вот ваш файл:', reply_markup=ReplyKeyboardRemove())
        tb.send_document(message.from_user.id, file)
        await bot.send_message(message.from_user.id, "Для того чтобы продолжить пользоваться ботом нажми /start")
        await state.finish()

    @dp.message_handler(state=que.Q)
    async def answer0_q1(message: types.Message, state: FSMContext):
        b_tema1 = KeyboardButton('Анонимно')
        kb_client_tema = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
        kb_client_tema.row(b_tema1)
        answer = message.text
        await state.update_data(answer0=answer)
        await bot.send_message(message.from_user.id, "Теперь представься, введи своё имя, фамилию если нужно:",
                               reply_markup=kb_client_tema)
        await que.next()

    @dp.message_handler(state=que.Q1)
    async def Fio(message: types.Message, state: FSMContext):
        answer = message.text
        await state.update_data(answer1=answer)
        await bot.send_message(message.from_user.id, "Напишите ваш отдел и город:", reply_markup=ReplyKeyboardRemove())
        await que.next()

    @dp.message_handler(state=que.Q2)
    async def answer1_q1(message: types.Message, state: FSMContext):
        answer = message.text
        await state.update_data(answer3=answer)
        if mode is not True:
            await bot.send_message(message.from_user.id, "Напиши текст обращения:", reply_markup=ReplyKeyboardRemove())
        else:
            await bot.send_message(message.from_user.id, "Напиши ваш вопрос:", reply_markup=ReplyKeyboardRemove())
        await que.next()

    @dp.message_handler(state=que.Q3)
    async def answer2_q1(message: types.message, state: FSMContext):
        data = await state.get_data()
        answer0 = data.get("answer0")
        answer1 = data.get("answer1")
        answer2 = message.text
        await send_mail(answer1, answer2, answer0,data.get("answer3"), message)
        await state.finish()
        await bot.send_message(message.from_user.id, "Спасибо за твою обратную связь, всё передам куда нужно!😉")
        await bot.send_message(message.from_user.id, "Для того чтобы продолжить пользоваться ботом нажми /start")

    @dp.message_handler()
    async def info_send(message: types.message):
        await bot.send_message(message.from_user.id, "Если вам нужно заполнить форму нажмите /feedback \n\n"
                                                     "Eсли вы хотите поулчить файл нажмите  /docs\n\n"
                                                     "Eсли же вы хотите задать вопрос нажмите /question\n\n")

    executor.start_polling(dp, skip_updates=True)
    # bot1.polling()


async def send_mail(name_chel, bodyMes, tema, city, message):
    global SAMPLE_RANGE_NAME
    if mode is not True:
        SAMPLE_RANGE_NAME = 'ОС'
    else:
        SAMPLE_RANGE_NAME = 'Вопросы'
    name_chel = name_chel + ' ' + city
    data = service.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                         range=SAMPLE_RANGE_NAME).execute()
    if len(data['values']) > 1:
        data['values'].pop(0)
        res = [x for x in data.get('values') if x]
        if len(res) != len(data['values']):
            for i in range(len(data['values'])):
                if mode is not True:
                    res.append(['', '', '', ''])
                else:
                    res.append(['', '', ''])
        arr = {'values': res}
        range1 = A1Range.create_a1range_from_list(SAMPLE_RANGE_NAME, 3, 1, arr.get('values')).format()
        response = service.update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                  range=range1,
                                  valueInputOption='USER_ENTERED',
                                  body=arr).execute()
    offset = datetime.timedelta(hours=3)
    tz = datetime.timezone(offset, name='МСК')
    dt = datetime.datetime.now(tz=tz)
    if mode is not True:
        array = {'values': [[name_chel, tema, bodyMes, dt.strftime("%d-%m-%Y %H:%M")]]}
    else:
        array = {'values': [[name_chel, bodyMes, dt.strftime("%d-%m-%Y %H:%M")]]}
    range_ = A1Range.create_a1range_from_list(SAMPLE_RANGE_NAME, 2, 1, array['values']).format()
    response = service.update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                              range=range_,
                              valueInputOption='USER_ENTERED',
                              body=array).execute()


if __name__ == '__main__':
    telegram_bot(token)