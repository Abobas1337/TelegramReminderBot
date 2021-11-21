from __future__ import print_function
import pandas as pd
import json
import csv
from google.oauth2 import service_account
import pygsheets


import httplib2

import datetime
import config
import telepot

from apiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials

import time
import schedule
from threading import Thread

import telebot
from telebot import types
bot = telebot.TeleBot('2084320484:AAFChX3RLZ035MEO69-8vco0TrUWQPTDF_U')






with open('service_account.json') as source:
    info = json.load(source)
credentials = service_account.Credentials.from_service_account_info(info)
client = pygsheets.authorize(service_account_file='service_account.json')
spreadsheet_url = "https://docs.google.com/spreadsheets/d/1CEy7qAOHZ6sUyxtVseVms0nB8GBI1nGkMqaiCCcu2Qk/edit?usp=sharing"
sheet_data = client.sheet.get('1CEy7qAOHZ6sUyxtVseVms0nB8GBI1nGkMqaiCCcu2Qk')
sheet = client.open_by_key("1CEy7qAOHZ6sUyxtVseVms0nB8GBI1nGkMqaiCCcu2Qk")
wks = sheet.worksheet_by_title('Лист1')

mat1_cells = wks.find("Математика", searchByRegex=False, matchCase=False, matchEntireCell=False, includeFormulas=False,
                      cols=(1, 1), rows=(1, 9), forceFetch=True)
mat1 = wks.get_value((mat1_cells[0].row, mat1_cells[0].col + 1))
print(mat1)


days = {"Понедельник", "Вторник", "Среда", "Четверг", "Пятница"}

@bot.message_handler(commands=['start'])
def menu(message):
	bot.reply_to(message, "Добро пожаловать! Напиши menu")


@bot.message_handler(func=lambda message: message.text.lower() == "menu" or message.text.lower() == "/menu" or
										  message.text.lower() == "меню")
def menu(message):
	markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
	itembtn1 = types.KeyboardButton('Дневник')
	markup.row(itembtn1)
	bot.send_message(message.from_user.id, "Выбери функцию:", reply_markup=markup)


@bot.message_handler(func=lambda message: message.text == "Дневник")
def week(message):
	markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
	itembtn1 = types.KeyboardButton('Понедельник')
	itembtn2 = types.KeyboardButton('Вторник')
	itembtn3 = types.KeyboardButton('Среда')
	itembtn4 = types.KeyboardButton('Четверг')
	itembtn5 = types.KeyboardButton('Пятница')
	itembtn6 = types.KeyboardButton('Меню')
	markup.row(itembtn1, itembtn2)
	markup.row(itembtn3, itembtn4, itembtn5)
	markup.row(itembtn6)
	bot.send_message(message.from_user.id, "Выбери день недели:", reply_markup=markup)

current_column = -1

@bot.message_handler(func=lambda message: message.text in days)
def homework(message):
	global current_column
	markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
	if message.text == "Понедельник":
		current_column = 1
		itembtn1 = types.KeyboardButton('1.Математика')
		itembtn2 = types.KeyboardButton('2.История')
		itembtn3 = types.KeyboardButton('3.Математика')
		itembtn4 = types.KeyboardButton('4.Физика')
		itembtn5 = types.KeyboardButton('5.Информатика')
		itembtn6 = types.KeyboardButton('6.Английский язык')
		itembtn7 = types.KeyboardButton('7.Литература')
		itembtn8 = types.KeyboardButton('8.Литература')
		itembtn9 = types.KeyboardButton('Меню')
		itembtn10 = types.KeyboardButton('Дневник')
		markup.row(itembtn1, itembtn2, itembtn3, itembtn4)
		markup.row(itembtn5, itembtn6, itembtn7, itembtn8)
		markup.row(itembtn9, itembtn10)
		bot.send_message(message.from_user.id, "Выбери Предмет:", reply_markup=markup)
	elif message.text == "Вторник":
		current_column = 3
		itembtn1 = types.KeyboardButton('1.Физическая культура')
		itembtn2 = types.KeyboardButton('2.Русский язык')
		itembtn3 = types.KeyboardButton('3.Информатика')
		itembtn4 = types.KeyboardButton('4.Физика')
		itembtn5 = types.KeyboardButton('5.Информатика')
		itembtn6 = types.KeyboardButton('6.Практикум по физике')
		itembtn7 = types.KeyboardButton('7.Индивидуальный проект')
		itembtn8 = types.KeyboardButton('Меню')
		itembtn9 = types.KeyboardButton('Дневник')
		markup.row(itembtn1, itembtn2, itembtn3, itembtn4)
		markup.row(itembtn5, itembtn6, itembtn7)
		markup.row(itembtn8, itembtn9)
		bot.send_message(message.from_user.id, "Выбери Предмет:", reply_markup=markup)
	elif message.text == "Среда":
		current_column = 5
		itembtn1 = types.KeyboardButton('1.Обществознание')
		itembtn2 = types.KeyboardButton('2.Физика')
		itembtn3 = types.KeyboardButton('3.Русский язык')
		itembtn4 = types.KeyboardButton('4.Математика')
		itembtn5 = types.KeyboardButton('5.Информатика')
		itembtn6 = types.KeyboardButton('6.Английский язык')
		itembtn7 = types.KeyboardButton('Меню')
		itembtn8 = types.KeyboardButton('Дневник')
		markup.row(itembtn1, itembtn2, itembtn3, itembtn4)
		markup.row(itembtn5, itembtn6)
		markup.row(itembtn7, itembtn8)
		bot.send_message(message.from_user.id, "Выбери Предмет:", reply_markup=markup)
	elif message.text == "Четверг":
		current_column = 7
		itembtn1 = types.KeyboardButton('1.Английский язык')
		itembtn2 = types.KeyboardButton('2.История')
		itembtn3 = types.KeyboardButton('3.Обществознание')
		itembtn4 = types.KeyboardButton('4.Математика')
		itembtn5 = types.KeyboardButton('5.Физика')
		itembtn6 = types.KeyboardButton('6.Математика')
		itembtn7 = types.KeyboardButton('7.Компьютерное моделирование')
		itembtn8 = types.KeyboardButton('Меню')
		itembtn9 = types.KeyboardButton('Дневник')
		markup.row(itembtn1, itembtn2, itembtn3, itembtn4)
		markup.row(itembtn5, itembtn6, itembtn7)
		markup.row(itembtn8, itembtn9)
		bot.send_message(message.from_user.id, "Выбери Предмет:", reply_markup=markup)
	elif message.text == "Пятница":
		current_column = 9
		itembtn1 = types.KeyboardButton('1.Математика')
		itembtn2 = types.KeyboardButton('2.Математика')
		itembtn3 = types.KeyboardButton('3.Литература')
		itembtn4 = types.KeyboardButton('4.Практикум по математике')
		itembtn5 = types.KeyboardButton('5.Физика')
		itembtn6 = types.KeyboardButton('6.Практикум по физике')
		itembtn7 = types.KeyboardButton('Меню')
		itembtn8 = types.KeyboardButton('Дневник')
		markup.row(itembtn1, itembtn2, itembtn3, itembtn4)
		markup.row(itembtn5, itembtn6)
		markup.row(itembtn7, itembtn8)
		bot.send_message(message.from_user.id, "Выбери Предмет:", reply_markup=markup)

@bot.message_handler(func=lambda message: message.text[0].isdigit())
def subject(message):
	global current_column
	if current_column == -1:
		bot.send_message(message.from_user.id, "Сначала выбери день")
		return

	subject_name = message.text
	subject_cells = wks.find(subject_name, searchByRegex=False, matchCase=False, matchEntireCell=False,
						  includeFormulas=False,
						  cols=(current_column, current_column), rows=(1, 9), forceFetch=True)

	if len(subject_cells) == 0:
		bot.send_message(message.from_user.id, "Такого предмета в этот день нет")
		return

	subj_dz = wks.get_value((subject_cells[0].row, subject_cells[0].col + 1))
	bot.send_message(message.from_user.id, "Домашнее задание: " + subj_dz)

# @bot.message_handler(func=lambda message: message.text == "Важные закладки")
# def week(message):
# 	markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
# 	itembtn1 = types.KeyboardButton('понедельник')
# 	itembtn2 = types.KeyboardButton('вторник')
# 	itembtn3 = types.KeyboardButton('среда')
# 	itembtn4 = types.KeyboardButton('четверг')
# 	itembtn5 = types.KeyboardButton('пятница')
# 	itembtn6 = types.KeyboardButton('Меню')
# 	markup.row(itembtn1, itembtn2)
# 	markup.row(itembtn3, itembtn4, itembtn5)
# 	markup.row(itembtn6)
# 	bot.send_message(message.from_user.id, "Выбери день недели:", reply_markup=markup)

@bot.message_handler(func=lambda m: True)
def echo_all(message):
	print(message.from_user.username + ":" + message.text)
	bot.reply_to(message, "Каво")


def job():
	TOKEN = '2084320484:AAFChX3RLZ035MEO69-8vco0TrUWQPTDF_U'
	client_secret_calendar = 'client_secret.json' # указываем путь к скачанному Json
	print("I'm working...")
	bot = telepot.Bot(TOKEN)
	credentials = ServiceAccountCredentials.from_json_keyfile_name(client_secret_calendar,
																   'https://www.googleapis.com/auth/calendar.readonly')
	http = credentials.authorize(httplib2.Http())
	service = discovery.build('calendar', 'v3', http=http)

	now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
	now_1day = round(time.time()) + 86400  # плюс сутки
	now_1day = datetime.datetime.fromtimestamp(now_1day).isoformat() + 'Z'

	print('Берем 100 событий')
	eventsResult = service.events().list(
		calendarId='kor.ol2005@yandex.ru', timeMin=now, timeMax=now_1day,
		maxResults=100, singleEvents=True,
		orderBy='startTime').execute()
	events = eventsResult.get('items', [])

	if not events:
		print('нет событий на ближайшие сутки')
		bot.sendMessage(-1001583873663, 'нет событий на ближайшие сутки')
	else:
		msg = '<b>События на ближайшие сутки:</b>\n'
		for event in events:
			start = event['start'].get('dateTime', event['start'].get('date'))
			print(start, ' ', event['summary'])
			if 'description' not in event:
				print('нет описания')
				ev_desc = 'нет описания'
			else:
				print(event['description'])
				ev_desc = event['description']

			ev_title = event['summary']
			cal_link = f"<a href=\"{event['htmlLink']}\">Подробнее...</a>"
			ev_start = event['start'].get('dateTime')
			print(cal_link)
			msg = msg + '%s\n%s\n%s\n%s\n\n' % (ev_title, ev_start, ev_desc, cal_link)
			print('===================================================================')
		bot.sendMessage(-1001583873663, msg, parse_mode='HTML')


def do_schedule():
	print("Sheduling...")
	schedule.every().day.at("17:27:00").do(job)
	while True:
		schedule.run_pending()
		time.sleep(1)


def main_loop():
	print("WORK STARTED)))))))")
	thread = Thread(target=do_schedule)
	thread.start()
	bot.polling(True)


if __name__ == '__main__':
	main_loop()



