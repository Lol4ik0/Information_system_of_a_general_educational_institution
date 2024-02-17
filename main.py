import telebot
import pygsheets
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from telebot import types
from config import telebot_token, tabel, spreadsheet_link, spreadsheet_link_worksheet_title, registration_list, syllabus, admin_gmail, admin_password

# Ініціалізуємо бота
bot = telebot.TeleBot(telebot_token)

# Авторизація в Google Диску
gc = pygsheets.authorize(service_file='molten-box-386414-5430574da430.json')

# Глобальні змінні
global send_button, gmail, password, name, enter_class, lesson, number_name, grade, grade_of_client, lesson_of_grade, date_of_lesson, name_of_trimestrs, trimestr, tabel_grade
gmail = ''
password = ''
send_button = ''
name = ''
enter_class = ''
lesson = ''
teacher = ''
number_name = 0
grade = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '']
grade_of_client = []
lesson_of_grade = []
date_of_lesson = []
trimestr = ''
tabel_grade = []
name_of_trimestrs = []

# Функція для пошуку значення в аркуші таблиці
def find_value_in_worksheet(worksheet, target_value):
    cell_list = worksheet.findall(target_value)
    if not cell_list:
        return None  # Значення не знайдено
    else:
        return cell_list  # Повертаємо список знайдених комірок

# Функція для пошуку облікового запису з вказаними поштою та паролем
def find_gmail_in_worksheet(registr_sheet, target_value, password_):
    print('Працює функція find_gmail_in_worksheet()')
    for wk in registr_sheet.worksheets():
        print(wk.title)
        found_cells = find_value_in_worksheet(wk, target_value)
        print(found_cells)
        if found_cells:
            for cell in found_cells:
                row, col = cell.row, cell.col
                cell_value = wk.cell(row, col).value
                password_cell_value = wk.cell(row, col + 1).value
                if cell_value == target_value and password_cell_value == password_:
                    global name, enter_class
                    name = wk.cell(row, col - 1).value
                    enter_class = wk.title
                    print(f'target_value - {target_value}, password_ - {password_}, name - {name}, enter_class - {enter_class}')
                    return 1
                elif cell_value == target_value and password_cell_value != password_:
                    global password
                    password = "password is incorrect"
                    print('Password is incorrect')
                    return None
        else:
            print(f'In {wk.title} found_cells is empty!')
    return None

def send_b(text):
    global send_button
    send_button = text

def delete_user():
    global gmail, password, name
    gmail, password, name = '', '', ''

# Обробник команди /help
@bot.message_handler(commands=['help'])
def help(message):
    help_text = [
        'Напишіть "/start", щоб пройти реєстрацію.',
        'Напишіть "/do", щоб обрати дію.',
        'Напишіть "/help", щоб побачити інструкцію ще раз.',
        'Напишіть "/no", щоб відмовитися від чогось.'
    ]
    for i in range(len(help_text)):
        bot.send_message(message.chat.id, help_text[i], parse_mode='html')

# Обробник команди /start
@bot.message_handler(commands=['start'])
def start(message):
    send_b('start')
    bot.send_message(message.chat.id, 'Введіть вашу пошту')

# Обробник команди /do
@bot.message_handler(commands=['do'])
def do(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    b1 = types.KeyboardButton('Дізнатися оцінки з предмету')
    b2 = types.KeyboardButton('Сформувати табель')
    b3 = types.KeyboardButton("Контактні дані вчителя")
    b4 = types.KeyboardButton("Мій розклад занять")
    markup.add(b1, b2, b3, b4)
    bot.send_message(message.chat.id, 'Оберіть дію', reply_markup=markup)

    send_b('do')

# Обробник повідомлень з текстом "no"
@bot.message_handler(regexp='no')
def no(message):
    send_b('')
    bot.send_message(message.chat.id, 'Ви відмовилися від цього.')
    bot.send_message(message.chat.id, 'Введіть "/help" для того, щоб почати інструкцію ще раз!')

# Обробник повідомлень користувача
@bot.message_handler()
def get_user_text(message):
    if send_button == 'start':
        global gmail
        gmail = message.text
        bot.send_message(message.chat.id, 'Напишіть ваш пароль.')
        send_b("enter_password")

    elif send_button == "enter_password":
        global password
        password = message.text

        if gmail == admin_gmail and password == admin_password:
          sh_registration_list = gc.open_by_key(registration_list)
          sh_tabel = gc.open_by_key(tabel)
          sh_spreadsheet_link = gc.open_by_key(spreadsheet_link)

          help_text = [
              f'Токен бота: {telebot_token}',
              f'Лист рейстрацій (пошта учня та її пароль): {sh_registration_list.url}',
              f'Предмети та посилання на них: {sh_spreadsheet_link.url}',
              f'Головний приклад табеля: {sh_tabel.url}'
          ]
          for i in range(len(help_text)):
              bot.send_message(message.chat.id, help_text[i], parse_mode='html')
          return 0

        please_wait_message = bot.send_message(message.chat.id, 'Будь ласка, зачекайте, триває реєстрація!')

        # Шлях до вашого JSON-файлу зі службовими обліковими даними
        credentials_file = 'molten-box-386414-5430574da430.json'

        # Зазначте список дозволів, які вам потрібні для доступу до Google таблиць
        scope = ['https://www.googleapis.com/auth/spreadsheets', 
                 'https://spreadsheets.google.com/feeds', 
                 'https://www.googleapis.com/auth/drive']

        # Аутентифікуйте службовий обліковий запис
        credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
        client = gspread.authorize(credentials)

        registr_sheet = client.open_by_key(registration_list)
        find_gmail_in_worksheet(registr_sheet, gmail, password)

        if password == 'password is incorrect':
            bot.edit_message_text('Пароль не вірний! Напишіть "/start", щоб розпочати знову!', message.chat.id, please_wait_message.id)
            delete_user()

        elif name == '':
            print('gmail is not found')
            bot.edit_message_text('Вибачте, щось пішло не так, спробуйте знову. Напишіть "/start", щоб розпочати знову!', message.chat.id, please_wait_message.id)
            delete_user()

        else:
            bot.edit_message_text('Реєстрація пройшла успішно!', message.chat.id, please_wait_message.id)
            bot.send_message(message.chat.id, 'Напишіть "/do", щоб обрати дію!')

    elif send_button == "do":
        if message.text == 'Мій розклад занять':
            if name == '' and gmail == '' and password == '':
                bot.send_message(message.chat.id, "Ви не зареєстувалися! Напишіть '/start', щоб зареєструватися!")
                return 0

            bot.send_message(message.chat.id, "Ось посилання на розклад:")
            bot.send_message(message.chat.id, syllabus)
            bot.send_message(message.chat.id, 'Напишіть "/do", щоб обрати дію!')

            send_b('')

        if message.text == 'Дізнатися оцінки з предмету':
            if name == '' and gmail == '' and password == '':
                bot.send_message(message.chat.id, "Ви не зареєстувалися! Напишіть '/start', щоб зареєструватися!")
                return 0

            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=3)

            sh = gc.open_by_key(spreadsheet_link)
            wk = sh.worksheet_by_title(spreadsheet_link_worksheet_title)

            list_of_value = wk.get_all_values()

            for i in list_of_value:
                value = i[0]
                if value == list_of_value[0][0]: 
                  continue
                markup.add(types.KeyboardButton(value))
            bot.send_message(message.chat.id, 'Оберіть предмет', reply_markup=markup)

            send_b('lesson')

        elif message.text == 'Сформувати табель':
          if name == '' and gmail == '' and password == '':
                bot.send_message(message.chat.id, "Ви не зареєстувалися! Напишіть '/start', щоб зареєструватися!")
                return 0

          please_wait_message = bot.send_message(message.chat.id, 'Будь ласка, зачекайте!')
          loading_message = bot.send_message(message.chat.id, 'Завантаження 0%')
          # Шлях до вашого JSON-файлу зі службовими обліковими даними
          credentials_file = 'molten-box-386414-5430574da430.json'

          # Зазначте список дозволів, які вам потрібні для доступу до Google таблиць
          scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

          # Аутентифікуйте службовий обліковий запис
          credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
          client = gspread.authorize(credentials)

          # Створіть нову таблицю
          tabel_table = client.copy(tabel, 'Табель', copy_comments=False)
          print('Таблицю успішно скопійовано!')
          bot.edit_message_text('Завантаження 2%', message.chat.id, loading_message.id)

          # Отримайте посилання на створену таблицю
          tabel_table_url = tabel_table.url
          print(f"Посилання на створену таблицю: {tabel_table_url}")

          # Отримати об'єкт таблиці
          tabel_spreadsheet = client.open_by_url(tabel_table_url)

          # Змінити налаштування доступу
          tabel_spreadsheet.share('', perm_type='anyone', role="reader")

          # Заповнюємо П.І.Б. та клас
          tabel_worksheet = tabel_spreadsheet.worksheet('Лист1')
          tabel_worksheet.update_cell(28, 1, name)
          tabel_worksheet.update_cell(20, 4, enter_class)

          loading = 4
          bot.edit_message_text(f'Завантаження {loading}%', message.chat.id, loading_message.id)

          sh = gc.open_by_key(spreadsheet_link)
          wk = sh.worksheet_by_title(spreadsheet_link_worksheet_title) 

          list_of_link = wk.get_all_values()

          name_of_trimestrs.clear()
          for row in list_of_link:
            value = row[2]
            if row != list_of_link[0] and value != "":
              name_of_trimestrs.append(value)

          for i in list_of_link:            
            #Якщо i - це перша строка, то не виконувати код
            if i == list_of_link[0] or i[1] == "": 
              continue

            link = i[0]
            link_url = i[1]

            sh = gc.open_by_url(link_url)
            wk = sh.worksheet_by_title(enter_class)
            list_of_value = wk.get_all_values()

            # Search "П.І.Б."
            count = 0
            for row in list_of_value:
              if row[1] == name:
                break
              count = count + 1

            number_of_name = count

            # Search "trimestr"
            row = list_of_value[4]
            count = 2

            for name_trimestr in name_of_trimestrs:
              value = row[count]
              print(f'Search "{name_trimestr}"')

              while value != name_trimestr:
                  count = count + 1
                  value = row[count]

              number_trim = count
              tabel_grade.append(list_of_value[number_of_name][number_trim])
              print(f'"{name_trimestr}" is found, this is: {wk.get_value((number_of_name, number_trim))}')
              print(tabel_grade)

              loading = loading + 1
              bot.edit_message_text(f'Завантаження {loading}%', message.chat.id, loading_message.id)

            # Шукаємо предмет, та записуємо у нього отримані дані
            found_cells = find_value_in_worksheet(tabel_worksheet, link)
            if found_cells:
              for cell in found_cells:
                tabel_worksheet.update_cell(cell.row, (cell.col)+3, tabel_grade[0])
                tabel_worksheet.update_cell(cell.row, (cell.col)+4, tabel_grade[1])
                tabel_worksheet.update_cell(cell.row, (cell.col)+5, tabel_grade[2])
                tabel_worksheet.update_cell(cell.row, (cell.col)+6, tabel_grade[3])
            else:
              print('found_cells is empty!')

            tabel_grade.clear()


          while loading < 90:
            loading = loading + 10
            bot.edit_message_text(f'Завантаження {loading}%', message.chat.id, loading_message.id)

          bot.delete_message(message.chat.id, please_wait_message.id)
          bot.delete_message(message.chat.id, loading_message.id)

          send_b('')

          # OUTPUT
          bot.send_message(message.chat.id, f'Ось посилання на створений табель: {tabel_table_url}', parse_mode='html')
          bot.send_message(message.chat.id, 'Напишіть "/do", Щоб обрати дію!')

        elif message.text == "Контактні дані вчителя":
            if name == '' and gmail == '' and password == '':
                bot.send_message(message.chat.id, "Ви не зареєстувалися! Напишіть '/start', щоб зареєструватися!")
                return 0

            markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=3)

            sh = gc.open_by_key(spreadsheet_link)
            wk = sh.worksheet_by_title(spreadsheet_link_worksheet_title)

            list_of_value = wk.get_all_values()

            for i in list_of_value:
                value = i[0]
                if value == list_of_value[0][0]: 
                  continue
                markup.add(types.KeyboardButton(value))
            bot.send_message(message.chat.id, 'Оберіть предмет', reply_markup=markup)

            send_b('Search teacher')

    elif send_button == 'Search teacher':
        send = bot.send_message(message.chat.id, 'Будь ласка, зачекайте!')
        sh = gc.open_by_key(spreadsheet_link)
        wk = sh.worksheet_by_title(spreadsheet_link_worksheet_title) 
        print('Search lesson')

        list_of_value = wk.get_all_values()

        lesson_ = None
        for row in list_of_value:
          value = row[0]
          if value == message.text:
            lesson_ = row[1]
            print(lesson_)
            break

        sh = gc.open_by_url(lesson_)
        wk = sh.worksheet_by_title(enter_class)

        bot.edit_message_text(f'Вашого вчителя з предмету {message.text} звуть {wk.get_value((3, 2))}', message.chat.id, send.id)
        bot.send_message(message.chat.id, 'Напишіть "/do", Щоб обрати дію!')

        send_b('')


    elif send_button == "lesson":
        sh = gc.open_by_key(spreadsheet_link)
        wk = sh.worksheet_by_title(spreadsheet_link_worksheet_title) 
        print('Search lesson')

        list_of_value = wk.get_all_values()

        for row in list_of_value:
          value = row[0]
          if value == message.text:
            global lesson
            lesson = row[1]
            print(lesson)
            break

        name_of_trimestrs.clear()
        for row in list_of_value:
          value = row[2]
          if row != list_of_value[0] and value != "":
            name_of_trimestrs.append(value)

        markup = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=3)
        for i in name_of_trimestrs:
          markup.add(types.KeyboardButton(i))
        bot.send_message(message.chat.id, 'Оберіть час', reply_markup=markup)

        send_b('found_grade')

    elif send_button == "found_grade":
        global trimestr
        trimestr = message.text
        please_wait_message = bot.send_message(message.chat.id, 'Будь ласка, зачекайте!')
        loading_message = bot.send_message(message.chat.id, 'Завантаження 0%')
        print('working')

        try:
            sh = gc.open_by_url(lesson)
            wk = sh.worksheet_by_title(enter_class)
            bot.edit_message_text('Завантаження 2%', message.chat.id, loading_message.id)
            print(wk)
            list_of_value = wk.get_all_values()

            # Search "trimestr"
            count = 0
            value = list_of_value[4][count]
            bot.edit_message_text('Завантаження 3%', message.chat.id, loading_message.id)
            print('Search "trimestr"')

            while value != trimestr:
                count = count + 1
                value = list_of_value[4][count]

            number_trimestr = count

            bot.edit_message_text('Завантаження 5%', message.chat.id, loading_message.id)
            print(f'Number of "trimestr is found: {number_trimestr}')

            # Search number_of_column
            count = 0
            value = list_of_value[1][count]
            bot.edit_message_text('Завантаження 7%', message.chat.id, loading_message.id)
            print('Search number_of_column')

            while value != trimestr:
                count = count + 1
                value = list_of_value[1][count]

            number_of_column = count

            bot.edit_message_text('Завантаження 10%', message.chat.id, loading_message.id)
            print(f'Number of column is found: {number_of_column}')

            # Search "П.І.Б."
            count = 0
            bot.edit_message_text('Завантаження 12%', message.chat.id, loading_message.id)
            print('Search "П.І.Б."')

            for row in list_of_value:
              if row[1] == name:
                break
              count = count + 1

            global number_name
            number_name = count

            bot.edit_message_text('Завантаження 14%', message.chat.id, loading_message.id)
            print(f'Number of "П.І.Б." is found: {number_name}')

            # Search grade and his lesson and date            
            is_grade = True
            number_of_string = number_name + 1
            value = list_of_value[number_name][2]
            loading = 15
            bot.edit_message_text(f'Завантаження {loading}%', message.chat.id, loading_message.id)
            while number_of_column < number_trimestr:

                number_of_column = number_of_column + 1
                value = list_of_value[number_name][number_of_column]

                if value != "":
                    if (loading + 5) < 100:
                      loading = loading + 5
                      bot.edit_message_text(f'Завантаження {loading}%', message.chat.id, loading_message.id)

                    is_grade = True
                    number_of_string = number_name + 1
                    grade_of_client.append(value)
                    print(f'grade is found this is {value}')

                    while is_grade == True:
                        number_of_string = number_of_string - 1
                        value = list_of_value[number_of_string][number_of_column]
                        n = 0

                        for i in grade:
                            if value == i:
                                n = 1

                        if n == 0:
                            is_grade = False

                            if (loading + 5) < 100:
                                loading = loading + 5
                                bot.edit_message_text(f'Завантаження {loading}%', message.chat.id, loading_message.id)

                            print(f'lesson_of_grade is found this is {value}')
                            lesson_of_grade.append(value)
                            date_of_lesson.append((wk.get_value((number_of_string-2, number_of_column))))


            print(f'grade_of_client: {grade_of_client}')
            print(f'lesson_of_grade: {lesson_of_grade}')
            print(f'date_of_lesson: {date_of_lesson}')

            while loading < 100:
              loading = loading + 5
              bot.edit_message_text(f'Завантаження {loading}%', message.chat.id, loading_message.id)

            bot.delete_message(message.chat.id, please_wait_message.id)
            bot.delete_message(message.chat.id, loading_message.id)

            send_b('')

            # OUTPUT
            bot.send_message(message.chat.id, f'За {trimestr} ви маєте наступні оцінки:', parse_mode='html')
            for i in range(len(grade_of_client)):
              bot.send_message(message.chat.id, f'{date_of_lesson[i]} {lesson_of_grade[i]}: {grade_of_client[i]}.', parse_mode='html')

            bot.send_message(message.chat.id, 'Напишіть "/do", Щоб обрати дію!')


            grade_of_client.clear()
            lesson_of_grade.clear()
            date_of_lesson.clear()



        except Exception as ex:
            bot.send_message(message.chat.id, f'Exception is: {ex}')
            print(f'Exception is: {ex}')

    else:
        bot.send_message(message.chat.id, 'Я вас не зрозумів.')
        bot.send_message(message.chat.id, 'Напишіть "/help", щоб побачити інструкцію!')
        bot.send_message(message.chat.id, 'Або напишіть "/do", щоб обрати дію!')


bot.polling(none_stop=True)