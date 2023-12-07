import asyncio
import logging
from aiogram import types
from aiogram import Bot, Dispatcher
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.filters import Command
from Excel import *
from start import *
import time
# Необходимо подключить openpyxl
logging.basicConfig(level=logging.INFO) # Включаем логирование, чтобы не пропустить важные сообщения
bot = Bot(token="6380923818:AAGKjLup6yvNM3D3ehgsM3MeZNkgzhFMYwg") # Объект бота
storage = MemoryStorage()
dp = Dispatcher(storage=storage) # Диспетчер
paths = []
class States(StatesGroup):
    subject = State()
    group = State()
    Familiya = State()
    GoNext = State()
    CheckSubject = State()
    CheckWork = State()
    Checksubject = State()
    CheckSub = State()
    Downloadfile = State()
    Git = State()
    NewSub = State()
    ready = State()
    ReallyReady = State()

spisok = [] # В этот список добавляюстся ID телеграмма для того, чтобы бот понимал, авторизировался студент ранеее или нет. Если не авторизировался, то сюда добавляется его айди, если авторизировался, то его айди уже будет здесь

@dp.message(Command("start")) # Начало использования бота
async def PersonWhoUsing(message: types.Message):
    kb = [
        [
            types.KeyboardButton(text="Иное лицо"),
            types.KeyboardButton(text="Студент")
        ],
    ]
    keyboard = types.ReplyKeyboardMarkup(
        keyboard=kb,
        resize_keyboard=True,
    )
    await message.answer("Для работы бота необходимо уточнить, вы?", reply_markup=keyboard)

@dp.message(lambda message: message.text == "Иное лицо") # Если пользователь бота не студент, то ему доступна лишь функция проверки документа на ГОСТ
async def OtherUsing(message:types.Message, state: FSMContext):
    kb = [
        [
            types.KeyboardButton(text="Проверить документ")
        ],
    ]
    keyboard = types.ReplyKeyboardMarkup(
        keyboard=kb,
        resize_keyboard=True,
    )
    await message.answer("Доступный вам функционал:\nПроверить word документ с расширением .docx на соответствие ГОСТ", reply_markup=keyboard)

@dp.message(lambda message: message.text == "Студент" or message.text == "Повторить") # Авторизация студента по БД + занесения ТГ-айди в БД (необходимо для поиска при повторной авторизации)
async def cmd_start(message:types.Message, state: FSMContext):
    identified = False
    for i in spisok: # Проверка на то, авторизировался ли студент ранее или нет
        if i == message.from_user.id:
            identified = True
    if (identified == False):
        kb = [
            [
                types.KeyboardButton(text="Основы профессиональной деятельности"),
                types.KeyboardButton(text="Объектно-ориентированное программирование"),
                types.KeyboardButton(text="Динамические языки")
            ],
        ]
        keyboard = types.ReplyKeyboardMarkup(
            keyboard=kb,
            resize_keyboard=True,
        )
        await message.answer("Для работы бота необходимо идентифицировать вас. Введите дисциплину, которую у вас преподает Кабанов Артемий Андреевич. \nНаиболее используемые ответы:\n-Основы профессиональной деятельности\n-Объектно-ориентированное программирование\n-Динамические языки\nТакже можно использовать сокращения 'ОПД','ООП','ДЯ'", reply_markup=keyboard)
        await state.set_state(States.subject)

        @dp.message(States.subject)
        async def process_first_question(message: types.Message, state: FSMContext):

            text = (message.text).lower()
            if text == "основы профессиональной деятельности" or text == "опд":
                text = "ОПД.xlsx"
            elif text == "объектно-ориентированное программирование" or text == "ооп":
                text = "ООП.xlsx"
            elif text == "динамические языки" or text == "дя":
                text = "ДЯ.xlsx"
            await state.update_data(subject=text)
            await message.answer("Введи вашу группу из студенческого билета. Например:\n'ПИН-221'\n'ИВТ-231'\n'АТП-211'")
            await state.set_state(States.group)

        @dp.message(States.group)
        async def process_second_question(message: types.Message, state: FSMContext):
            await state.update_data(group=(message.text).upper())
            await state.set_state(States.Familiya)
            await message.reply("Введите полную фамилию и имя в именительном падеже. Например:\n'Кит Денис'\n'Фирс Данил'\n'Шляхтин Роман'\n'Гриневич Илья'")

        @dp.message(States.Familiya)
        async def process_third_question(message: types.Message, state: FSMContext):
            await message.answer("Ожидайте")

            await state.update_data(Familiya=message.text)
            data = await state.get_data()
            print(str(data['subject']))
            var = await authorization_student(data['subject'], data['group'], data['Familiya'])
            boolvar = await set_telegram_id(data['subject'], data['group'], data['Familiya'], message.from_user.id)
            print("Статус boolvar'a " + str(boolvar))
            if (var == True and boolvar == True):
                spisok.append(message.from_user.id)
                kb = [
                    [
                        types.KeyboardButton(text="Начать")
                    ],
                ]
                keyboard = types.ReplyKeyboardMarkup(
                    keyboard=kb,
                    resize_keyboard=True,
                )
                await message.answer('Вы идентифицированы, для использования функционала бота необходимо нажать "Начать"', reply_markup=keyboard)

            else: # Студент не авторизирован
                kb = [
                    [
                        types.KeyboardButton(text="Повторить")
                    ],
                ]
                keyboard = types.ReplyKeyboardMarkup(
                    keyboard=kb,
                    resize_keyboard=True,
                )
                await message.answer('Вы не идентифицированы. Повторите попытку, перепроверив данные, которые вы вводите. \n\nВ случае возникновения проблем обратитесь к преподавателю', reply_markup=keyboard)

    else:
        kb = [
            [
                types.KeyboardButton(text="Начать")
            ],
        ]
        keyboard = types.ReplyKeyboardMarkup(
            keyboard=kb,
            resize_keyboard=True,
        )
        await message.answer('Введи "Начать", чтобы продолжить',reply_markup=keyboard)

@dp.message(lambda message: message.text == "начать" or message.text == "Начать" or message.text == "Назад") # Основное меню
async def GoNext(message: types.Message):
    kb = [
        [
            types.KeyboardButton(text="Проверить документ"),
            types.KeyboardButton(text="Задания и баллы"),
            types.KeyboardButton(text="Материалы"),
            types.KeyboardButton(text="Профиль")
        ],
    ]
    keyboard = types.ReplyKeyboardMarkup(
        keyboard=kb,
        resize_keyboard=True
    )
    await message.answer("Доступный функционал:\n- Проверить word документ с расширением .docx на соответствие ГОСТ\n- Узнать состояние работ и количество баллов по дисциплине\n- Дополнительные материалы для изучения\n- Профиль студента", reply_markup=keyboard)

@dp.message(lambda message: message.text == "Проверить документ" or message.text == "Прислать ещё") # Проверка документа на ГОСТ
async def docs(message: types.Message, state: FSMContext):
        await message.answer("Пришите word документ с расширением .docx")
        await state.set_state(States.Downloadfile)

        @dp.message(States.Downloadfile)
        async def process_download(message: types.Message, state: FSMContext,  content_types=types.ContentType.DOCUMENT):
            await message.answer("Идет процесс обработки документа")
            try:
                file = await bot.get_file(message.document.file_id)
                if file.endswith('.docx'): # Если прислали не .docx файл, то удаляет его и сообщает об ошибочном присланном файле
                    await bot.download_file(file.file_path, str(message.document.file_name))
                    filecheck() # Проверка файла на соответствие ГОСТ
                    await bot.send_document(message.chat.id, types.FSInputFile(f"C:/Users/Desktop/PycharmProjects/pythonProject4/" + str(message.document.file_name)))
                    await message.answer("В следующем документе указаны недочеты, которые необходимо исправить\n\nЕсли вы не обнаружили в файле примечаний с указанием о необходимости что-то исправить, то ваш документ соответствует ГОСТ")
                    os.remove(FileName())
                else:
                    os.remove(FileName())
                    kb = [
                        [
                            types.KeyboardButton(text="Прислать ещё")
                        ],
                    ]
                    keyboard = types.ReplyKeyboardMarkup(
                        keyboard=kb,
                        resize_keyboard=True
                    )
                    await message.answer("Необходим документ с расширением .docx\nПовторите попытку",reply_markup=keyboard)
            except:
                kb = [
                    [
                        types.KeyboardButton(text="Прислать ещё")
                    ],
                ]
                keyboard = types.ReplyKeyboardMarkup(
                    keyboard=kb,
                    resize_keyboard=True
                )
                await message.answer(
                    "Ошибка документа, попробуйте снова",
                    reply_markup=keyboard)


@dp.message(lambda message: message.text == "Задания и баллы")
async def works(message: types.Message,  state: FSMContext):
    kb = [ # Необходимо для уточнениня по какой дисциплине необходимо узнать баллы и работы
        [
            types.KeyboardButton(text="Основы профессиональной деятельности"),
            types.KeyboardButton(text="Объектно-ориентированное программирование"),
            types.KeyboardButton(text="Динамические языки")
        ],
    ]
    keyboard = types.ReplyKeyboardMarkup(
        keyboard=kb,
        resize_keyboard=True,
    )
    await message.answer("Введите название дисциплины, по которой хотите узнать статус заданий и баллы", reply_markup=keyboard)
    await state.set_state(States.CheckSubject)

    @dp.message(States.CheckSubject)
    async def Sub(message: types.Message, state: FSMContext):

        data = await state.get_data()
        await message.answer("Ожидайте")
        Student = data['Familiya']
        Group = data['group']
        text = (message.text).lower()
        if text == "основы профессиональной деятельности" or text == "опд":
            text = "ОПД.xlsx"
        elif text == "объектно-ориентированное программирование" or text == "ооп":
            text = "ООП.xlsx"
        elif text == "динамические языки" or text == "дя":
            text = "ДЯ.xlsx"
        kb = [
            [
                types.KeyboardButton(text="Назад")
            ],
        ]
        keyboard = types.ReplyKeyboardMarkup(
            keyboard=kb,
            resize_keyboard=True,
        )
        status = await check_status(text,Group,Student)
        await message.answer(status, reply_markup=keyboard)

@dp.message(lambda message: message.text == "Материалы")
async def materials(message: types.Message):

    await message.reply("здесь ссылка на материалы")

@dp.message(lambda message: message.text == "Загрузить GITHUB")
async def ADDGit(message: types.Message, state: FSMContext):
    kb = [
        [
            types.KeyboardButton(text="Основы профессиональной деятельности"),
            types.KeyboardButton(text="Объектно-ориентированное программирование"),
            types.KeyboardButton(text="Динамические языки")
        ],
    ]
    keyboard = types.ReplyKeyboardMarkup(
        keyboard=kb,
        resize_keyboard=True,
    )
    await message.answer("Введите название дисциплины", reply_markup=keyboard)
    await state.set_state(States.NewSub)

    @dp.message(States.NewSub)
    async def ADDGits(message: types.Message, state: FSMContext):
        text = (message.text).lower()
        if text == "основы профессиональной деятельности" or text == "опд":
            text = "ОПД.xlsx"
        elif text == "объектно-ориентированное программирование" or text == "ооп":
            text = "ООП.xlsx"
        elif text == "динамические языки" or text == "дя":
            text = "ДЯ.xlsx"
        await state.update_data(NewSub=text)
        await message.reply("Пришлите ссылку на Github. Например:\nhttps://github.com/Ifpatr1ck\n\n В случае наличия ошибок в ссылке, она не будет добавлена")
        await state.set_state(States.Git)

    @dp.message(States.Git)
    async def ADDGits(message: types.Message, state: FSMContext):
        data = await state.get_data()
        Student = data['Familiya']
        Group = data['group']
        await message.answer("Ожидайте")
        text = await change_github(data['NewSub'],Group, Student,message.text)
        kb = [
            [
                types.KeyboardButton(text="Назад")
            ],
        ]
        keyboard = types.ReplyKeyboardMarkup(
            keyboard=kb,
            resize_keyboard=True,
        )
        await message.answer(text, reply_markup=keyboard)

@dp.message(lambda message: message.text == "Профиль" or message.text == "Вернуться")
async def infomenu(message: types.Message,state: FSMContext):
    data = await state.get_data()
    textOfMessage = data['subject']
    Student = data['Familiya']
    Group = data['group']
    if textOfMessage == "ОПД.xlsx":
        textOfMessage = "Основы профессиональной деятельности"
    elif textOfMessage == "ООП.xlsx":
        textOfMessage = "Объектно-ориентированное программирование"
    elif textOfMessage == "ДЯ.xlsx":
        textOfMessage = "Динамические языки"

    await message.reply("Информация об используемом аккаунте:\nИмя студента: "+ Student + '\nГруппа студента: ' + Group + '\nИзучаемая дисциплина: ' + textOfMessage)
    kb = [
        [
            types.KeyboardButton(text="Назад"),
            types.KeyboardButton(text="Загрузить GITHUB"),
            types.KeyboardButton(text="Сообщить о загруженной работе")
        ],
    ]

    keyboard = types.ReplyKeyboardMarkup(
        keyboard=kb,
        resize_keyboard=True,
    )
    await message.answer("Выберите функцию", reply_markup=keyboard)

@dp.message(lambda message: message.text == "Сообщить о загруженной работе" or message.text == "Сообщить ещё")
async def ElseReady(message: types.Message,state: FSMContext):
    time.sleep(5)
    await message.answer("Данный раздел предназначен для того, чтоб вы могли указать о том, какую работу необходимо проверить преподавателю")
    kb = [
        [
            types.KeyboardButton(text="Основы профессиональной деятельности"),
            types.KeyboardButton(text="Объектно-ориентированное программирование"),
            types.KeyboardButton(text="Динамические языки")
        ],
    ]
    keyboard = types.ReplyKeyboardMarkup(
        keyboard=kb,
        resize_keyboard=True,
    )
    await message.answer("Выберите дисциплину", reply_markup=keyboard)
    await state.set_state(States.ready)

    @dp.message(States.ready)
    async def Ready(message: types.Message, state: FSMContext):
        await message.answer("Ожидайте")
        text = (message.text).lower()
        if text == "основы профессиональной деятельности" or text == "опд":
            text = "ОПД.xlsx"
        elif text == "объектно-ориентированное программирование" or text == "ооп":
            text = "ООП.xlsx"
        elif text == "динамические языки" or text == "дя":
            text = "ДЯ.xlsx"
        await state.update_data(ready=text)
        data = await state.get_data()
        Group = data['group']
        try:
            countoflab = await kolvo_lab(text,Group)
            text = "Всего лабораторных работ:\n"
            for i in range(1, int(countoflab) + 1):
                text += "ЛР" + str(i) + "\n"
            await message.answer(
                text + "\n\nВведите какую работу хотите отметить как необходимую проверить. Например\nЛР1 \nЛР2")
            await state.set_state(States.ReallyReady)
        except:
            kb = [
                [
                    types.KeyboardButton(text="Сообщить о загруженной работе")
                ],
            ]
            keyboard = types.ReplyKeyboardMarkup(
                keyboard=kb,
                resize_keyboard=True,
            )
            await message.answer("Данные введены некорректно, попробуйте снова",reply_markup=keyboard)

        @dp.message(States.ReallyReady)
        async def Ready(message: types.Message, state: FSMContext):
            await message.answer("Ожидайте")
            proverka = False
            for i in range(1, int(countoflab) + 1):
                text = "ЛР" + str(i)
                if text == (message.text).upper:
                    proverka = True
            if proverka == False:
                kb = [
                    [
                        types.KeyboardButton(text="Основы профессиональной деятельности"),
                        types.KeyboardButton(text="Объектно-ориентированное программирование"),
                        types.KeyboardButton(text="Динамические языки")
                    ],
                ]
                keyboard = types.ReplyKeyboardMarkup(
                    keyboard=kb,
                    resize_keyboard=True,
                )
                await message.answer("Работа с таким названием не найдена, повторите заного выбрав дисциплину заного:", reply_markup=keyboard)
                await state.set_state(States.ready)

            else:
                data = await state.get_data()
                Student = data['Familiya']
                Group = data['group']
                Status = await set_status_ready_for_inspection(data['ready'], Group, Student, (message.text).upper())
                await message.answer(Status)
                kb = [
                    [
                        types.KeyboardButton(text="Вернуться"),
                        types.KeyboardButton(text="Сообщить ещё")
                    ],
                ]
                keyboard = types.ReplyKeyboardMarkup(
                    keyboard=kb,
                    resize_keyboard=True,
                )
                await message.answer("Выберите действие", reply_markup=keyboard)

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())