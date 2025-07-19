import logging
from aiogram import Bot, Dispatcher, types, executor
from aiogram.types import ReplyKeyboardRemove
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.dispatcher.filters import Command
from datetime import datetime, time
import asyncio
import openpyxl
import os

API_TOKEN = '7128425992:AAGbgXkXqUEzMTicL8Nv0Hgk8T2mst9G-sQ'
GROUP_CHAT_ID = -1002521462361 # Заменить на реальный ID группы

logging.basicConfig(level=logging.INFO)

bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot, storage=MemoryStorage())

class Form(StatesGroup):
    supplier = State()
    amount = State()
    agent_name = State()
    agent_phone = State()
    delivery_date = State()
    admin_name = State()

# Временное хранилище заявок
user_requests = {}

@dp.message_handler(lambda message: message.chat.type == 'private', commands=['start', 'заявка'])
@dp.message_handler(lambda message: message.chat.type == 'private' and message.text.lower() == 'заявка')
async def start_form(message: types.Message):
    await Form.supplier.set()
    await message.reply("✏️ Введи поставщика:", reply_markup=ReplyKeyboardRemove())

@dp.message_handler(state=Form.supplier)
async def step_supplier(message: types.Message, state: FSMContext):
    await state.update_data(supplier=message.text)
    await Form.next()
    await message.reply("💰 Введи сумму:")

@dp.message_handler(state=Form.amount)
async def step_amount(message: types.Message, state: FSMContext):
    await state.update_data(amount=message.text)
    await Form.next()
    await message.reply("👤 Имя агента:")

@dp.message_handler(state=Form.agent_name)
async def step_agent_name(message: types.Message, state: FSMContext):
    await state.update_data(agent_name=message.text)
    await Form.next()
    await message.reply("📞 Номер агента:")

@dp.message_handler(state=Form.agent_phone)
async def step_agent_phone(message: types.Message, state: FSMContext):
    await state.update_data(agent_phone=message.text)
    await Form.next()
    await message.reply("📅 Дата поставки (дд.мм.гггг):")

@dp.message_handler(state=Form.delivery_date)
async def step_delivery_date(message: types.Message, state: FSMContext):
    await state.update_data(delivery_date=message.text)
    await Form.next()
    await message.reply("🧑‍💼 Имя админа (кто подаёт заявку):")

@dp.message_handler(state=Form.admin_name)
async def step_admin_name(message: types.Message, state: FSMContext):
    await state.update_data(admin_name=message.text)
    data = await state.get_data()

    user_id = message.from_user.id
    username = message.from_user.full_name

    if user_id not in user_requests:
        user_requests[user_id] = []

    request = {
        'Поставщик': data['supplier'],
        'Сумма': data['amount'],
        'Имя агента': data['agent_name'],
        'Номер агента': data['agent_phone'],
        'Дата поставки': data['delivery_date'],
        'Имя админа': data['admin_name'],
        'От кого': username
    }

    user_requests[user_id].append(request)

    await message.reply("✅ Заявка сохранена!")

    # Отправка в группу
    await bot.send_message(GROUP_CHAT_ID, f"""
📦 Новая заявка от {username}:

Поставщик: {data['supplier']}
Сумма: {data['amount']}
Имя агента: {data['agent_name']}
Номер: {data['agent_phone']}
Дата поставки: {data['delivery_date']}
Имя админа: {data['admin_name']}
""")

    await state.finish()

@dp.message_handler(commands=['заявки'])
async def list_requests(message: types.Message):
    user_id = message.from_user.id
    if user_id not in user_requests or not user_requests[user_id]:
        await message.reply("У тебя нет заявок на сегодня.")
        return

    text = "📋 Твои заявки:\n"
    for i, req in enumerate(user_requests[user_id], 1):
        text += f"\n{i}) Поставщик: {req['Поставщик']}\nСумма: {req['Сумма']}\nИмя агента: {req['Имя агента']}\nНомер: {req['Номер агента']}\nДата: {req['Дата поставки']}\nИмя админа: {req['Имя админа']}\n"
    await message.reply(text)

def generate_excel(requests: dict):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Заявки"

    headers = ['Поставщик', 'Сумма', 'Имя агента', 'Номер агента', 'Дата поставки', 'Имя админа', 'От кого']
    ws.append(headers)
    total = 0

    for user_list in requests.values():
        for req in user_list:
            row = [req[h] for h in headers]
            ws.append(row)
            try:
                total += float(str(req['Сумма']).replace(' ', '').replace(',', '.'))
            except:
                pass

    file_path = f"requests_{datetime.now().date()}.xlsx"
    wb.save(file_path)
    return file_path, total

async def send_daily_summary():
    if not user_requests:
        return

    text = "📦 Заявки на завтра:\n"
    all_requests = []
    for uid, requests in user_requests.items():
        all_requests.extend(requests)

    file_path, total = generate_excel(user_requests)

    for i, req in enumerate(all_requests, 1):
        text += f"\n{i}) Поставщик: {req['Поставщик']}\nСумма: {req['Сумма']}\nИмя агента: {req['Имя агента']}\nНомер: {req['Номер агента']}\nДата: {req['Дата поставки']}\nИмя админа: {req['Имя админа']}\n"

    text += f"\n💰 Общая сумма: {total} сом"

    await bot.send_message(GROUP_CHAT_ID, text)
    await bot.send_document(GROUP_CHAT_ID, open(file_path, 'rb'))
    os.remove(file_path)

    user_requests.clear()

async def scheduler():
    while True:
        now = datetime.now().time()
        if now >= time(20, 0) and now < time(20, 1):
            await send_daily_summary()
            await asyncio.sleep(60)
        await asyncio.sleep(30)

if __name__ == '__main__':
    loop = asyncio.get_event_loop()
    loop.create_task(scheduler())
    executor.start_polling(dp, skip_updates=True)
