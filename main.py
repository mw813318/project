import logging
import asyncio
import os
import sqlite3
from datetime import datetime, time, timedelta
from aiogram import Bot, Dispatcher, types
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters import Command
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.utils import executor
import openpyxl
from collections import defaultdict
from openpyxl.utils import get_column_letter

API_TOKEN = '7128425992:AAGbgXkXqUEzMTicL8Nv0Hgk8T2mst9G-sQ'
GROUP_CHAT_ID = -1002521462361

bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)

DB_PATH = "requests.db"

logging.basicConfig(level=logging.INFO)


# --- FSM ---
class Form(StatesGroup):
    supplier = State()
    amount = State()
    agent_name = State()
    agent_phone = State()
    delivery_date = State()
    admin_name = State()


# --- DB ---
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            username TEXT,
            supplier TEXT,
            amount REAL,
            agent_name TEXT,
            agent_phone TEXT,
            delivery_date TEXT,
            admin_name TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()


# --- Excel ---
def generate_excel_by_weekdays():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT supplier, amount, agent_name, agent_phone, delivery_date, admin_name, username FROM requests")
    rows = cur.fetchall()
    conn.close()

    if not rows:
        return None

    grouped = defaultdict(list)
    for r in rows:
        try:
            date_obj = datetime.strptime(r[4], "%Y-%m-%d").date()
            weekday = date_obj.strftime("%A")
            grouped[weekday].append(r)
        except:
            continue

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for day, entries in grouped.items():
        ws = wb.create_sheet(title=day)
        ws.append(['Поставщик','Сумма','Имя агента','Номер','Дата','Админ','От кого'])
        for e in entries:
            ws.append(e)
        for col in ws.columns:
            max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 3

    filename = f"weekly_requests_{datetime.now().strftime('%Y%m%d')}.xlsx"
    wb.save(filename)
    return filename


# --- Handlers ---
@dp.message_handler(commands=['start', 'заявка'], chat_type=types.ChatType.PRIVATE)
async def start_form(message: types.Message):
    await Form.supplier.set()
    await message.answer("Введи поставщика:")


@dp.message_handler(state=Form.supplier, chat_type=types.ChatType.PRIVATE)
async def step_supplier(message: types.Message, state: FSMContext):
    await state.update_data(supplier=message.text)
    await Form.next()
    await message.answer("Введи сумму:")


@dp.message_handler(state=Form.amount, chat_type=types.ChatType.PRIVATE)
async def step_amount(message: types.Message, state: FSMContext):
    try:
        amount = float(message.text.replace(',', '.').replace(' ', ''))
    except:
        return await message.answer("Некорректная сумма. Введи число.")
    await state.update_data(amount=amount)
    await Form.next()
    await message.answer("Имя агента:")


@dp.message_handler(state=Form.agent_name, chat_type=types.ChatType.PRIVATE)
async def step_agent_name(message: types.Message, state: FSMContext):
    await state.update_data(agent_name=message.text)
    await Form.next()
    await message.answer("Номер агента:")


@dp.message_handler(state=Form.agent_phone, chat_type=types.ChatType.PRIVATE)
async def step_agent_phone(message: types.Message, state: FSMContext):
    await state.update_data(agent_phone=message.text)
    await Form.next()
    await message.answer("Дата поставки (дд.мм.гггг):")


@dp.message_handler(state=Form.delivery_date, chat_type=types.ChatType.PRIVATE)
async def step_delivery_date(message: types.Message, state: FSMContext):
    try:
        date_obj = datetime.strptime(message.text, "%d.%m.%Y").date()
        await state.update_data(delivery_date=date_obj.strftime("%Y-%m-%d"))
        await Form.next()
        await message.answer("Имя админа:")
    except:
        await message.answer("Неверный формат даты. Введи ДД.ММ.ГГГГ")


@dp.message_handler(state=Form.admin_name, chat_type=types.ChatType.PRIVATE)
async def step_admin_name(message: types.Message, state: FSMContext):
    data = await state.get_data()
    await state.finish()

    user_id = message.from_user.id
    username = message.from_user.full_name

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO requests (user_id, username, supplier, amount, agent_name, agent_phone, delivery_date, admin_name)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        user_id, username, data['supplier'], data['amount'],
        data['agent_name'], data['agent_phone'],
        data['delivery_date'], message.text
    ))
    conn.commit()
    conn.close()

    await message.answer("✅ Заявка сохранена.")
    await bot.send_message(GROUP_CHAT_ID,
                           f"Новая заявка от {username}:\n"
                           f"Поставщик: {data['supplier']}\n"
                           f"Сумма: {data['amount']}\n"
                           f"Агент: {data['agent_name']}\n"
                           f"Номер: {data['agent_phone']}\n"
                           f"Дата: {data['delivery_date']}\n"
                           f"Админ: {message.text}")


@dp.message_handler(commands=['заявки'], chat_type=types.ChatType.PRIVATE)
async def list_requests(message: types.Message):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        SELECT supplier, amount, agent_name, agent_phone, delivery_date, admin_name
        FROM requests WHERE user_id = ?
    """, (message.from_user.id,))
    rows = cur.fetchall()
    conn.close()

    if not rows:
        return await message.answer("У тебя нет заявок.")

    text = "Твои заявки:\n"
    for i, r in enumerate(rows, 1):
        text += (f"\n{i}) Поставщик: {r[0]}\n"
                 f"Сумма: {r[1]}\n"
                 f"Агент: {r[2]}\n"
                 f"Номер: {r[3]}\n"
                 f"Дата: {r[4]}\n"
                 f"Админ: {r[5]}\n")
    await message.answer(text)


@dp.message_handler(commands=['экспорт'], chat_type=types.ChatType.PRIVATE)
async def export_requests(message: types.Message):
    filename = generate_excel_by_weekdays()
    if filename:
        await message.answer_document(types.InputFile(filename))
        os.remove(filename)
    else:
        await message.answer("Нет заявок для экспорта.")


# --- Планировщик ---
async def scheduler():
    while True:
        now = datetime.now()
        if now.time().hour == 20 and now.time().minute == 0:
            tomorrow = (now + timedelta(days=1)).date().strftime("%Y-%m-%d")
            conn = sqlite3.connect(DB_PATH)
            cur = conn.cursor()
            cur.execute("""
                SELECT supplier, amount, agent_name, agent_phone, delivery_date, admin_name, username
                FROM requests WHERE delivery_date = ?
            """, (tomorrow,))
            rows = cur.fetchall()
            conn.close()

            if rows:
                msg = "Заявки на завтра:\n"
                total = 0
                for i, r in enumerate(rows, 1):
                    total += float(r[1])
                    msg += (f"\n{i}) Поставщик: {r[0]}\n"
                            f"Сумма: {r[1]}\n"
                            f"Агент: {r[2]}\n"
                            f"Номер: {r[3]}\n"
                            f"Дата: {r[4]}\n"
                            f"Админ: {r[5]}\n")

                filename = generate_excel_by_week
