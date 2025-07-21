import logging
import asyncio
import os
import sqlite3
from datetime import datetime, timedelta
from collections import defaultdict

from aiogram import Bot, Dispatcher, Router, F
from aiogram.types import Message, FSInputFile
from aiogram.filters import CommandStart, Command
from aiogram.fsm.state import StatesGroup, State
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.memory import MemoryStorage
import openpyxl
from openpyxl.utils import get_column_letter

API_TOKEN = '7128425992:AAGbgXkXqUEzMTicL8Nv0Hgk8T2mst9G-sQ'
GROUP_CHAT_ID = -1002521462361

logging.basicConfig(level=logging.INFO)

# Create bot and dispatcher
bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)
router = Router()

DB_PATH = 'requests.db'

# FSM States
class Form(StatesGroup):
    supplier = State()
    amount = State()
    agent_name = State()
    agent_phone = State()
    delivery_date = State()
    admin_name = State()

# Database initialization
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

# Excel generation
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

# Command handlers
@router.message(CommandStart())
@router.message(Command("заявка"))
async def start_form(message: Message, state: FSMContext):
    await state.set_state(Form.supplier)
    await message.answer("Введи поставщика:")

@router.message(Form.supplier)
async def step_supplier(message: Message, state: FSMContext):
    await state.update_data(supplier=message.text)
    await state.set_state(Form.amount)
    await message.answer("Введи сумму:")

@router.message(Form.amount)
async def step_amount(message: Message, state: FSMContext):
    try:
        amount = float(message.text.replace(',', '.').replace(' ', ''))
    except:
        return await message.answer("Некорректная сумма. Введи число.")
    await state.update_data(amount=amount)
    await state.set_state(Form.agent_name)
    await message.answer("Имя агента:")

@router.message(Form.agent_name)
async def step_agent_name(message: Message, state: FSMContext):
    await state.update_data(agent_name=message.text)
    await state.set_state(Form.agent_phone)
    await message.answer("Номер агента:")

@router.message(Form.agent_phone)
async def step_agent_phone(message: Message, state: FSMContext):
    await state.update_data(agent_phone=message.text)
    await state.set_state(Form.delivery_date)
    await message.answer("Дата поставки (дд.мм.гггг):")
@router.message(Form.delivery_date)
async def step_delivery_date(message: Message, state: FSMContext):
    try:
        date_obj = datetime.strptime(message.text, "%d.%m.%Y").date()
        await state.update_data(delivery_date=date_obj.strftime("%Y-%m-%d"))
        await state.set_state(Form.admin_name)
        await message.answer("Имя админа:")
    except:
        await message.answer("Неверный формат даты. Введи ДД.ММ.ГГГГ")

@router.message(Form.admin_name)
async def step_admin_name(message: Message, state: FSMContext):
    data = await state.update_data(admin_name=message.text)
    user = message.from_user

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO requests (user_id, username, supplier, amount, agent_name, agent_phone, delivery_date, admin_name)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        user.id, user.full_name, data['supplier'], data['amount'],
        data['agent_name'], data['agent_phone'], data['delivery_date'], data['admin_name']
    ))
    conn.commit()
    conn.close()

    await state.clear()
    await message.answer("✅ Заявка сохранена.")
    await bot.send_message(GROUP_CHAT_ID,
        f"📦 Новая заявка от {user.full_name}:\n\n"
        f"Поставщик: {data['supplier']}\n"
        f"Сумма: {data['amount']}\n"
        f"Агент: {data['agent_name']}\n"
        f"Номер: {data['agent_phone']}\n"
        f"Дата: {data['delivery_date']}\n"
        f"Админ: {data['admin_name']}\n")

@router.message(Command("заявки"))
async def list_requests(message: Message):
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

    text = "📦 Твои заявки:\n"
    for i, r in enumerate(rows, 1):
        text += (f"\n{i}) Поставщик: {r[0]}\n"
                 f"Сумма: {r[1]}\n"
                 f"Агент: {r[2]}\n"
                 f"Номер: {r[3]}\n"
                 f"Дата: {r[4]}\n"
                 f"Админ: {r[5]}\n")
    await message.answer(text)

@router.message(Command("экспорт"))
async def export_requests(message: Message):
    filename = generate_excel_by_weekdays()
    if filename:
        await message.answer_document(FSInputFile(filename))
        os.remove(filename)
    else:
        await message.answer("Нет заявок для экспорта.")

# Scheduled task (use apscheduler for production)
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
                msg = "📦 Заявки на завтра:\n"
                total = 0
                for i, r in enumerate(rows, 1):
                    total += float(r[1])
                    msg += (f"\n{i}) Поставщик: {r[0]}\n"
                            f"Сумма: {r[1]}\n"
                            f"Агент: {r[2]}\n"
                            f"Номер: {r[3]}\n"
                            f"Дата: {r[4]}\n"
                            f"Админ: {r[5]}\n")
                await bot.send_message(GROUP_CHAT_ID, msg)
        await asyncio.sleep(60)

async def main():
    init_db()
    dp.include_router(router)
    await dp.start_polling(bot, polling_timeout=30)

if __name__ == '__main__':
    asyncio.run(main())