import os
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta, date
from flask import Flask, render_template, request, redirect, url_for, session, send_file
from flask_sqlalchemy import SQLAlchemy
from openpyxl import Workbook
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
import pytz

app = Flask(__name__)
app.secret_key = 'replace-this-with-secret-key-in-production'  # Обязательно поменяйте!

# ---------- НАСТРОЙКИ БАЗЫ ----------
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///registrations.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# ---------- МОДЕЛЬ БАЗЫ ----------
class Registration(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    last_name = db.Column(db.String(50), nullable=False)
    first_name = db.Column(db.String(50), nullable=False)
    middle_name = db.Column(db.String(50), nullable=True)
    birth_date = db.Column(db.String(10), nullable=False)  # DD.MM.YYYY
    phone = db.Column(db.String(20), nullable=False)
    passport_series = db.Column(db.String(10), nullable=False)
    passport_number = db.Column(db.String(20), nullable=False)
    passport_issued_by = db.Column(db.String(200), nullable=False)
    passport_issued_date = db.Column(db.String(10), nullable=False)
    address = db.Column(db.String(300), nullable=False)
    reg_date = db.Column(db.String(10), nullable=False)    # DD.MM.YYYY
    reg_time = db.Column(db.String(5), nullable=False)     # HH:MM
    created_at = db.Column(db.DateTime, default=datetime.now)

    def __repr__(self):
        return f"{self.last_name} {self.first_name} - {self.reg_date} {self.reg_time}"

# ---------- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ----------
def parse_date(d_str):
    """Преобразует DD.MM.YYYY в datetime.date"""
    return datetime.strptime(d_str, '%d.%m.%Y').date()

def get_available_dates():
    """Возвращает список дат (строки DD.MM.YYYY) на ближайшие 30 дней,
       где: будни, не сегодня и не вчера, свободно < 16 мест"""
    today = date.today()
    result = []
    for i in range(30):
        d = today + timedelta(days=i)
        if d.weekday() in [5, 6]:  # Сб, Вс — пропускаем
            continue
        # Пропускаем сегодня и вчера
        if d <= today:
            continue
        d_str = d.strftime('%d.%m.%Y')
        count = Registration.query.filter_by(reg_date=d_str).count()
        if count < 16:
            result.append(d_str)
    return result

def get_booked_times(date_str):
    """Возвращает список уже занятых слотов на указанную дату (строки HH:MM)"""
    recs = Registration.query.filter_by(reg_date=date_str).all()
    return [r.reg_time for r in recs]

def generate_time_slots():
    """Генерирует слоты с 08:00 до 11:45 с шагом 15 минут"""
    slots = []
    start = datetime.strptime('08:00', '%H:%M')
    end = datetime.strptime('11:45', '%H:%M')
    current = start
    while current <= end:
        slots.append(current.strftime('%H:%M'))
        current += timedelta(minutes=15)
    return slots

def is_valid_phone(phone):
    """Простая валидация телефона (цифры, +, пробелы, скобки)"""
    cleaned = re.sub(r'[\s\(\)\-]', '', phone)
    return re.match(r'^\+?\d{10,15}$', cleaned) is not None

# ---------- ГАРАНТИРУЕМ СОЗДАНИЕ ТАБЛИЦ ПРИ СТАРТЕ ----------
with app.app_context():
    db.create_all()
# ---------- ПУБЛИЧНАЯ ФОРМА (ДЛЯ ИНОСТРАНЦЕВ) ----------
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Получаем все данные из формы
        data = {
            'last_name': request.form.get('last_name', ''),
            'first_name': request.form.get('first_name', ''),
            'middle_name': request.form.get('middle_name', ''),
            'birth_date': request.form.get('birth_date', ''),
            'phone': request.form.get('phone', ''),
            'passport_series': request.form.get('passport_series', ''),
            'passport_number': request.form.get('passport_number', ''),
            'passport_issued_by': request.form.get('passport_issued_by', ''),
            'passport_issued_date': request.form.get('passport_issued_date', ''),
            'address': request.form.get('address', ''),
            'reg_date': request.form.get('reg_date', ''),
            'reg_time': request.form.get('reg_time', ''),
        }
        
        # Проверяем, была ли нажата кнопка "Записаться" (финальная отправка)
        if 'submit_registration' in request.form:
            # === ВАЛИДАЦИЯ ===
            errors = []
            if not data['last_name']: errors.append('Фамилия обязательна')
            if not data['first_name']: errors.append('Имя обязательно')
            # ... добавьте остальные проверки из вашего кода ...
            
            if errors:
                # Если ошибки — показываем форму с данными и ошибками
                available_dates = get_available_dates()
                all_slots = generate_time_slots()
                booked_slots = get_booked_times(data['reg_date']) if data['reg_date'] else []
                return render_template('form.html',
                                       errors=errors,
                                       data=data,
                                       available_dates=available_dates,
                                       all_slots=all_slots,
                                       booked_slots=booked_slots,
                                       clinic_address="г. Ульяновск, Московское шоссе, 92",
                                       clinic_phone="8 (8422) 22-97-80",
                                       work_hours="08:00 – 12:00, будни (кроме праздников)")
            
            # Сохраняем в базу
            new_reg = Registration(
                last_name=data['last_name'],
                first_name=data['first_name'],
                middle_name=data['middle_name'],
                birth_date=data['birth_date'],
                phone=data['phone'],
                passport_series=data['passport_series'],
                passport_number=data['passport_number'],
                passport_issued_by=data['passport_issued_by'],
                passport_issued_date=data['passport_issued_date'],
                address=data['address'],
                reg_date=data['reg_date'],
                reg_time=data['reg_time']
            )
            # Перед сохранением
            print(f"Попытка сохранить: {data['last_name']} {data['first_name']} на {data['reg_date']} {data['reg_time']}")
            
            try:
                db.session.add(new_reg)
                db.session.commit()
                print("✅ Запись сохранена успешно!")
            except Exception as e:
                print(f"❌ Ошибка сохранения: {e}")
                db.session.rollback()
            
            return render_template('success.html',
                                   reg_date=data['reg_date'],
                                   reg_time=data['reg_time'],
                                   clinic_address="г. Ульяновск, Московское шоссе, 92",
                                   clinic_phone="8 (8422) 22-97-80")
        
        # Если это не финальная отправка, значит, это выбор даты
        # Показываем форму с уже заполненными данными
        available_dates = get_available_dates()
        all_slots = generate_time_slots()
        booked_slots = get_booked_times(data['reg_date']) if data['reg_date'] else []
        
        return render_template('form.html',
                               errors=[],
                               data=data,  # ← передаём заполненные данные!
                               available_dates=available_dates,
                               all_slots=all_slots,
                               booked_slots=booked_slots,
                               clinic_address="г. Ульяновск, Московское шоссе, 92",
                               clinic_phone="8 (8422) 22-97-80",
                               work_hours="08:00 – 12:00, будни (кроме праздников)")
    
    # GET-запрос — показываем пустую форму
    available_dates = get_available_dates()
    all_slots = generate_time_slots()
    return render_template('form.html',
                           errors=[],
                           data={},
                           available_dates=available_dates,
                           all_slots=all_slots,
                           booked_slots=[],
                           clinic_address="г. Ульяновск, Московское шоссе, 92",
                           clinic_phone="8 (8422) 22-97-80",
                           work_hours="08:00 – 12:00, будни (кроме праздников)")
# ---------- АДМИНКА (ДЛЯ ВЛАДЕЛЬЦА) ----------
@app.route('/admin', methods=['GET', 'POST'])
def admin():
    # Простейшая защита — пароль в строке запроса (для демо)
    # В реальном проекте поставьте нормальную авторизацию!
    if request.args.get('key') != 'ADMIN2026':
        return "Доступ запрещён. Укажите ?key=ADMIN2026", 403

    if request.method == 'POST':
        # Выгрузка Excel с фильтром по дате
        filter_date = request.form.get('filter_date', '').strip()
        query = Registration.query.order_by(Registration.reg_date, Registration.reg_time)
        if filter_date:
            try:
                datetime.strptime(filter_date, '%d.%m.%Y')
                query = query.filter_by(reg_date=filter_date)
            except ValueError:
                pass

        registrations = query.all()

        wb = Workbook()
        ws = wb.active
        ws.title = "Записи на медосмотр"

        # Заголовки
        headers = [
            'ID', 'Фамилия', 'Имя', 'Отчество', 'Дата рождения',
            'Телефон', 'Серия паспорта', 'Номер паспорта',
            'Паспорт выдан (кем)', 'Паспорт выдан (дата)',
            'Адрес пребывания', 'Дата записи', 'Время записи',
            'Дата создания заявки'
        ]
        ws.append(headers)

        for r in registrations:
            ws.append([
                r.id,
                r.last_name,
                r.first_name,
                r.middle_name,
                r.birth_date,
                r.phone,
                r.passport_series,
                r.passport_number,
                r.passport_issued_by,
                r.passport_issued_date,
                r.address,
                r.reg_date,
                r.reg_time,
                r.created_at.strftime('%d.%m.%Y %H:%M')
            ])

        # Сохраняем во временный файл
        filepath = f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        wb.save(filepath)
        return send_file(filepath, as_attachment=True, download_name='zapisi_medosmotr.xlsx')

    # GET — показываем таблицу записей
    registrations = Registration.query.order_by(
        Registration.reg_date.desc(),
        Registration.reg_time.desc()
    ).all()
    return render_template('admin.html', registrations=registrations)
# ---------- МАРШРУТ ДЛЯ УДАЛЕНИЯ ЗАПИСЕЙ в АДМИНКЕ (ДЛЯ ВЛАДЕЛЬЦА) ----------
@app.route('/delete/<int:record_id>', methods=['POST'])
def delete_record(record_id):
    # Проверяем пароль (защита)
    if request.args.get('key') != 'ADMIN2026':
        return "Доступ запрещён", 403
    
    record = Registration.query.get(record_id)
    if record:
        db.session.delete(record)
        db.session.commit()
        return redirect(url_for('admin', key='ADMIN2026'))
    else:
        return "Запись не найдена", 404

# ---------- АВТОМАТИЧЕСКАЯ ОТПРАВКА В 11:00 (мск)----------
def send_daily_report():
    with app.app_context():
        today_str = date.today().strftime('%d.%m.%Y')
        
        # Получаем все записи от сегодняшней даты и дальше
        registrations = Registration.query.filter(
            Registration.reg_date >= today_str
        ).order_by(Registration.reg_date, Registration.reg_time).all()
        
        if not registrations:
            print(f"[{datetime.now()}] Нет записей от {today_str} и дальше, письмо не отправлено")
            return

        # Создаём Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "Записи на медосмотр"
        
        headers = ['Фамилия', 'Имя', 'Отчество', 'Дата рождения', 'Телефон',
                   'Серия паспорта', 'Номер паспорта', 'Кем выдан', 'Дата выдачи',
                   'Адрес', 'Дата записи', 'Время записи']
        ws.append(headers)
        
        for r in registrations:
            ws.append([
                r.last_name, r.first_name, r.middle_name,
                r.birth_date, r.phone, r.passport_series, r.passport_number,
                r.passport_issued_by, r.passport_issued_date,
                r.address, r.reg_date, r.reg_time
            ])
        
        filepath = f"daily_{today_str}.xlsx"
        wb.save(filepath)

        # === НАСТРОЙКИ SMTP (ЗАМЕНИТЕ НА СВОИ) ===
        # Рекомендую использовать Яндекс.Почту или Mail.ru для отправки
        SMTP_SERVER = 'smtp.yandex.ru'   # или smtp.mail.ru
        SMTP_PORT = 465
        SMTP_USER = 'dyadyk.yurij@yandex.ru'
        SMTP_PASSWORD = 'skskixmzczcmleih'  # Специальный пароль для приложений
        TO_EMAIL = 'dyadyk.yurij@yandex.ru'

        msg = MIMEMultipart()
        msg['From'] = SMTP_USER
        msg['To'] = TO_EMAIL
        msg['Subject'] = f'Список иностранцев, записавшихся на медосмотр с {today_str} и далее (ежедневная выгрузка)'

        body = f'Во вложении — список из {len(registrations)} иностранцев, записавшихся на медосмотр на сегодня ({today_str}) и далее.\n\n' \
               f'Адрес клиники: г. Ульяновск, Московское шоссе, 92\n' \
               f'Телефон регистратуры: 8 (8422) 22-97-80'
        msg.attach(MIMEText(body, 'plain'))

        with open(filepath, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename=zapisi_{today_str}.xlsx')
            msg.attach(part)

        try:
            with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
                server.login(SMTP_USER, SMTP_PASSWORD)
                server.send_message(msg)
            print(f"[{datetime.now()}] Отчёт по записавшимся на медосмотр на {today_str} и далее отправлен успешно")
        except Exception as e:
            print(f"[{datetime.now()}] Ошибка отправки: {e}")

        # Удаляем временный файл
        os.remove(filepath)

# ---------- ЗАПУСК ПЛАНИРОВЩИКА (улучшенная версия) ----------
# Импортируем модуль os для проверки окружения
import os

# Создаём планировщик
scheduler = BackgroundScheduler()
scheduler.add_job(
    func=send_daily_report,
    trigger=CronTrigger(hour=11, minute=0, timezone=pytz.timezone('Europe/Moscow')),
    id='daily_report',
    misfire_grace_time=3600  # Даём задаче 1 час на выполнение, если она пропущена
)

# Запускаем планировщик ТОЛЬКО если:
# 1. Приложение НЕ в режиме отладки (debug=False)
# 2. Это НЕ дочерний процесс Werkzeug (чтобы не запускать дважды при разработке)
# 3. Планировщик ещё не запущен
if not app.debug and not os.environ.get('WERKZEUG_RUN_MAIN') and not scheduler.running:
    scheduler.start()
    print("✅ APScheduler успешно запущен!")

# ---------- МАРШРУТ ДЛЯ ПРИНУДИТЕЛЬНОЙ ОТПРАВКИ (ЗАПАСНОЙ) ----------
@app.route('/trigger-report')
def trigger_report():
    """Принудительный запуск отправки отчёта (для тестирования или по cron-запросу)"""
    try:
        send_daily_report()
        return "✅ Отчёт отправлен принудительно", 200
    except Exception as e:
        return f"❌ Ошибка: {e}", 500

# ---------- ТОЧКА ВХОДА ----------
if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(host='0.0.0.0', port=5000, debug=False)
