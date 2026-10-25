# -*- coding: utf-8 -*-
"""
سیستم جامع مدیریت پایان‌نامه دانشگاهی
نسخه ۱.۰

توضیحات کامل سیستم:
این برنامه یک نرم‌افزار دسکتاپ کامل برای مدیریت فرآیند پایان‌نامه در دانشگاه است.
سه نقش اصلی وجود دارد: ادمین (مسئول آموزش)، استاد راهنما، و دانشجو.
امکانات اصلی شامل ثبت‌نام دانشجو، تأیید توسط ادمین، ارسال رمز عبور موقت،
ساخت استاد توسط ادمین، درخواست راهنمایی توسط دانشجو و پذیرش توسط استاد،
مدیریت پایان‌نامه، جلسات، فایل‌ها، پیام‌ها، ارزیابی دفاع، بایگانی و گزارش‌ها می‌باشد.

نحوه اجرا:
۱. نصب پیش‌نیازها با دستور زیر:
   pip install PyQt6 pyqtgraph pandas openpyxl sqlalchemy qdarkstyle loguru schedule plyer pillow PyPDF2 markdown jdatetime

   توجه: بسته plyer برای نوتیفیکیشن‌های دسکتاپ (اختیاری) و jdatetime برای تاریخ شمسی استفاده می‌شود.
   برای استفاده از وب‌کم، بسته‌های PyQt6-Qt6-Multimedia نیز ممکن است لازم باشد.

۲. فایل main.py را با پایتون ۳.۱۱ اجرا کنید:
   python main.py

۳. در اولین اجرا، پایگاه داده theses.db ساخته می‌شود و یک ادمین پیش‌فرض ایجاد می‌گردد:
   نام کاربری: admin
   رمز عبور: admin

۴. برای ارسال ایمیل، متغیرهای محیطی زیر باید تنظیم شوند (قابل تنظیم در کد):
   SMTP_SERVER : آدرس سرور SMTP (مثلاً smtp.gmail.com)
   SMTP_PORT : پورت (مثلاً 587)
   SENDER_EMAIL : ایمیل فرستنده
   SENDER_PASSWORD : رمز عبور یا رمز برنامه

   در غیر این صورت، قسمت ارسال ایمیل غیرفعال می‌شود و رمزها در کنسول چاپ می‌شوند.

تمامی متن‌های رابط کاربری، پیغام‌ها و کامنت‌ها به زبان فارسی روان می‌باشند.
فونت استفاده شده Vazirmatn است (باید روی سیستم نصب باشد).
ظاهر برنامه مدرن با تم تاریک/روشن و دکمه‌های گرد و گرادینت می‌باشد.

نویسنده: تیم حرفه‌ای توسعه‌دهندگان پایتون
"""
from PyQt6.QtGui import QDesktopServices
import sys
import os
import hashlib
import random
import string
import shutil
import json
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, List

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from PIL import Image, ImageQt
import markdown


# PyQt6
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QTableWidget, QTableWidgetItem,
    QStackedWidget, QMessageBox, QFileDialog, QCheckBox, QTextEdit,
    QTabWidget, QGroupBox, QGridLayout, QSplitter, QScrollArea, QProgressBar,
    QStatusBar, QToolBar, QHeaderView, QMenu, QRadioButton, QButtonGroup,
    QDialog, QDialogButtonBox, QFormLayout, QSpinBox, QDoubleSpinBox, QDateEdit,
    QCalendarWidget, QListWidget, QListWidgetItem, QFrame, QSpacerItem, QSizePolicy,
    QInputDialog
)
from PyQt6.QtGui import (
    QPixmap, QImage, QFont, QIcon, QAction, QColor, QPalette, QLinearGradient,
    QBrush, QPainter, QPen, QCursor, QKeySequence
)
from PyQt6.QtCore import (
    Qt, QThread, pyqtSignal, QDate, QTimer, QTime, QUrl, QDir
)

# PyQtGraph
import pyqtgraph as pg
pg.setConfigOption('background', 'w')  # پیش‌فرض سفید
pg.setConfigOption('foreground', 'k')

# SQLAlchemy
from sqlalchemy import (
    create_engine, Column, Integer, String, Float, Boolean, DateTime, ForeignKey, Text, JSON, Date
)
from sqlalchemy.orm import declarative_base, sessionmaker, relationship, joinedload

# Loguru
from loguru import logger

# Email
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Notifications (اختیاری)
try:
    from plyer import notification
    PLYER_AVAILABLE = True
except ImportError:
    PLYER_AVAILABLE = False

# jdatetime for Shamsi date
try:
    import jdatetime
    JALALI_AVAILABLE = True
except ImportError:
    JALALI_AVAILABLE = False

# ----------------------------------------------------------------------
# تنظیمات لاگ
# ----------------------------------------------------------------------
LOG_DIR = Path('logs')
LOG_DIR.mkdir(exist_ok=True)
logger.add(LOG_DIR / "app_{time}.log", rotation="10 MB", retention="30 days", level="INFO")
logger.info("شروع برنامه")

# ----------------------------------------------------------------------
# متغیرهای محیطی ایمیل (پیش‌فرض‌ها)
# ----------------------------------------------------------------------
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
SENDER_EMAIL = 'bodys9895@gmail.com'
SENDER_PASSWORD = 'brfwlljnzoovztdk'   # رمز برنامه ۱۶ رقمی معتبر

def can_send_email():
    return True   
# ----------------------------------------------------------------------
# مسیرها
# ----------------------------------------------------------------------
DB_PATH = 'theses.db'
BACKUP_DIR = Path('backups')
BACKUP_DIR.mkdir(exist_ok=True)
ARCHIVE_DIR = Path('archive')
ARCHIVE_DIR.mkdir(exist_ok=True)
UPLOAD_DIR = Path('uploads')
UPLOAD_DIR.mkdir(exist_ok=True)

# ----------------------------------------------------------------------
# پایگاه داده SQLAlchemy
# ----------------------------------------------------------------------
Base = declarative_base()

class User(Base):
    __tablename__ = 'users'
    id = Column(Integer, primary_key=True)
    username = Column(String, unique=True, nullable=False)
    password_hash = Column(String, nullable=False)
    role = Column(String, nullable=False)  # admin, professor, student
    related_id = Column(Integer, nullable=True)  # کلید خارجی به جدول دانشجو یا استاد
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.now)

class Student(Base):
    __tablename__ = 'students'
    id = Column(Integer, primary_key=True)
    national_id = Column(String, unique=True)
    student_id = Column(String, unique=True)
    first_name = Column(String)
    last_name = Column(String)
    field_of_study = Column(String)
    grade = Column(String)  # مقطع
    entry_year = Column(Integer)
    email = Column(String)
    phone = Column(String)
    photo_path = Column(String)
    is_approved_by_admin = Column(Boolean, default=False)
    # رابطه
    theses = relationship("Thesis", back_populates="student", uselist=False)
    user = relationship("User", primaryjoin="User.related_id==Student.id", foreign_keys=id, uselist=False, viewonly=True)

class Professor(Base):
    __tablename__ = 'professors'
    id = Column(Integer, primary_key=True)
    professor_code = Column(String, unique=True)
    first_name = Column(String)
    last_name = Column(String)
    expertise = Column(String)
    email = Column(String)
    phone = Column(String)
    max_students = Column(Integer, default=5)
    # رابطه‌ها
    theses_supervised = relationship("Thesis", back_populates="supervisor", foreign_keys="Thesis.supervisor_id")
    theses_advised = relationship("Thesis", back_populates="advisor", foreign_keys="Thesis.advisor_id")

class Thesis(Base):
    __tablename__ = 'theses'
    id = Column(Integer, primary_key=True)
    student_id = Column(Integer, ForeignKey('students.id'), unique=True)
    title = Column(String)
    supervisor_id = Column(Integer, ForeignKey('professors.id'), nullable=True)
    advisor_id = Column(Integer, ForeignKey('professors.id'), nullable=True)
    proposal_approval_date = Column(Date, nullable=True)
    start_date = Column(Date, nullable=True)
    defense_date = Column(Date, nullable=True)
    project_status = Column(String, default="در حال انجام")  # "در حال انجام", "متوقف", "دفاع شده"
    progress_percent = Column(Float, default=0.0)
    final_grade = Column(Float, nullable=True)
    is_archived = Column(Boolean, default=False)
    # روابط
    student = relationship("Student", back_populates="theses")
    supervisor = relationship("Professor", foreign_keys=[supervisor_id], back_populates="theses_supervised")
    advisor = relationship("Professor", foreign_keys=[advisor_id], back_populates="theses_advised")
    chapters = relationship("Chapter", back_populates="thesis")
    files = relationship("UploadedFile", back_populates="thesis")
    meetings = relationship("Meeting", back_populates="thesis")
    defense_evaluation = relationship("DefenseEvaluation", back_populates="thesis", uselist=False)

class Request(Base):
    __tablename__ = 'requests'
    id = Column(Integer, primary_key=True)
    student_id = Column(Integer, ForeignKey('students.id'))
    professor_id = Column(Integer, ForeignKey('professors.id'))
    request_date = Column(Date, default=datetime.now().date())
    status = Column(String, default="pending")  # pending, accepted, rejected
    student_message = Column(Text)

class Meeting(Base):
    __tablename__ = 'meetings'
    id = Column(Integer, primary_key=True)
    student_id = Column(Integer, ForeignKey('students.id'))
    professor_id = Column(Integer, ForeignKey('professors.id'))
    thesis_id = Column(Integer, ForeignKey('theses.id'))
    meeting_date = Column(Date)
    topic = Column(String)
    professor_notes = Column(Text)
    assigned_tasks = Column(Text)
    next_meeting_date = Column(Date, nullable=True)
    # رابطه
    thesis = relationship("Thesis", back_populates="meetings")

class Chapter(Base):
    __tablename__ = 'chapters'
    id = Column(Integer, primary_key=True)
    thesis_id = Column(Integer, ForeignKey('theses.id'))
    chapter_number = Column(Integer)
    title = Column(String)
    file_path = Column(String)
    upload_date = Column(Date, default=datetime.now().date())
    is_approved = Column(Boolean, default=False)
    approval_date = Column(Date, nullable=True)
    revision_reason = Column(Text)
    # رابطه
    thesis = relationship("Thesis", back_populates="chapters")

class UploadedFile(Base):
    __tablename__ = 'files'
    id = Column(Integer, primary_key=True)
    thesis_id = Column(Integer, ForeignKey('theses.id'))
    file_type = Column(String)  # proposal, chapter, article, final_thesis, defense_presentation
    original_name = Column(String)
    stored_path = Column(String)
    upload_date = Column(Date, default=datetime.now().date())
    version = Column(Integer, default=1)
    thesis = relationship("Thesis", back_populates="files")

class Message(Base):
    __tablename__ = 'messages'
    id = Column(Integer, primary_key=True)
    from_user_id = Column(Integer, ForeignKey('users.id'))
    to_user_id = Column(Integer, ForeignKey('users.id'))
    message_text = Column(Text)
    timestamp = Column(DateTime, default=datetime.now)
    is_read = Column(Boolean, default=False)

class Alert(Base):
    __tablename__ = 'alerts'
    id = Column(Integer, primary_key=True)
    user_id = Column(Integer, ForeignKey('users.id'))
    alert_text = Column(String)
    alert_date = Column(DateTime, default=datetime.now)
    is_done = Column(Boolean, default=False)

class DefenseEvaluation(Base):
    __tablename__ = 'defense_evaluation'
    id = Column(Integer, primary_key=True)
    thesis_id = Column(Integer, ForeignKey('theses.id'), unique=True)
    professor_id = Column(Integer, ForeignKey('professors.id'))
    evaluation_data_json = Column(JSON)  # ذخیره امتیازات به صورت دیکشنری
    final_score = Column(Float, nullable=True)
    is_final = Column(Boolean, default=False)
    thesis = relationship("Thesis", back_populates="defense_evaluation")

engine = create_engine(f'sqlite:///{DB_PATH}')
Base.metadata.create_all(engine)
Session = sessionmaker(bind=engine)

def get_session():
    return Session()

# ----------------------------------------------------------------------
# توابع کمکی
# ----------------------------------------------------------------------
def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode('utf-8')).hexdigest()

def generate_temp_password(length=8) -> str:
    chars = string.ascii_letters + string.digits
    return ''.join(random.choice(chars) for _ in range(length))

class EmailSender(QThread):
    finished = pyqtSignal(bool, str)  # success, message

    def __init__(self, to_email, subject, body):
        super().__init__()
        self.to_email = to_email
        self.subject = subject
        self.body = body

    def run(self):
        if not can_send_email():
            self.finished.emit(False, "تنظیمات ایمیل یافت نشد. لطفاً متغیرهای محیطی را تنظیم کنید.")
            return
        try:
            msg = MIMEMultipart()
            msg['From'] = SENDER_EMAIL
            msg['To'] = self.to_email
            msg['Subject'] = self.subject
            msg.attach(MIMEText(self.body, 'plain', 'utf-8'))

            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(SENDER_EMAIL, SENDER_PASSWORD)
                server.send_message(msg)
            self.finished.emit(True, "ایمیل با موفقیت ارسال شد.")
            logger.info(f"ایمیل به {self.to_email} ارسال شد.")
        except Exception as e:
            self.finished.emit(False, f"خطا در ارسال ایمیل: {str(e)}")
            logger.error(f"خطا در ارسال ایمیل: {e}")

def show_notification(title, message):
    if PLYER_AVAILABLE:
        try:
            notification.notify(title=title, message=message, app_name='پایان‌نامه', timeout=5)
        except:
            pass

# تابع تبدیل تاریخ میلادی به شمسی
def to_jalali(date_obj):
    if JALALI_AVAILABLE and date_obj:
        if isinstance(date_obj, (datetime, QDate)):
            if isinstance(date_obj, QDate):
                date_obj = date_obj.toPyDate()
            j = jdatetime.date.fromgregorian(date=date_obj)
            return j.strftime("%Y/%m/%d")
        return str(date_obj)
    return str(date_obj)

def to_jalali_datetime(dt):
    if JALALI_AVAILABLE and dt:
        j = jdatetime.datetime.fromgregorian(datetime=dt)
        return j.strftime("%Y/%m/%d %H:%M")
    return str(dt)

# ----------------------------------------------------------------------
# استایل‌های QSS
# ----------------------------------------------------------------------
DARK_STYLE = """
QMainWindow { background-color: #1e1e2f; }
QLabel { color: #f0f0f0; font-family: 'Vazirmatn'; font-size: 10pt; }
QPushButton {
    border-radius: 8px;
    padding: 8px 16px;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #4a90d9, stop:1 #2b5f9e);
    color: white;
    font-family: 'Vazirmatn';
    font-size: 10pt;
    border: none;
}
QPushButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #5aa0e9, stop:1 #3b6fae);
}
QPushButton:pressed {
    background: #1f4a7a;
}
QLineEdit, QTextEdit, QComboBox, QSpinBox, QDoubleSpinBox, QDateEdit {
    border-radius: 4px;
    padding: 4px;
    background-color: #2d2d44;
    color: #f0f0f0;
    font-family: 'Vazirmatn';
    border: 1px solid #4a4a6a;
}
QTableWidget {
    background-color: #2d2d44;
    color: #f0f0f0;
    gridline-color: #4a4a6a;
    font-family: 'Vazirmatn';
}
QHeaderView::section {
    background-color: #3d3d5c;
    color: white;
    padding: 4px;
    border: 1px solid #4a4a6a;
}
QTabWidget::pane { background: #2d2d44; border: 1px solid #4a4a6a; }
QTabBar::tab {
    background: #3d3d5c;
    color: white;
    padding: 8px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    margin-right: 2px;
}
QTabBar::tab:selected { background: #5a5a8a; }
QStatusBar { background-color: #2d2d44; color: #f0f0f0; }
QGroupBox {
    border: 1px solid #4a4a6a;
    border-radius: 6px;
    margin-top: 12px;
    font-family: 'Vazirmatn';
    color: #f0f0f0;
}
QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }
"""

LIGHT_STYLE = """
QMainWindow { background-color: #f5f5f5; }
QLabel { color: #333333; font-family: 'Vazirmatn'; font-size: 10pt; }
QPushButton {
    border-radius: 8px;
    padding: 8px 16px;
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #5a9bd4, stop:1 #3871a6);
    color: white;
    font-family: 'Vazirmatn';
    font-size: 10pt;
    border: none;
}
QPushButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #6aabe4, stop:1 #4881b6);
}
QPushButton:pressed { background: #2d5a8c; }
QLineEdit, QTextEdit, QComboBox, QSpinBox, QDoubleSpinBox, QDateEdit {
    border-radius: 4px;
    padding: 4px;
    background-color: white;
    color: #333333;
    font-family: 'Vazirmatn';
    border: 1px solid #cccccc;
}
QTableWidget {
    background-color: white;
    color: #333333;
    gridline-color: #cccccc;
    font-family: 'Vazirmatn';
}
QHeaderView::section {
    background-color: #e0e0e0;
    color: #333333;
    padding: 4px;
    border: 1px solid #cccccc;
}
QTabWidget::pane { background: white; border: 1px solid #cccccc; }
QTabBar::tab {
    background: #e0e0e0;
    color: #333333;
    padding: 8px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    margin-right: 2px;
}
QTabBar::tab:selected { background: #ffffff; }
QStatusBar { background-color: #f0f0f0; color: #333333; }
QGroupBox {
    border: 1px solid #cccccc;
    border-radius: 6px;
    margin-top: 12px;
    font-family: 'Vazirmatn';
    color: #333333;
}
QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }
"""

# ----------------------------------------------------------------------
# پنجره اصلی
# ----------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("سیستم جامع مدیریت پایان‌نامه")
        self.resize(1200, 800)
        self.current_user = None
        self.current_theme = "dark"
        self.apply_style(DARK_STYLE)

        # تنظیم فونت پیش‌فرض
        font = QFont("Vazirmatn", 10)
        self.setFont(font)

        # نوار وضعیت
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_label = QLabel("آماده")
        self.status_bar.addPermanentWidget(self.status_label)

        # نوار ابزار
        self.create_toolbar()

        # استک اصلی
        self.central_stack = QStackedWidget()
        self.setCentralWidget(self.central_stack)

        # صفحه‌ها
        self.login_page = LoginPage(self)
        self.registration_page = RegistrationPage(self)
        self.admin_panel = None
        self.professor_panel = None
        self.student_panel = None

        self.central_stack.addWidget(self.login_page)
        self.central_stack.addWidget(self.registration_page)

        # راه‌اندازی تایمر برای بروزرسانی ساعت
        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self.update_status_clock)
        self.clock_timer.start(1000)

        # بک‌آپ خودکار
        self.backup_timer = QTimer(self)
        self.backup_timer.timeout.connect(self.auto_backup)
        self.backup_timer.start(30 * 60 * 1000)  # هر 30 دقیقه

    def create_toolbar(self):
        toolbar = QToolBar("نوار ابزار")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        theme_btn = QPushButton("تغییر تم")
        theme_btn.clicked.connect(self.toggle_theme)
        toolbar.addWidget(theme_btn)

        toolbar.addSeparator()
        logout_btn = QPushButton("خروج")
        logout_btn.clicked.connect(self.logout)
        toolbar.addWidget(logout_btn)

    def toggle_theme(self):
        if self.current_theme == "dark":
            self.apply_style(LIGHT_STYLE)
            self.current_theme = "light"
        else:
            self.apply_style(DARK_STYLE)
            self.current_theme = "dark"

    def apply_style(self, style):
        self.setStyleSheet(style)
        # تنظیم فونت password fields
        # (در جای خود)

    def update_status_clock(self):
        now = datetime.now()
        jalali_str = to_jalali_datetime(now)
        self.status_label.setText(f"آخرین عملیات: {jalali_str}")

    def auto_backup(self):
        try:
            backup_path = BACKUP_DIR / f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
            shutil.copy2(DB_PATH, backup_path)
            # نگهداری 10 نسخه آخر
            backups = sorted(BACKUP_DIR.glob("backup_*.db"))
            if len(backups) > 10:
                for old in backups[:-10]:
                    old.unlink()
            logger.info("بک‌آپ خودکار انجام شد.")
        except Exception as e:
            logger.error(f"خطا در بک‌آپ خودکار: {e}")

    def login_successful(self, user):
        self.current_user = user
        role = user.role
        if role == 'admin':
            if self.admin_panel is None:
                self.admin_panel = AdminPanel(self)
                self.central_stack.addWidget(self.admin_panel)
            self.central_stack.setCurrentWidget(self.admin_panel)
            self.admin_panel.refresh()
        elif role == 'professor':
            if self.professor_panel is None:
                self.professor_panel = ProfessorPanel(self)
                self.central_stack.addWidget(self.professor_panel)
            self.central_stack.setCurrentWidget(self.professor_panel)
            self.professor_panel.refresh()
        elif role == 'student':
            if self.student_panel is None:
                self.student_panel = StudentPanel(self)
                self.central_stack.addWidget(self.student_panel)
            self.central_stack.setCurrentWidget(self.student_panel)
            self.student_panel.refresh()

    def logout(self):
        self.current_user = None
        self.central_stack.setCurrentWidget(self.login_page)
        self.status_label.setText("خروج انجام شد")
        # پاک‌سازی پنل‌ها برای بارگذاری مجدد
        if self.admin_panel:
            self.central_stack.removeWidget(self.admin_panel)
            self.admin_panel.deleteLater()
            self.admin_panel = None
        if self.professor_panel:
            self.central_stack.removeWidget(self.professor_panel)
            self.professor_panel.deleteLater()
            self.professor_panel = None
        if self.student_panel:
            self.central_stack.removeWidget(self.student_panel)
            self.student_panel.deleteLater()
            self.student_panel = None

# ----------------------------------------------------------------------
# صفحه ورود
# ----------------------------------------------------------------------
class LoginPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main = main_window
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title = QLabel("ورود به سیستم مدیریت پایان‌نامه")
        title.setStyleSheet("font-size: 18pt; font-weight: bold;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        role_label = QLabel("نقش:")
        self.role_combo = QComboBox()
        self.role_combo.addItems(["دانشجو", "استاد", "ادمین"])

        username_label = QLabel("نام کاربری:")
        self.username_edit = QLineEdit()
        password_label = QLabel("رمز عبور:")
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.show_pwd_cb = QCheckBox("نمایش رمز")
        self.show_pwd_cb.toggled.connect(lambda checked: self.password_edit.setEchoMode(QLineEdit.EchoMode.Normal if checked else QLineEdit.EchoMode.Password))

        # دکمه‌ها
        login_btn = QPushButton("ورود")
        login_btn.clicked.connect(self.login)
        register_btn = QPushButton("ثبت‌نام دانشجو")
        register_btn.clicked.connect(lambda: self.main.central_stack.setCurrentWidget(self.main.registration_page))
        recover_btn = QPushButton("بازیابی رمز")
        recover_btn.clicked.connect(self.recover_password)

        form_layout = QFormLayout()
        form_layout.addRow(role_label, self.role_combo)
        form_layout.addRow(username_label, self.username_edit)
        form_layout.addRow(password_label, self.password_edit)
        form_layout.addRow(self.show_pwd_cb)

        btn_layout = QHBoxLayout()
        btn_layout.addWidget(login_btn)
        btn_layout.addWidget(register_btn)
        btn_layout.addWidget(recover_btn)

        layout.addStretch()
        layout.addWidget(title)
        layout.addSpacing(20)
        layout.addLayout(form_layout)
        layout.addLayout(btn_layout)
        layout.addStretch()
        self.setLayout(layout)

    def login(self):
        username = self.username_edit.text().strip()
        password = self.password_edit.text().strip()
        role_map = {"دانشجو": "student", "استاد": "professor", "ادمین": "admin"}
        role = role_map[self.role_combo.currentText()]

        session = get_session()
        user = session.query(User).filter_by(username=username, role=role).first()
        if user and user.password_hash == hash_password(password) and user.is_active:
            self.main.login_successful(user)
            logger.info(f"کاربر {username} وارد شد.")
        else:
            QMessageBox.warning(self, "خطا", "نام کاربری یا رمز عبور اشتباه است یا حساب غیرفعال.")
        session.close()

    def recover_password(self):
        # پیاده‌سازی ساده: دریافت ایمیل و ارسال رمز جدید
        email, ok = QInputDialog.getText(self, "بازیابی رمز", "لطفاً ایمیل خود را وارد کنید:")
        if ok and email:
            session = get_session()
            # جستجو در دانشجویان و اساتید
            student = session.query(Student).filter_by(email=email).first()
            professor = session.query(Professor).filter_by(email=email).first()
            user = None
            if student:
                user = session.query(User).filter_by(related_id=student.id, role='student').first()
            elif professor:
                user = session.query(User).filter_by(related_id=professor.id, role='professor').first()
            if user:
                new_pwd = generate_temp_password()
                user.password_hash = hash_password(new_pwd)
                session.commit()
                # ارسال ایمیل
                self.send_password_email(email, new_pwd)
                QMessageBox.information(self, "موفقیت", "رمز جدید به ایمیل شما ارسال شد.")
            else:
                QMessageBox.warning(self, "خطا", "حسابی با این ایمیل یافت نشد.")
            session.close()

    def send_password_email(self, email, password):
        if can_send_email():
            subject = "بازیابی رمز عبور - سیستم پایان‌نامه"
            body = f"رمز عبور جدید شما: {password}\nلطفاً پس از ورود آن را تغییر دهید."
            self.email_thread = EmailSender(email, subject, body)
            self.email_thread.finished.connect(lambda s, m: print(m))
            self.email_thread.start()
        else:
            QMessageBox.information(self, "تذکر", f"امکان ارسال ایمیل وجود ندارد. رمز جدید: {password}")

# ----------------------------------------------------------------------
# صفحه ثبت‌نام دانشجو
# ----------------------------------------------------------------------
class RegistrationPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main = main_window
        self.photo_path = None
        layout = QVBoxLayout()
        title = QLabel("فرم ثبت‌نام دانشجو")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 14pt; font-weight: bold;")

        form_layout = QFormLayout()
        self.fields = {}
        field_names = [
            ("نام", "first_name"), ("نام خانوادگی", "last_name"), ("کد ملی", "national_id"),
            ("شماره دانشجویی", "student_id"), ("رشته", "field_of_study"), ("مقطع", "grade"),
            ("سال ورود", "entry_year"), ("ایمیل", "email"), ("شماره تماس", "phone")
        ]
        for label, key in field_names:
            edit = QLineEdit()
            self.fields[key] = edit
            form_layout.addRow(label + ":", edit)

        # عکس
        photo_layout = QHBoxLayout()
        self.photo_label = QLabel("عکس انتخاب نشده")
        self.photo_label.setFixedSize(100, 100)
        self.photo_label.setStyleSheet("border: 1px solid gray;")
        upload_btn = QPushButton("آپلود عکس")
        upload_btn.clicked.connect(self.upload_photo)
        cam_btn = QPushButton("عکس با وب‌کم")
        cam_btn.clicked.connect(self.capture_camera)
        photo_layout.addWidget(self.photo_label)
        photo_layout.addWidget(upload_btn)
        photo_layout.addWidget(cam_btn)

        register_btn = QPushButton("ثبت‌نام")
        register_btn.clicked.connect(self.register)
        back_btn = QPushButton("بازگشت به ورود")
        back_btn.clicked.connect(lambda: self.main.central_stack.setCurrentWidget(self.main.login_page))

        layout.addWidget(title)
        layout.addLayout(form_layout)
        layout.addLayout(photo_layout)
        layout.addWidget(register_btn)
        layout.addWidget(back_btn)
        self.setLayout(layout)

    def upload_photo(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "انتخاب عکس", "", "تصاویر (*.png *.jpg *.jpeg)")
        if file_path:
            self.photo_path = file_path
            pixmap = QPixmap(file_path).scaled(100, 100, Qt.AspectRatioMode.KeepAspectRatio)
            self.photo_label.setPixmap(pixmap)

    def capture_camera(self):
        try:
            import cv2
            cap = cv2.VideoCapture(0)
            if not cap.isOpened():
                raise Exception("وب‌کم در دسترس نیست")
            ret, frame = cap.read()
            if ret:
                # ذخیره موقت
                temp_path = UPLOAD_DIR / f"webcam_{datetime.now().strftime('%Y%m%d%H%M%S')}.jpg"
                cv2.imwrite(str(temp_path), frame)
                self.photo_path = str(temp_path)
                pixmap = QPixmap(str(temp_path)).scaled(100, 100, Qt.AspectRatioMode.KeepAspectRatio)
                self.photo_label.setPixmap(pixmap)
                cap.release()
            else:
                raise Exception("گرفتن عکس ناموفق")
        except Exception as e:
            QMessageBox.warning(self, "خطا", f"امکان استفاده از وب‌کم وجود ندارد: {str(e)}. لطفاً عکس را آپلود کنید.")

    def register(self):
        # اعتبارسنجی
        for key, edit in self.fields.items():
            if not edit.text().strip():
                QMessageBox.warning(self, "خطا", f"فیلد {key} الزامی است.")
                return
        national_id = self.fields['national_id'].text().strip()
        student_id = self.fields['student_id'].text().strip()
        if not national_id.isdigit() or len(national_id) != 10:
            QMessageBox.warning(self, "خطا", "کد ملی باید 10 رقم باشد.")
            return
        try:
            entry_year = int(self.fields['entry_year'].text().strip())
        except ValueError:
            QMessageBox.warning(self, "خطا", "سال ورود نامعتبر.")
            return

        session = get_session()
        existing = session.query(Student).filter_by(national_id=national_id).first()
        if existing:
            QMessageBox.warning(self, "خطا", "کد ملی تکراری است.")
            session.close()
            return

        # ذخیره عکس
        photo_path_final = None
        if self.photo_path:
            dest = UPLOAD_DIR / f"{national_id}_photo{os.path.splitext(self.photo_path)[1]}"
            shutil.copy2(self.photo_path, dest)
            photo_path_final = str(dest)

        new_student = Student(
            national_id=national_id,
            student_id=student_id,
            first_name=self.fields['first_name'].text().strip(),
            last_name=self.fields['last_name'].text().strip(),
            field_of_study=self.fields['field_of_study'].text().strip(),
            grade=self.fields['grade'].text().strip(),
            entry_year=entry_year,
            email=self.fields['email'].text().strip(),
            phone=self.fields['phone'].text().strip(),
            photo_path=photo_path_final,
            is_approved_by_admin=False
        )
        session.add(new_student)
        session.commit()
        session.close()
        QMessageBox.information(self, "موفقیت", "ثبت‌نام شما با موفقیت انجام شد. پس از تأیید ادمین می‌توانید وارد شوید.")
        # بازگشت به لاگین
        self.main.central_stack.setCurrentWidget(self.main.login_page)

# ----------------------------------------------------------------------
# پنل ادمین
# ----------------------------------------------------------------------
class AdminPanel(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main = main_window
        self.session = None
        layout = QVBoxLayout()
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)
        self.setLayout(layout)

        # ساخت تب‌ها
        self.student_tab = QWidget()
        self.professor_tab = QWidget()
        self.archive_tab = QWidget()
        self.report_tab = QWidget()
        self.backup_tab = QWidget()

        self.tabs.addTab(self.student_tab, "مدیریت دانشجویان")
        self.tabs.addTab(self.professor_tab, "مدیریت اساتید")
        self.tabs.addTab(self.archive_tab, "بایگانی پایان‌نامه‌ها")
        self.tabs.addTab(self.report_tab, "گزارش‌ها")
        self.tabs.addTab(self.backup_tab, "بک‌آپ")

        self.setup_student_tab()
        self.setup_professor_tab()
        self.setup_archive_tab()
        self.setup_report_tab()
        self.setup_backup_tab()

    def refresh(self):
        self.load_students()
        self.load_professors()
        self.load_archive()
        # گزارش‌ها بعداً لود می‌شوند

    def get_session(self):
        if self.session is None or not self.session.is_active:
            self.session = get_session()
        return self.session

    # --------------------- تب دانشجویان ---------------------
    def setup_student_tab(self):
        layout = QVBoxLayout()
        filter_layout = QHBoxLayout()
        self.student_filter_combo = QComboBox()
        self.student_filter_combo.addItems(["همه", "در انتظار تأیید", "تأیید شده", "رد شده"])
        self.student_filter_combo.currentTextChanged.connect(self.load_students)
        self.student_search = QLineEdit()
        self.student_search.setPlaceholderText("جستجو...")
        self.student_search.textChanged.connect(self.load_students)

        filter_layout.addWidget(QLabel("فیلتر:"))
        filter_layout.addWidget(self.student_filter_combo)
        filter_layout.addWidget(QLabel("جستجو:"))
        filter_layout.addWidget(self.student_search)

        self.student_table = QTableWidget()
        self.student_table.setColumnCount(9)
        self.student_table.setHorizontalHeaderLabels(["کد ملی", "شماره دانشجویی", "نام", "نام خانوادگی", "رشته", "مقطع", "ایمیل", "وضعیت", "عملیات"])
        self.student_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        btn_layout = QHBoxLayout()
        approve_btn = QPushButton("تأیید")
        approve_btn.clicked.connect(lambda: self.process_student_action("approve"))
        edit_btn = QPushButton("ویرایش")
        edit_btn.clicked.connect(lambda: self.process_student_action("edit"))
        reject_btn = QPushButton("رد")
        reject_btn.clicked.connect(lambda: self.process_student_action("reject"))
        delete_btn = QPushButton("حذف")
        delete_btn.clicked.connect(lambda: self.process_student_action("delete"))
        view_photo_btn = QPushButton("مشاهده عکس")
        view_photo_btn.clicked.connect(self.view_student_photo)

        btn_layout.addWidget(approve_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(reject_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addWidget(view_photo_btn)

        layout.addLayout(filter_layout)
        layout.addWidget(self.student_table)
        layout.addLayout(btn_layout)
        self.student_tab.setLayout(layout)

    def load_students(self):
        session = self.get_session()
        filter_text = self.student_filter_combo.currentText()
        search = self.student_search.text().strip()
        query = session.query(Student)
        if filter_text == "در انتظار تأیید":
            query = query.filter_by(is_approved_by_admin=False)
        elif filter_text == "تأیید شده":
            query = query.filter_by(is_approved_by_admin=True)
        # رد شده منطق خاصی ندارد، می‌توان وضعیت دیگری اضافه کرد
        if search:
            query = query.filter(
                (Student.first_name.contains(search)) | (Student.last_name.contains(search)) |
                (Student.student_id.contains(search)) | (Student.national_id.contains(search))
            )
        students = query.all()
        self.student_table.setRowCount(len(students))
        for i, stu in enumerate(students):
            self.student_table.setItem(i, 0, QTableWidgetItem(stu.national_id))
            self.student_table.setItem(i, 1, QTableWidgetItem(stu.student_id))
            self.student_table.setItem(i, 2, QTableWidgetItem(stu.first_name))
            self.student_table.setItem(i, 3, QTableWidgetItem(stu.last_name))
            self.student_table.setItem(i, 4, QTableWidgetItem(stu.field_of_study))
            self.student_table.setItem(i, 5, QTableWidgetItem(stu.grade))
            self.student_table.setItem(i, 6, QTableWidgetItem(stu.email))
            status = "تأیید شده" if stu.is_approved_by_admin else "در انتظار"
            self.student_table.setItem(i, 7, QTableWidgetItem(status))
            btn = QPushButton("عملیات")
            btn.clicked.connect(lambda _, r=i: self.show_student_actions(r))
            self.student_table.setCellWidget(i, 8, btn)
        self.student_table.resizeColumnsToContents()

    def show_student_actions(self, row):
        # منو برای عملیات
        pass

    def process_student_action(self, action):
        row = self.student_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "خطا", "لطفاً یک دانشجو انتخاب کنید.")
            return
        national_id = self.student_table.item(row, 0).text()
        session = self.get_session()
        student = session.query(Student).filter_by(national_id=national_id).first()
        if not student:
            return
        if action == "approve":
            student.is_approved_by_admin = True
            # ایجاد کاربر و رمز موقت
            user = session.query(User).filter_by(role='student', related_id=student.id).first()
            temp_pwd = generate_temp_password()
            if not user:
                user = User(username=student.national_id, password_hash=hash_password(temp_pwd), role='student', related_id=student.id)
                session.add(user)
            else:
                user.password_hash = hash_password(temp_pwd)
            session.commit()
            self.send_password_email(student.email, temp_pwd)
            QMessageBox.information(self, "موفقیت", "دانشجو تأیید و رمز به ایمیل ارسال شد.")
            self.load_students()
        elif action == "edit":
            self.edit_student(student)
        elif action == "reject":
            # می‌توانیم یک فیلد reject اضافه کنیم، فعلاً کاربر را غیرفعال می‌کنیم
            user = session.query(User).filter_by(role='student', related_id=student.id).first()
            if user:
                user.is_active = False
            session.commit()
            QMessageBox.information(self, "موفقیت", "حساب دانشجو رد و غیرفعال شد.")
            self.load_students()
        elif action == "delete":
            reply = QMessageBox.question(self, "تأیید", "آیا مطمئن به حذف هستید؟", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                user = session.query(User).filter_by(role='student', related_id=student.id).first()
                if user:
                    session.delete(user)
                session.delete(student)
                session.commit()
                self.load_students()

    def edit_student(self, student):
        dialog = QDialog(self)
        dialog.setWindowTitle("ویرایش دانشجو")
        layout = QFormLayout()
        edits = {}
        for field in ['first_name', 'last_name', 'field_of_study', 'grade', 'email', 'phone']:
            edit = QLineEdit(getattr(student, field))
            edits[field] = edit
            layout.addRow(field, edit)
        entry_edit = QSpinBox()
        entry_edit.setValue(student.entry_year)
        layout.addRow("سال ورود", entry_edit)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addRow(button_box)
        dialog.setLayout(layout)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            for field in edits:
                setattr(student, field, edits[field].text())
            student.entry_year = entry_edit.value()
            self.get_session().commit()
            self.load_students()

    def view_student_photo(self):
        row = self.student_table.currentRow()
        if row < 0:
            return
        national_id = self.student_table.item(row, 0).text()
        session = self.get_session()
        student = session.query(Student).filter_by(national_id=national_id).first()
        if student and student.photo_path and os.path.exists(student.photo_path):
            pixmap = QPixmap(student.photo_path)
            dlg = QDialog(self)
            dlg.setWindowTitle("عکس دانشجو")
            layout = QVBoxLayout()
            lbl = QLabel()
            lbl.setPixmap(pixmap.scaled(300, 300, Qt.AspectRatioMode.KeepAspectRatio))
            layout.addWidget(lbl)
            dlg.setLayout(layout)
            dlg.exec()

    def send_password_email(self, email, password):
        subject = "اطلاعات ورود - سیستم پایان‌نامه"
        body = f"نام کاربری شما: {email}\nرمز عبور موقت: {password}\nلطفاً پس از ورود رمز خود را تغییر دهید."
        if can_send_email():
            self.email_thread = EmailSender(email, subject, body)
            self.email_thread.start()
        else:
            QMessageBox.information(self, "تذکر", f"ایمیل ارسال نشد. رمز: {password}")

    # --------------------- تب اساتید ---------------------
    def setup_professor_tab(self):
        layout = QVBoxLayout()
        self.prof_table = QTableWidget()
        self.prof_table.setColumnCount(6)
        self.prof_table.setHorizontalHeaderLabels(["کد استادی", "نام", "نام خانوادگی", "تخصص", "ایمیل", "ظرفیت"])
        btn_layout = QHBoxLayout()
        add_btn = QPushButton("افزودن استاد")
        add_btn.clicked.connect(self.add_professor)
        edit_btn = QPushButton("ویرایش")
        edit_btn.clicked.connect(lambda: self.process_prof_action("edit"))
        delete_btn = QPushButton("حذف")
        delete_btn.clicked.connect(lambda: self.process_prof_action("delete"))
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(delete_btn)
        layout.addWidget(self.prof_table)
        layout.addLayout(btn_layout)
        self.professor_tab.setLayout(layout)

    def load_professors(self):
        session = self.get_session()
        professors = session.query(Professor).all()
        self.prof_table.setRowCount(len(professors))
        for i, prof in enumerate(professors):
            self.prof_table.setItem(i, 0, QTableWidgetItem(prof.professor_code))
            self.prof_table.setItem(i, 1, QTableWidgetItem(prof.first_name))
            self.prof_table.setItem(i, 2, QTableWidgetItem(prof.last_name))
            self.prof_table.setItem(i, 3, QTableWidgetItem(prof.expertise))
            self.prof_table.setItem(i, 4, QTableWidgetItem(prof.email))
            self.prof_table.setItem(i, 5, QTableWidgetItem(str(prof.max_students)))
        self.prof_table.resizeColumnsToContents()

    def add_professor(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("افزودن استاد جدید")
        layout = QFormLayout()
        edits = {}
        for label, key in [("کد استادی", "professor_code"), ("نام", "first_name"), ("نام خانوادگی", "last_name"),
                           ("تخصص", "expertise"), ("ایمیل", "email"), ("شماره تماس", "phone")]:
            edit = QLineEdit()
            edits[key] = edit
            layout.addRow(label, edit)
        max_edit = QSpinBox()
        max_edit.setRange(1, 20)
        max_edit.setValue(5)
        layout.addRow("حداکثر دانشجو", max_edit)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addRow(button_box)
        dialog.setLayout(layout)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            code = edits['professor_code'].text().strip()
            session = self.get_session()
            if session.query(Professor).filter_by(professor_code=code).first():
                QMessageBox.warning(self, "خطا", "کد استادی تکراری است.")
                return
            prof = Professor(
                professor_code=code,
                first_name=edits['first_name'].text(),
                last_name=edits['last_name'].text(),
                expertise=edits['expertise'].text(),
                email=edits['email'].text(),
                phone=edits['phone'].text(),
                max_students=max_edit.value()
            )
            session.add(prof)
            session.flush()
            # ایجاد کاربر
            temp_pwd = generate_temp_password()
            user = User(username=code, password_hash=hash_password(temp_pwd), role='professor', related_id=prof.id)
            session.add(user)
            session.commit()
            self.send_password_email(prof.email, temp_pwd)
            QMessageBox.information(self, "موفقیت", "استاد اضافه شد و رمز به ایمیل ارسال شد.")
            self.load_professors()

    def process_prof_action(self, action):
        row = self.prof_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "خطا", "استادی انتخاب نشده.")
            return
        code = self.prof_table.item(row, 0).text()
        session = self.get_session()
        prof = session.query(Professor).filter_by(professor_code=code).first()
        if not prof:
            return
        if action == "edit":
            dialog = QDialog(self)
            dialog.setWindowTitle("ویرایش استاد")
            layout = QFormLayout()
            edits = {}
            for field, label in [("first_name", "نام"), ("last_name", "نام خانوادگی"), ("expertise", "تخصص"), ("email", "ایمیل"), ("phone", "تلفن")]:
                edit = QLineEdit(getattr(prof, field))
                edits[field] = edit
                layout.addRow(label, edit)
            max_edit = QSpinBox()
            max_edit.setValue(prof.max_students)
            layout.addRow("حداکثر دانشجو", max_edit)
            button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
            button_box.accepted.connect(dialog.accept)
            button_box.rejected.connect(dialog.reject)
            layout.addRow(button_box)
            dialog.setLayout(layout)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                for field in edits:
                    setattr(prof, field, edits[field].text())
                prof.max_students = max_edit.value()
                session.commit()
                self.load_professors()
        elif action == "delete":
            reply = QMessageBox.question(self, "تأیید", "حذف استاد؟", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                user = session.query(User).filter_by(role='professor', related_id=prof.id).first()
                if user:
                    session.delete(user)
                session.delete(prof)
                session.commit()
                self.load_professors()

    # --------------------- تب بایگانی ---------------------
    def setup_archive_tab(self):
        layout = QVBoxLayout()
        filter_layout = QHBoxLayout()
        self.archive_search = QLineEdit()
        self.archive_search.setPlaceholderText("جستجو بر اساس نام، شماره دانشجویی...")
        self.archive_search.textChanged.connect(self.load_archive)
        self.archive_year_combo = QComboBox()
        self.archive_year_combo.addItem("همه سال‌ها")
        self.archive_year_combo.currentTextChanged.connect(self.load_archive)
        filter_layout.addWidget(QLabel("جستجو:"))
        filter_layout.addWidget(self.archive_search)
        filter_layout.addWidget(QLabel("سال دفاع:"))
        filter_layout.addWidget(self.archive_year_combo)

        self.archive_table = QTableWidget()
        self.archive_table.setColumnCount(6)
        self.archive_table.setHorizontalHeaderLabels(["دانشجو", "شماره دانشجویی", "عنوان", "تاریخ دفاع", "نمره", "عملیات"])

        btn_layout = QHBoxLayout()
        download_btn = QPushButton("دریافت بایگانی")
        download_btn.clicked.connect(self.download_archive)
        excel_btn = QPushButton("خروجی اکسل")
        excel_btn.clicked.connect(self.export_archive_excel)
        btn_layout.addWidget(download_btn)
        btn_layout.addWidget(excel_btn)

        layout.addLayout(filter_layout)
        layout.addWidget(self.archive_table)
        layout.addLayout(btn_layout)
        self.archive_tab.setLayout(layout)

    def load_archive(self):
        session = self.get_session()
        query = session.query(Thesis).filter_by(is_archived=True)
        search = self.archive_search.text().strip()
        if search:
            query = query.join(Student).filter(
                (Student.first_name.contains(search)) | (Student.last_name.contains(search)) |
                (Student.student_id.contains(search)) | (Thesis.title.contains(search))
            )
        # فیلتر سال دفاع
        year = self.archive_year_combo.currentText()
        if year != "همه سال‌ها":
            # فرض سال شمسی؟ ساده‌سازی می‌کنیم
            pass
        theses = query.all()
        self.archive_table.setRowCount(len(theses))
        for i, thesis in enumerate(theses):
            student = thesis.student
            self.archive_table.setItem(i, 0, QTableWidgetItem(f"{student.first_name} {student.last_name}"))
            self.archive_table.setItem(i, 1, QTableWidgetItem(student.student_id))
            self.archive_table.setItem(i, 2, QTableWidgetItem(thesis.title))
            self.archive_table.setItem(i, 3, QTableWidgetItem(str(thesis.defense_date) if thesis.defense_date else ""))
            self.archive_table.setItem(i, 4, QTableWidgetItem(str(thesis.final_grade) if thesis.final_grade else ""))
            btn = QPushButton("دانلود")
            btn.clicked.connect(lambda _, t=thesis: self.download_archive_item(t))
            self.archive_table.setCellWidget(i, 5, btn)
        self.archive_table.resizeColumnsToContents()

    def download_archive_item(self, thesis):
        # ساخت پوشه آرشیو
        student = thesis.student
        folder_name = f"{student.student_id}_{thesis.title}"
        dest_dir = ARCHIVE_DIR / folder_name
        dest_dir.mkdir(exist_ok=True)
        # کپی فایل‌ها
        for file in thesis.files:
            if os.path.exists(file.stored_path):
                shutil.copy(file.stored_path, dest_dir / file.original_name)
        # metadata
        meta = {
            "student": {
                "first_name": student.first_name,
                "last_name": student.last_name,
                "student_id": student.student_id,
                "field": student.field_of_study
            },
            "supervisor": f"{thesis.supervisor.first_name} {thesis.supervisor.last_name}" if thesis.supervisor else "",
            "defense_date": str(thesis.defense_date),
            "final_grade": thesis.final_grade,
            "meetings": [{"date": str(m.meeting_date), "topic": m.topic} for m in thesis.meetings]
        }
        with open(dest_dir / "metadata.json", "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        QMessageBox.information(self, "موفقیت", f"بایگانی در {dest_dir} ایجاد شد.")

    def download_archive(self):
        row = self.archive_table.currentRow()
        if row < 0:
            return
        # دریافت thesis مرتبط
        title = self.archive_table.item(row, 2).text()
        session = self.get_session()
        thesis = session.query(Thesis).filter_by(title=title, is_archived=True).first()
        if thesis:
            self.download_archive_item(thesis)

    def export_archive_excel(self):
        session = self.get_session()
        theses = session.query(Thesis).filter_by(is_archived=True).all()
        data = []
        for t in theses:
            data.append({
                "دانشجو": f"{t.student.first_name} {t.student.last_name}",
                "شماره دانشجویی": t.student.student_id,
                "عنوان": t.title,
                "تاریخ دفاع": str(t.defense_date),
                "نمره": t.final_grade
            })
        df = pd.DataFrame(data)
        path, _ = QFileDialog.getSaveFileName(self, "ذخیره اکسل", "", "Excel files (*.xlsx)")
        if path:
            df.to_excel(path, index=False)
            QMessageBox.information(self, "موفقیت", "خروجی اکسل ذخیره شد.")

    # --------------------- تب گزارش‌ها ---------------------
    def setup_report_tab(self):
        layout = QVBoxLayout()
        self.report_widget = pg.GraphicsLayoutWidget()
        layout.addWidget(self.report_widget)
        refresh_btn = QPushButton("به‌روزرسانی گزارش")
        refresh_btn.clicked.connect(self.update_reports)
        layout.addWidget(refresh_btn)
        self.report_tab.setLayout(layout)

    def update_reports(self):
        self.report_widget.clear()
        # نمودار میله‌ای: تعداد دانشجو به تفکیک استاد
        session = self.get_session()
        professors = session.query(Professor).all()
        names = [f"{p.last_name}" for p in professors]
        counts = [len(p.theses_supervised) for p in professors]
        if names:
            p1 = self.report_widget.addPlot(title="تعداد دانشجویان تحت راهنمایی")
            bar = pg.BarGraphItem(x=list(range(len(names))), height=counts, width=0.6, brush='b')
            p1.addItem(bar)
            p1.getAxis('bottom').setTicks([list(enumerate(names))])
        # نمودار دیگر خطی پیشرفت (نمونه کلی)
        p2 = self.report_widget.addPlot(title="میانگین پیشرفت دانشجویان")
        # اینجا می‌توان داده‌های تاریخی را شبیه‌سازی کرد
        p2.plot([1,2,3,4], [20,40,55,70], pen='g')

    # --------------------- تب بک‌آپ ---------------------
    def setup_backup_tab(self):
        layout = QVBoxLayout()
        backup_btn = QPushButton("بک‌آپ دستی")
        backup_btn.clicked.connect(self.main.auto_backup)
        restore_btn = QPushButton("بازیابی بک‌آپ")
        restore_btn.clicked.connect(self.restore_backup)
        layout.addWidget(backup_btn)
        layout.addWidget(restore_btn)
        self.backup_tab.setLayout(layout)

    def restore_backup(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل بک‌آپ", str(BACKUP_DIR), "Database (*.db)")
        if file_path:
            reply = QMessageBox.question(self, "تأیید", "بازیابی باعث جایگزینی پایگاه داده فعلی می‌شود. ادامه؟", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                shutil.copy2(file_path, DB_PATH)
                QMessageBox.information(self, "موفقیت", "بازیابی انجام شد. لطفاً برنامه را مجدداً راه‌اندازی کنید.")

# ----------------------------------------------------------------------
# پنل استاد راهنما (بسیار کامل)
# ----------------------------------------------------------------------
class ProfessorPanel(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main = main_window
        self.session = None
        self.current_professor = None
        layout = QVBoxLayout()
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)
        self.setLayout(layout)

        self.dashboard_tab = QWidget()
        self.students_list_tab = QWidget()
        self.requests_tab = QWidget()
        self.detail_tab = QWidget()  # جزئیات دانشجو
        self.reports_tab = QWidget()

        self.tabs.addTab(self.dashboard_tab, "داشبورد")
        self.tabs.addTab(self.students_list_tab, "دانشجویان من")
        self.tabs.addTab(self.requests_tab, "درخواست‌های راهنمایی")
        self.tabs.addTab(self.reports_tab, "گزارش‌ها")

        self.setup_dashboard()
        self.setup_students_list()
        self.setup_requests()
        self.setup_reports()

    def refresh(self):
        user = self.main.current_user
        if user.role != 'professor':
            return
        self.session = get_session()
        self.current_professor = self.session.query(Professor).filter_by(id=user.related_id).first()
        if not self.current_professor:
            QMessageBox.critical(self, "خطا", "اطلاعات استاد یافت نشد.")
            self.main.logout()
            return
        self.load_dashboard()
        self.load_students_list()
        self.load_requests()

    def get_session(self):
        return self.session

    # --------------------- داشبورد ---------------------
    def setup_dashboard(self):
        self.dash_layout = QVBoxLayout()
        self.dashboard_tab.setLayout(self.dash_layout)

    def load_dashboard(self):
        # پاکسازی
        for i in reversed(range(self.dash_layout.count())):
            w = self.dash_layout.itemAt(i).widget()
            if w:
                w.deleteLater()
        prof = self.current_professor
        # کارت‌ها
        cards_layout = QHBoxLayout()
        card1 = self.create_card("تعداد دانشجویان", str(len(prof.theses_supervised)))
        cards_layout.addWidget(card1)
        card2 = self.create_card("جلسات امروز", "۰")
        cards_layout.addWidget(card2)
        self.dash_layout.addLayout(cards_layout)
        # هشدارها و یادآوری‌ها
        alerts_group = QGroupBox("یادآوری‌ها")
        alerts_layout = QVBoxLayout()
        # جلسات آینده
        upcoming_meetings = self.session.query(Meeting).filter(
            Meeting.professor_id == prof.id,
            Meeting.meeting_date >= datetime.now().date(),
            Meeting.meeting_date <= datetime.now().date() + timedelta(days=7)
        ).all()
        for m in upcoming_meetings:
            lbl = QLabel(f"جلسه با {m.thesis.student.first_name} در تاریخ {m.meeting_date}")
            alerts_layout.addWidget(lbl)
        alerts_group.setLayout(alerts_layout)
        self.dash_layout.addWidget(alerts_group)

    def create_card(self, title, value):
        frame = QFrame()
        frame.setStyleSheet("background-color: #3d3d5c; border-radius: 10px; padding: 10px;")
        layout = QVBoxLayout()
        lbl_title = QLabel(title)
        lbl_title.setStyleSheet("font-weight: bold; color: white;")
        lbl_val = QLabel(value)
        lbl_val.setStyleSheet("font-size: 18pt; color: white;")
        layout.addWidget(lbl_title)
        layout.addWidget(lbl_val)
        frame.setLayout(layout)
        return frame

    # --------------------- لیست دانشجویان ---------------------
    def setup_students_list(self):
        layout = QVBoxLayout()
        self.students_table = QTableWidget()
        self.students_table.setColumnCount(5)
        self.students_table.setHorizontalHeaderLabels(["نام", "شماره دانشجویی", "عنوان پایان‌نامه", "پیشرفت", "وضعیت"])
        self.students_table.cellClicked.connect(self.open_student_detail)
        layout.addWidget(self.students_table)
        self.students_list_tab.setLayout(layout)

    def load_students_list(self):
        prof = self.current_professor
        theses = prof.theses_supervised
        self.students_table.setRowCount(len(theses))
        for i, thesis in enumerate(theses):
            student = thesis.student
            self.students_table.setItem(i, 0, QTableWidgetItem(f"{student.first_name} {student.last_name}"))
            self.students_table.setItem(i, 1, QTableWidgetItem(student.student_id))
            self.students_table.setItem(i, 2, QTableWidgetItem(thesis.title or ""))
            self.students_table.setItem(i, 3, QTableWidgetItem(f"%{thesis.progress_percent}"))
            self.students_table.setItem(i, 4, QTableWidgetItem(thesis.project_status))
            # رنگ‌بندی
            if thesis.project_status == "دفاع شده":
                color = QColor(100, 200, 100)
            elif thesis.project_status == "متوقف":
                color = QColor(255, 100, 100)
            else:
                color = QColor(255, 255, 100)
            for col in range(5):
                self.students_table.item(i, col).setBackground(color)
        self.students_table.resizeColumnsToContents()

    def open_student_detail(self, row, col):
        prof = self.current_professor
        thesis = prof.theses_supervised[row]  # هم‌تراز
        self.show_student_detail(thesis)

    # --------------------- جزئیات دانشجو (پنجره مرکزی) ---------------------
    def show_student_detail(self, thesis):
        # ساخت یک تب جدید یا پنجره جدا
        if self.tabs.count() == 4:
            self.tabs.addTab(self.detail_tab, "جزئیات دانشجو")
        self.tabs.setCurrentWidget(self.detail_tab)
        # پاکسازی محتویات قبلی
        if self.detail_tab.layout():
            QWidget().setLayout(self.detail_tab.layout())  # trick
        layout = QVBoxLayout()
        # اطلاعات پایه
        student = thesis.student
        info_group = QGroupBox("اطلاعات دانشجو")
        info_layout = QGridLayout()
        info_layout.addWidget(QLabel("نام:"), 0, 0)
        info_layout.addWidget(QLabel(f"{student.first_name} {student.last_name}"), 0, 1)
        if student.photo_path and os.path.exists(student.photo_path):
            lbl_photo = QLabel()
            lbl_photo.setPixmap(QPixmap(student.photo_path).scaled(80, 80, Qt.AspectRatioMode.KeepAspectRatio))
            info_layout.addWidget(lbl_photo, 0, 2, 2, 1)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # پایان‌نامه
        thesis_group = QGroupBox("پایان‌نامه")
        t_layout = QFormLayout()
        self.thesis_title_edit = QLineEdit(thesis.title or "")
        t_layout.addRow("عنوان:", self.thesis_title_edit)
        save_title_btn = QPushButton("ذخیره عنوان")
        save_title_btn.clicked.connect(lambda: self.save_thesis_title(thesis))
        t_layout.addRow(save_title_btn)
        # تاریخ‌ها
        self.proposal_date_edit = QDateEdit()
        if thesis.proposal_approval_date:
            self.proposal_date_edit.setDate(QDate(thesis.proposal_approval_date))
        t_layout.addRow("تاریخ تصویب پروپوزال:", self.proposal_date_edit)
        # دکمه‌های مدیریت
        thesis_group.setLayout(t_layout)
        layout.addWidget(thesis_group)

        # تب‌های زیرمجموعه
        subtabs = QTabWidget()
        self.detail_subtabs = subtabs

        # تب جلسات
        meetings_tab = QWidget()
        self.setup_meetings_tab(meetings_tab, thesis)
        subtabs.addTab(meetings_tab, "جلسات")

        # تب فصل‌ها و پیشرفت
        chapters_tab = QWidget()
        self.setup_chapters_tab(chapters_tab, thesis)
        subtabs.addTab(chapters_tab, "فصل‌ها و پیشرفت")

        # تب فایل‌ها
        files_tab = QWidget()
        self.setup_files_tab(files_tab, thesis)
        subtabs.addTab(files_tab, "فایل‌ها")

        # تب یادداشت‌های خصوصی
        notes_tab = QWidget()
        self.setup_notes_tab(notes_tab, thesis)
        subtabs.addTab(notes_tab, "یادداشت‌ها")

        # تب پیام‌ها
        messages_tab = QWidget()
        self.setup_messages_tab(messages_tab, thesis)
        subtabs.addTab(messages_tab, "پیام‌ها")

        # تب تخمین نمره و دفاع
        grade_tab = QWidget()
        self.setup_grade_tab(grade_tab, thesis)
        subtabs.addTab(grade_tab, "نمره و دفاع")

        layout.addWidget(subtabs)
        self.detail_tab.setLayout(layout)

    def save_thesis_title(self, thesis):
        thesis.title = self.thesis_title_edit.text()
        self.get_session().commit()
        QMessageBox.information(self, "موفقیت", "عنوان ذخیره شد.")

    # جلسات
    def setup_meetings_tab(self, tab, thesis):
        layout = QVBoxLayout()
        self.meetings_table = QTableWidget()
        self.meetings_table.setColumnCount(5)
        self.meetings_table.setHorizontalHeaderLabels(["تاریخ", "موضوع", "یادداشت استاد", "کارها", "جلسه بعدی"])
        add_btn = QPushButton("افزودن جلسه")
        add_btn.clicked.connect(lambda: self.add_meeting(thesis))
        layout.addWidget(self.meetings_table)
        layout.addWidget(add_btn)
        tab.setLayout(layout)
        self.load_meetings(thesis)

    def load_meetings(self, thesis):
        meetings = self.get_session().query(Meeting).filter_by(thesis_id=thesis.id).order_by(Meeting.meeting_date).all()
        self.meetings_table.setRowCount(len(meetings))
        for i, m in enumerate(meetings):
            self.meetings_table.setItem(i, 0, QTableWidgetItem(str(m.meeting_date)))
            self.meetings_table.setItem(i, 1, QTableWidgetItem(m.topic or ""))
            self.meetings_table.setItem(i, 2, QTableWidgetItem(m.professor_notes or ""))
            self.meetings_table.setItem(i, 3, QTableWidgetItem(m.assigned_tasks or ""))
            self.meetings_table.setItem(i, 4, QTableWidgetItem(str(m.next_meeting_date) if m.next_meeting_date else ""))
        self.meetings_table.resizeColumnsToContents()

    def add_meeting(self, thesis):
        dialog = QDialog(self)
        dialog.setWindowTitle("جلسه جدید")
        layout = QFormLayout()
        date_edit = QDateEdit()
        date_edit.setDate(QDate.currentDate())
        topic_edit = QLineEdit()
        notes_edit = QTextEdit()
        tasks_edit = QTextEdit()
        next_date_edit = QDateEdit()
        next_date_edit.setDate(QDate.currentDate().addDays(7))
        layout.addRow("تاریخ:", date_edit)
        layout.addRow("موضوع:", topic_edit)
        layout.addRow("یادداشت:", notes_edit)
        layout.addRow("کارهای تعیین شده:", tasks_edit)
        layout.addRow("تاریخ جلسه بعدی:", next_date_edit)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addRow(button_box)
        dialog.setLayout(layout)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            meeting = Meeting(
                student_id=thesis.student_id,
                professor_id=thesis.supervisor_id,
                thesis_id=thesis.id,
                meeting_date=date_edit.date().toPyDate(),
                topic=topic_edit.text(),
                professor_notes=notes_edit.toPlainText(),
                assigned_tasks=tasks_edit.toPlainText(),
                next_meeting_date=next_date_edit.date().toPyDate() if next_date_edit.date() != QDate() else None
            )
            self.get_session().add(meeting)
            self.get_session().commit()
            self.load_meetings(thesis)

    # فصل‌ها و پیشرفت
    def setup_chapters_tab(self, tab, thesis):
        layout = QVBoxLayout()
        self.chapters_table = QTableWidget()
        self.chapters_table.setColumnCount(5)
        self.chapters_table.setHorizontalHeaderLabels(["شماره فصل", "عنوان", "تاریخ تحویل", "وضعیت", "عملیات"])
        add_chapter_btn = QPushButton("افزودن فصل (دانشجو)")
        layout.addWidget(self.chapters_table)
        layout.addWidget(add_chapter_btn)
        # درصد پیشرفت
        self.progress_bar = QProgressBar()
        layout.addWidget(QLabel("درصد پیشرفت:"))
        layout.addWidget(self.progress_bar)
        tab.setLayout(layout)
        self.load_chapters(thesis)

    def load_chapters(self, thesis):
        chapters = self.get_session().query(Chapter).filter_by(thesis_id=thesis.id).all()
        self.chapters_table.setRowCount(len(chapters))
        for i, ch in enumerate(chapters):
            self.chapters_table.setItem(i, 0, QTableWidgetItem(str(ch.chapter_number)))
            self.chapters_table.setItem(i, 1, QTableWidgetItem(ch.title))
            self.chapters_table.setItem(i, 2, QTableWidgetItem(str(ch.upload_date)))
            status = "تأیید شده" if ch.is_approved else "در انتظار"
            self.chapters_table.setItem(i, 3, QTableWidgetItem(status))
            btn = QPushButton("تأیید/رد")
            btn.clicked.connect(lambda _, c=ch: self.review_chapter(c))
            self.chapters_table.setCellWidget(i, 4, btn)
        # محاسبه درصد
        total = len(chapters)
        approved = len([ch for ch in chapters if ch.is_approved])
        percent = (approved / total * 100) if total > 0 else 0
        self.progress_bar.setValue(int(percent))
        thesis.progress_percent = percent
        self.get_session().commit()

    def review_chapter(self, chapter):
        dialog = QDialog(self)
        dialog.setWindowTitle("بازبینی فصل")
        layout = QVBoxLayout()
        lbl = QLabel(f"فصل {chapter.chapter_number}: {chapter.title}")
        layout.addWidget(lbl)
        reason_edit = QTextEdit()
        layout.addWidget(QLabel("دلیل (در صورت رد):"))
        layout.addWidget(reason_edit)
        btn_box = QDialogButtonBox()
        approve_btn = QPushButton("تأیید")
        reject_btn = QPushButton("رد")
        btn_box.addButton(approve_btn, QDialogButtonBox.ButtonRole.AcceptRole)
        btn_box.addButton(reject_btn, QDialogButtonBox.ButtonRole.RejectRole)
        layout.addWidget(btn_box)
        dialog.setLayout(layout)
        def on_approve():
            chapter.is_approved = True
            chapter.approval_date = datetime.now().date()
            chapter.revision_reason = ""
            self.get_session().commit()
            self.load_chapters(chapter.thesis)
            dialog.accept()
        def on_reject():
            chapter.is_approved = False
            chapter.revision_reason = reason_edit.toPlainText()
            self.get_session().commit()
            self.load_chapters(chapter.thesis)
            dialog.accept()
        approve_btn.clicked.connect(on_approve)
        reject_btn.clicked.connect(on_reject)
        dialog.exec()

    # فایل‌ها
    def setup_files_tab(self, tab, thesis):
        layout = QVBoxLayout()
        self.files_table = QTableWidget()
        self.files_table.setColumnCount(4)
        self.files_table.setHorizontalHeaderLabels(["نوع", "نام فایل", "تاریخ آپلود", "دانلود"])
        layout.addWidget(self.files_table)
        tab.setLayout(layout)
        self.load_files(thesis)

    def load_files(self, thesis):
        files = self.get_session().query(UploadedFile).filter_by(thesis_id=thesis.id).all()
        self.files_table.setRowCount(len(files))
        for i, f in enumerate(files):
            self.files_table.setItem(i, 0, QTableWidgetItem(f.file_type))
            self.files_table.setItem(i, 1, QTableWidgetItem(f.original_name))
            self.files_table.setItem(i, 2, QTableWidgetItem(str(f.upload_date)))
            btn = QPushButton("دانلود")
            btn.clicked.connect(lambda _, p=f.stored_path: QDesktopServices.openUrl(QUrl.fromLocalFile(p)))
            self.files_table.setCellWidget(i, 3, btn)

    # یادداشت‌های خصوصی
    def setup_notes_tab(self, tab, thesis):
        layout = QVBoxLayout()
        self.notes_edit = QTextEdit()
        layout.addWidget(QLabel("یادداشت‌های خصوصی استاد:"))
        layout.addWidget(self.notes_edit)
        save_btn = QPushButton("ذخیره")
        save_btn.clicked.connect(lambda: self.save_notes(thesis))
        layout.addWidget(save_btn)
        tab.setLayout(layout)

    def save_notes(self, thesis):
        # در حال حاضر یادداشت در thesis ذخیره نمی‌شود، می‌توان به thesis field اضافه کرد
        pass

    # پیام‌ها
    def setup_messages_tab(self, tab, thesis):
        layout = QVBoxLayout()
        self.messages_display = QTextEdit()
        self.messages_display.setReadOnly(True)
        self.message_input = QLineEdit()
        send_btn = QPushButton("ارسال")
        send_btn.clicked.connect(lambda: self.send_message(thesis))
        layout.addWidget(QLabel("گفتگو:"))
        layout.addWidget(self.messages_display)
        layout.addWidget(self.message_input)
        layout.addWidget(send_btn)
        tab.setLayout(layout)
        self.load_messages(thesis)

    def load_messages(self, thesis):
        student_user = self.get_session().query(User).filter_by(role='student', related_id=thesis.student_id).first()
        prof_user = self.main.current_user
        if not student_user or not prof_user:
            return
        msgs = self.get_session().query(Message).filter(
            ((Message.from_user_id == prof_user.id) & (Message.to_user_id == student_user.id)) |
            ((Message.from_user_id == student_user.id) & (Message.to_user_id == prof_user.id))
        ).order_by(Message.timestamp).all()
        text = ""
        for m in msgs:
            if m.from_user_id == prof_user.id:
                sender = "استاد"
            else:
                sender = "دانشجو"
            text += f"[{to_jalali_datetime(m.timestamp)}] {sender}: {m.message_text}\n"
        self.messages_display.setPlainText(text)

    def send_message(self, thesis):
        text = self.message_input.text().strip()
        if not text:
            return
        student_user = self.get_session().query(User).filter_by(role='student', related_id=thesis.student_id).first()
        msg = Message(from_user_id=self.main.current_user.id, to_user_id=student_user.id, message_text=text)
        self.get_session().add(msg)
        self.get_session().commit()
        self.message_input.clear()
        self.load_messages(thesis)

    # تخمین نمره و دفاع
    def setup_grade_tab(self, tab, thesis):
        layout = QVBoxLayout()
        # وزن‌ها
        w_group = QGroupBox("وزن‌های نمره تخمینی")
        w_layout = QFormLayout()
        self.w_progress = QDoubleSpinBox()
        self.w_progress.setRange(0, 100)
        self.w_progress.setValue(30)
        self.w_files = QDoubleSpinBox()
        self.w_files.setRange(0, 100)
        self.w_files.setValue(30)
        self.w_meetings = QDoubleSpinBox()
        self.w_meetings.setRange(0, 100)
        self.w_meetings.setValue(40)
        w_layout.addRow("پیشرفت (درصد):", self.w_progress)
        w_layout.addRow("کیفیت فایل‌ها:", self.w_files)
        w_layout.addRow("عملکرد جلسات:", self.w_meetings)
        w_group.setLayout(w_layout)
        layout.addWidget(w_group)
        self.grade_label = QLabel("نمره تخمینی: 0")
        layout.addWidget(self.grade_label)
        calc_btn = QPushButton("محاسبه نمره")
        calc_btn.clicked.connect(lambda: self.calculate_estimated_grade(thesis))
        layout.addWidget(calc_btn)

        # دفاع
        defense_btn = QPushButton("شروع ارزیابی دفاع")
        defense_btn.clicked.connect(lambda: self.open_defense_evaluation(thesis))
        layout.addWidget(defense_btn)
        tab.setLayout(layout)

    def calculate_estimated_grade(self, thesis):
        # محاسبه بر اساس وزن‌ها
        progress = thesis.progress_percent / 100  # 0-1
        # کیفیت فایل‌ها: می‌توان تعداد فایل‌های آپلود شده
        files_score = min(1.0, len(thesis.files) / 5)
        # جلسات: تعداد جلسات
        meetings_score = min(1.0, len(thesis.meetings) / 3)
        total = (progress * self.w_progress.value() + files_score * self.w_files.value() + meetings_score * self.w_meetings.value()) / 100
        grade = total * 20
        self.grade_label.setText(f"نمره تخمینی: {grade:.2f}")

    def open_defense_evaluation(self, thesis):
        dialog = DefenseEvaluationDialog(self, thesis)
        dialog.exec()
        # بعد از ذخیره نهایی، thesis.final_grade آپدیت می‌شود

    # --------------------- درخواست‌های راهنمایی ---------------------
    def setup_requests(self):
        layout = QVBoxLayout()
        self.requests_table = QTableWidget()
        self.requests_table.setColumnCount(6)
        self.requests_table.setHorizontalHeaderLabels(["دانشجو", "کد ملی", "رشته", "تاریخ درخواست", "پیام", "عملیات"])
        layout.addWidget(self.requests_table)
        self.requests_tab.setLayout(layout)

    def load_requests(self):
        prof = self.current_professor
        requests = self.get_session().query(Request).filter_by(professor_id=prof.id, status='pending').all()
        self.requests_table.setRowCount(len(requests))
        for i, req in enumerate(requests):
            student = self.get_session().query(Student).filter_by(id=req.student_id).first()
            self.requests_table.setItem(i, 0, QTableWidgetItem(f"{student.first_name} {student.last_name}"))
            self.requests_table.setItem(i, 1, QTableWidgetItem(student.national_id))
            self.requests_table.setItem(i, 2, QTableWidgetItem(student.field_of_study))
            self.requests_table.setItem(i, 3, QTableWidgetItem(str(req.request_date)))
            self.requests_table.setItem(i, 4, QTableWidgetItem(req.student_message or ""))
            accept_btn = QPushButton("قبول")
            reject_btn = QPushButton("رد")
            accept_btn.clicked.connect(lambda _, r=req: self.accept_request(r))
            reject_btn.clicked.connect(lambda _, r=req: self.reject_request(r))
            btn_widget = QWidget()
            btn_layout = QHBoxLayout()
            btn_layout.addWidget(accept_btn)
            btn_layout.addWidget(reject_btn)
            btn_widget.setLayout(btn_layout)
            self.requests_table.setCellWidget(i, 5, btn_widget)

    def accept_request(self, req):
        prof = self.current_professor
        if len(prof.theses_supervised) >= prof.max_students:
            QMessageBox.warning(self, "خطا", "ظرفیت شما تکمیل است.")
            return
        req.status = 'accepted'
        # ایجاد thesis
        student = self.get_session().query(Student).filter_by(id=req.student_id).first()
        thesis = Thesis(student_id=student.id, supervisor_id=prof.id, title="", project_status="در حال انجام")
        self.get_session().add(thesis)
        self.get_session().commit()
        self.load_requests()
        QMessageBox.information(self, "موفقیت", "درخواست پذیرفته شد و پایان‌نامه ایجاد گردید.")

    def reject_request(self, req):
        req.status = 'rejected'
        self.get_session().commit()
        self.load_requests()

    # --------------------- گزارش‌های استاد ---------------------
    def setup_reports(self):
        layout = QVBoxLayout()
        self.prof_report_widget = pg.GraphicsLayoutWidget()
        layout.addWidget(self.prof_report_widget)
        export_btn = QPushButton("خروجی اکسل")
        export_btn.clicked.connect(self.export_students_excel)
        layout.addWidget(export_btn)
        self.reports_tab.setLayout(layout)

    def export_students_excel(self):
        prof = self.current_professor
        data = []
        for thesis in prof.theses_supervised:
            data.append({
                "نام": thesis.student.first_name,
                "نام خانوادگی": thesis.student.last_name,
                "شماره دانشجویی": thesis.student.student_id,
                "عنوان": thesis.title,
                "پیشرفت": thesis.progress_percent,
                "وضعیت": thesis.project_status
            })
        df = pd.DataFrame(data)
        path, _ = QFileDialog.getSaveFileName(self, "ذخیره گزارش", "", "Excel files (*.xlsx)")
        if path:
            df.to_excel(path, index=False)
            QMessageBox.information(self, "موفقیت", "گزارش ذخیره شد.")

# ----------------------------------------------------------------------
# دیالوگ ارزیابی دفاع
# ----------------------------------------------------------------------
class DefenseEvaluationDialog(QDialog):
    def __init__(self, parent, thesis):
        super().__init__(parent)
        self.thesis = thesis
        self.setWindowTitle("ارزیابی دفاع")
        self.setMinimumSize(600, 500)
        self.categories = {}
        layout = QVBoxLayout()
        # دسته‌ها
        cat_names = ["ارائه و ساختار", "روش‌شناسی تحقیق", "نتایج و تحلیل", "پاسخ به سوالات", "نوآوری"]
        for cat in cat_names:
            group = QGroupBox(cat)
            cat_layout = QVBoxLayout()
            # هر دسته چند زیرسوال با دکمه رادیویی
            scores = {"ضعیف": 1, "متوسط": 2, "خوب": 3, "عالی": 4}
            self.categories[cat] = {}
            for sub, label in enumerate(["سوال ۱", "سوال ۲"], start=1):
                sub_group = QHBoxLayout()
                sub_group.addWidget(QLabel(label))
                radio_group = QButtonGroup(self)
                for text, val in scores.items():
                    radio = QRadioButton(text)
                    radio.setProperty("score", val)
                    radio_group.addButton(radio)
                    sub_group.addWidget(radio)
                # انتخاب پیش‌فرض متوسط
                radio_group.buttons()[1].setChecked(True)
                self.categories[cat][f"sub_{sub}"] = radio_group
                cat_layout.addLayout(sub_group)
            group.setLayout(cat_layout)
            layout.addWidget(group)

        # دکمه‌ها
        btn_layout = QHBoxLayout()
        save_draft_btn = QPushButton("ذخیره پیش‌نویس")
        save_draft_btn.clicked.connect(lambda: self.save_evaluation(False))
        submit_btn = QPushButton("ارسال نهایی")
        submit_btn.clicked.connect(lambda: self.save_evaluation(True))
        btn_layout.addWidget(save_draft_btn)
        btn_layout.addWidget(submit_btn)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def save_evaluation(self, final):
        # جمع‌آوری امتیازات
        total_score = 0
        count = 0
        eval_data = {}
        for cat, subs in self.categories.items():
            eval_data[cat] = {}
            for sub, radio_group in subs.items():
                checked = radio_group.checkedButton()
                if checked:
                    score = checked.property("score")
                    eval_data[cat][sub] = score
                    total_score += score
                    count += 1
        avg = total_score / count if count > 0 else 0
        # تبدیل به نمره 0-20
        final_grade = round(avg * 5, 2)  # حداکثر 4*5=20

        # ذخیره در دیتابیس
        session = get_session()
        existing = session.query(DefenseEvaluation).filter_by(thesis_id=self.thesis.id).first()
        if not existing:
            existing = DefenseEvaluation(thesis_id=self.thesis.id, professor_id=self.thesis.supervisor_id)
            session.add(existing)
        existing.evaluation_data_json = eval_data
        existing.final_score = final_grade
        existing.is_final = final
        if final:
            self.thesis.final_grade = final_grade
            self.thesis.project_status = "دفاع شده"
        session.commit()
        session.close()
        QMessageBox.information(self, "موفقیت", "ارزیابی ذخیره شد." if not final else "نمره نهایی ثبت و دفاع پایان یافت.")
        if final:
            self.accept()
        else:
            self.accept()

# ----------------------------------------------------------------------
# پنل دانشجو
# ----------------------------------------------------------------------
class StudentPanel(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main = main_window
        self.session = None
        self.current_student = None
        layout = QVBoxLayout()
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)
        self.setLayout(layout)

        self.profile_tab = QWidget()
        self.thesis_tab = QWidget()
        self.meetings_tab = QWidget()
        self.messages_tab = QWidget()
        self.progress_tab = QWidget()
        self.defense_tab = QWidget()
        self.request_guidance_tab = QWidget()

        self.tabs.addTab(self.profile_tab, "پروفایل")
        self.tabs.addTab(self.thesis_tab, "پایان‌نامه")
        self.tabs.addTab(self.meetings_tab, "جلسات")
        self.tabs.addTab(self.messages_tab, "پیام‌ها")
        self.tabs.addTab(self.progress_tab, "پیشرفت")
        self.tabs.addTab(self.defense_tab, "ارزیابی دفاع")
        self.tabs.addTab(self.request_guidance_tab, "درخواست راهنمایی")

        self.setup_profile()
        self.setup_thesis()
        self.setup_meetings()
        self.setup_messages()
        self.setup_progress()
        self.setup_defense()
        self.setup_request_guidance()

    def refresh(self):
        user = self.main.current_user
        if user.role != 'student':
            return
        self.session = get_session()
        self.current_student = self.session.query(Student).filter_by(id=user.related_id).first()
        if not self.current_student:
            QMessageBox.critical(self, "خطا", "اطلاعات دانشجو یافت نشد.")
            self.main.logout()
            return
        self.load_profile()
        self.load_thesis()
        self.load_meetings()
        self.load_messages()
        self.load_progress()
        self.load_defense()
        self.load_requests()

    # پروفایل
    def setup_profile(self):
        layout = QFormLayout()
        self.profile_info = {}
        for label, key in [("نام", "first_name"), ("نام خانوادگی", "last_name"), ("کد ملی", "national_id"),
                           ("شماره دانشجویی", "student_id"), ("رشته", "field_of_study"), ("مقطع", "grade"),
                           ("سال ورود", "entry_year"), ("ایمیل", "email"), ("شماره تماس", "phone")]:
            lbl = QLabel()
            self.profile_info[key] = lbl
            layout.addRow(label + ":", lbl)
        self.profile_photo = QLabel()
        layout.addRow("عکس:", self.profile_photo)
        edit_btn = QPushButton("ویرایش (ایمیل/تلفن)")
        edit_btn.clicked.connect(self.edit_profile)
        change_sup_btn = QPushButton("درخواست تغییر استاد راهنما")
        change_sup_btn.clicked.connect(self.request_change_supervisor)
        layout.addRow(edit_btn)
        layout.addRow(change_sup_btn)
        self.profile_tab.setLayout(layout)

    def load_profile(self):
        stu = self.current_student
        for key, lbl in self.profile_info.items():
            lbl.setText(str(getattr(stu, key, "")))
        if stu.photo_path and os.path.exists(stu.photo_path):
            self.profile_photo.setPixmap(QPixmap(stu.photo_path).scaled(100, 100, Qt.AspectRatioMode.KeepAspectRatio))

    def edit_profile(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("ویرایش اطلاعات")
        layout = QFormLayout()
        email_edit = QLineEdit(self.current_student.email)
        phone_edit = QLineEdit(self.current_student.phone)
        layout.addRow("ایمیل:", email_edit)
        layout.addRow("تلفن:", phone_edit)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addRow(button_box)
        dialog.setLayout(layout)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.current_student.email = email_edit.text()
            self.current_student.phone = phone_edit.text()
            self.session.commit()
            self.load_profile()

    def request_change_supervisor(self):
        # درخواست به ادمین (می‌توان یک جدول درخواست‌های خاص داشت)
        QMessageBox.information(self, "توجه", "درخواست شما ثبت شد. لطفاً با مسئول آموزش تماس بگیرید.")

    # پایان‌نامه
    def setup_thesis(self):
        layout = QFormLayout()
        self.thesis_title_label = QLabel()
        self.thesis_status_label = QLabel()
        layout.addRow("عنوان:", self.thesis_title_label)
        layout.addRow("وضعیت:", self.thesis_status_label)
        # ویرایش عنوان (قبل از تأیید)
        self.title_edit = QLineEdit()
        self.title_edit.setPlaceholderText("عنوان جدید")
        save_title_btn = QPushButton("ذخیره عنوان")
        save_title_btn.clicked.connect(self.save_title)
        layout.addRow(self.title_edit, save_title_btn)
        # آپلود فایل
        file_type_combo = QComboBox()
        file_type_combo.addItems(["proposal", "chapter", "article", "final_thesis", "defense_presentation"])
        self.file_type_combo = file_type_combo
        upload_btn = QPushButton("آپلود فایل")
        upload_btn.clicked.connect(self.upload_file)
        layout.addRow("نوع فایل:", file_type_combo)
        layout.addRow(upload_btn)
        # لیست فایل‌ها
        self.student_files_table = QTableWidget()
        self.student_files_table.setColumnCount(3)
        self.student_files_table.setHorizontalHeaderLabels(["نوع", "نام", "تاریخ"])
        layout.addRow(self.student_files_table)
        self.thesis_tab.setLayout(layout)

    def load_thesis(self):
        stu = self.current_student
        thesis = stu.theses if hasattr(stu, 'theses') else None
        if thesis:
            self.thesis_title_label.setText(thesis.title or "بدون عنوان")
            self.thesis_status_label.setText(thesis.project_status)
        else:
            self.thesis_title_label.setText("ندارد")
            self.thesis_status_label.setText("-")
        self.load_student_files()

    def save_title(self):
        thesis = self.current_student.theses
        if thesis and not thesis.supervisor:  # قبل از تأیید استاد
            thesis.title = self.title_edit.text()
            self.session.commit()
            self.load_thesis()
        else:
            QMessageBox.warning(self, "خطا", "امکان ویرایش عنوان وجود ندارد.")

    def upload_file(self):
        thesis = self.current_student.theses
        if not thesis:
            QMessageBox.warning(self, "خطا", "هنوز پایان‌نامه‌ای تعریف نشده.")
            return
        file_path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل")
        if file_path:
            dest = UPLOAD_DIR / os.path.basename(file_path)
            shutil.copy2(file_path, dest)
            ftype = self.file_type_combo.currentText()
            file_record = UploadedFile(
                thesis_id=thesis.id,
                file_type=ftype,
                original_name=os.path.basename(file_path),
                stored_path=str(dest),
                version=1
            )
            self.session.add(file_record)
            self.session.commit()
            self.load_student_files()

    def load_student_files(self):
        thesis = self.current_student.theses
        if thesis:
            files = self.session.query(UploadedFile).filter_by(thesis_id=thesis.id).all()
            self.student_files_table.setRowCount(len(files))
            for i, f in enumerate(files):
                self.student_files_table.setItem(i, 0, QTableWidgetItem(f.file_type))
                self.student_files_table.setItem(i, 1, QTableWidgetItem(f.original_name))
                self.student_files_table.setItem(i, 2, QTableWidgetItem(str(f.upload_date)))
        else:
            self.student_files_table.setRowCount(0)

    # جلسات
    def setup_meetings(self):
        layout = QVBoxLayout()
        self.meetings_table_student = QTableWidget()
        self.meetings_table_student.setColumnCount(4)
        self.meetings_table_student.setHorizontalHeaderLabels(["تاریخ", "موضوع", "یادداشت استاد", "وظایف"])
        request_btn = QPushButton("درخواست جلسه جدید")
        request_btn.clicked.connect(self.request_meeting)
        layout.addWidget(self.meetings_table_student)
        layout.addWidget(request_btn)
        self.meetings_tab.setLayout(layout)

    def load_meetings(self):
        stu = self.current_student
        thesis = stu.theses
        if thesis:
            meetings = self.session.query(Meeting).filter_by(thesis_id=thesis.id).all()
            self.meetings_table_student.setRowCount(len(meetings))
            for i, m in enumerate(meetings):
                self.meetings_table_student.setItem(i, 0, QTableWidgetItem(str(m.meeting_date)))
                self.meetings_table_student.setItem(i, 1, QTableWidgetItem(m.topic or ""))
                self.meetings_table_student.setItem(i, 2, QTableWidgetItem(m.professor_notes or ""))
                self.meetings_table_student.setItem(i, 3, QTableWidgetItem(m.assigned_tasks or ""))
        else:
            self.meetings_table_student.setRowCount(0)

    def request_meeting(self):
        # اینجا می‌توان پیامی به استاد فرستاد
        QMessageBox.information(self, "توجه", "درخواست شما ثبت شد. استاد راهنما مطلع می‌شود.")

    # پیام‌ها
    def setup_messages(self):
        layout = QVBoxLayout()
        self.student_msgs_display = QTextEdit()
        self.student_msgs_display.setReadOnly(True)
        self.student_msg_input = QLineEdit()
        send_btn = QPushButton("ارسال")
        send_btn.clicked.connect(self.send_message_student)
        layout.addWidget(QLabel("پیام‌ها:"))
        layout.addWidget(self.student_msgs_display)
        layout.addWidget(self.student_msg_input)
        layout.addWidget(send_btn)
        self.messages_tab.setLayout(layout)

    def load_messages(self):
        stu = self.current_student
        thesis = stu.theses
        if not thesis or not thesis.supervisor:
            self.student_msgs_display.setPlainText("استاد راهنما مشخص نشده.")
            return
        prof_user = self.session.query(User).filter_by(role='professor', related_id=thesis.supervisor_id).first()
        student_user = self.main.current_user
        if not prof_user:
            return
        msgs = self.session.query(Message).filter(
            ((Message.from_user_id == student_user.id) & (Message.to_user_id == prof_user.id)) |
            ((Message.from_user_id == prof_user.id) & (Message.to_user_id == student_user.id))
        ).order_by(Message.timestamp).all()
        text = ""
        for m in msgs:
            if m.from_user_id == student_user.id:
                sender = "شما"
            else:
                sender = "استاد"
            text += f"[{to_jalali_datetime(m.timestamp)}] {sender}: {m.message_text}\n"
        self.student_msgs_display.setPlainText(text)

    def send_message_student(self):
        text = self.student_msg_input.text().strip()
        if not text:
            return
        stu = self.current_student
        thesis = stu.theses
        if not thesis or not thesis.supervisor:
            return
        prof_user = self.session.query(User).filter_by(role='professor', related_id=thesis.supervisor_id).first()
        msg = Message(from_user_id=self.main.current_user.id, to_user_id=prof_user.id, message_text=text)
        self.session.add(msg)
        self.session.commit()
        self.student_msg_input.clear()
        self.load_messages()

    # پیشرفت
    def setup_progress(self):
        layout = QVBoxLayout()
        self.progress_bar_student = QProgressBar()
        layout.addWidget(QLabel("درصد پیشرفت:"))
        layout.addWidget(self.progress_bar_student)
        # نمودار
        self.progress_plot = pg.PlotWidget()
        layout.addWidget(self.progress_plot)
        self.progress_tab.setLayout(layout)

    def load_progress(self):
        stu = self.current_student
        thesis = stu.theses
        if thesis:
            self.progress_bar_student.setValue(int(thesis.progress_percent))
            # نمودار ساده
            self.progress_plot.clear()
            self.progress_plot.plot([1,2,3], [thesis.progress_percent]*3, pen='b')

    # ارزیابی دفاع
    def setup_defense(self):
        layout = QVBoxLayout()
        self.defense_grade_label = QLabel("نمره دفاع: نامشخص")
        layout.addWidget(self.defense_grade_label)
        self.defense_tab.setLayout(layout)

    def load_defense(self):
        stu = self.current_student
        thesis = stu.theses
        if thesis and thesis.defense_evaluation and thesis.defense_evaluation.is_final:
            self.defense_grade_label.setText(f"نمره نهایی: {thesis.defense_evaluation.final_score}")
        else:
            self.defense_grade_label.setText("هنوز نمره‌ای ثبت نشده.")

    # درخواست راهنمایی
    def setup_request_guidance(self):
        layout = QVBoxLayout()
        self.prof_list_table = QTableWidget()
        self.prof_list_table.setColumnCount(3)
        self.prof_list_table.setHorizontalHeaderLabels(["نام استاد", "تخصص", "ظرفیت باقی‌مانده"])
        select_btn = QPushButton("درخواست راهنمایی")
        select_btn.clicked.connect(self.send_guidance_request)
        layout.addWidget(QLabel("لیست اساتید قابل انتخاب:"))
        layout.addWidget(self.prof_list_table)
        layout.addWidget(select_btn)
        self.request_guidance_tab.setLayout(layout)

    def load_requests(self):
        # اگر دانشجو هنوز استاد ندارد، لیست اساتید را نمایش دهد
        thesis = self.current_student.theses
        if thesis and thesis.supervisor:
            self.prof_list_table.setRowCount(0)
            return
        professors = self.session.query(Professor).all()
        self.prof_list_table.setRowCount(len(professors))
        for i, prof in enumerate(professors):
            self.prof_list_table.setItem(i, 0, QTableWidgetItem(f"{prof.first_name} {prof.last_name}"))
            self.prof_list_table.setItem(i, 1, QTableWidgetItem(prof.expertise))
            remaining = prof.max_students - len(prof.theses_supervised)
            self.prof_list_table.setItem(i, 2, QTableWidgetItem(str(remaining)))

    def send_guidance_request(self):
        row = self.prof_list_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "خطا", "استادی انتخاب نشده.")
            return
        # دریافت استاد
        professors = self.session.query(Professor).all()
        selected_prof = professors[row]
        # بررسی وجود درخواست قبلی
        existing = self.session.query(Request).filter_by(student_id=self.current_student.id, professor_id=selected_prof.id, status='pending').first()
        if existing:
            QMessageBox.warning(self, "خطا", "درخواست قبلاً ثبت شده.")
            return
        # دریافت پیام (اختیاری)
        text, ok = QInputDialog.getText(self, "پیام", "متن درخواست (اختیاری):")
        if ok:
            req = Request(student_id=self.current_student.id, professor_id=selected_prof.id, student_message=text)
            self.session.add(req)
            self.session.commit()
            QMessageBox.information(self, "موفقیت", "درخواست شما برای استاد ارسال شد.")

# ----------------------------------------------------------------------
# تابع اصلی
# ----------------------------------------------------------------------
def create_default_admin():
    session = get_session()
    admin_user = session.query(User).filter_by(username='admin', role='admin').first()
    if not admin_user:
        logger.info("ایجاد ادمین پیش‌فرض")
        admin = User(
            username='admin',
            password_hash=hash_password('admin'),
            role='admin',
            is_active=True
        )
        session.add(admin)
        session.commit()
    session.close()

if __name__ == '__main__':
    # اولین اجرا: ساخت دیتابیس و ادمین
    create_default_admin()
    app = QApplication(sys.argv)
    # تنظیم فونت پیش‌فرض
    font = QFont("Vazirmatn", 10)
    app.setFont(font)
    # استایل اولیه تاریک
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec())