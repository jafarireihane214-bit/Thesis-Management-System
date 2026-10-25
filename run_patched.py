# -*- coding: utf-8 -*-
"""
فایل اجرای اصلاح‌شده با رفع مشکلات بایگانی ادمین، نمودار دانشجو و تنظیم دستی پیشرفت
اجرا: python run_patched.py
"""
import sys
import os
import importlib.util
import shutil
import json
import re  # برای پاک‌سازی نام پوشه

# ======================== ایمپورت‌های PyQt6 =========================
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QTableWidget, QTableWidgetItem,
    QStackedWidget, QMessageBox, QFileDialog, QCheckBox, QTextEdit,
    QTabWidget, QGroupBox, QGridLayout, QProgressBar,
    QStatusBar, QToolBar, QDialog, QDialogButtonBox, QFormLayout,
    QSpinBox, QDoubleSpinBox, QDateEdit, QInputDialog, QRadioButton, QButtonGroup
)
from PyQt6.QtGui import QPixmap, QFont, QColor
from PyQt6.QtCore import Qt, QDate, QUrl
from PyQt6.QtGui import QDesktopServices

# ======================== بارگذاری پویای main.py =========================
file_path = os.path.join(os.path.dirname(__file__), "main.py")
spec = importlib.util.spec_from_file_location("thesis_module", file_path)
main_module = importlib.util.module_from_spec(spec)
sys.modules["thesis_module"] = main_module
spec.loader.exec_module(main_module)

# دریافت کلاس‌ها و توابع مورد نیاز
ProfessorPanel = main_module.ProfessorPanel
StudentPanel = main_module.StudentPanel
AdminPanel = main_module.AdminPanel
DefenseEvaluationDialog = main_module.DefenseEvaluationDialog
get_session = main_module.get_session
Chapter = main_module.Chapter
MainWindow = main_module.MainWindow
create_default_admin = main_module.create_default_admin
pg = main_module.pg
UploadedFile = main_module.UploadedFile
Thesis = main_module.Thesis
Student = main_module.Student
ARCHIVE_DIR = main_module.ARCHIVE_DIR
UPLOAD_DIR = main_module.UPLOAD_DIR
# اضافه کردن or_ از sqlalchemy
from sqlalchemy import or_

# ============================================================
# ۱. متدهای اصلاح‌شده برای استاد (دسترسی دستی پیشرفت)
# ============================================================
def save_manual_progress(self, thesis):
    new_progress = self.progress_spinbox.value()
    thesis.progress_percent = new_progress
    get_session().commit()
    self.progress_bar.setValue(int(new_progress))
    QMessageBox.information(self, "موفقیت", "درصد پیشرفت به‌روز شد.")

def patched_setup_chapters_tab(self, tab, thesis):
    layout = QVBoxLayout()
    self.chapters_table = QTableWidget()
    self.chapters_table.setColumnCount(5)
    self.chapters_table.setHorizontalHeaderLabels(
        ["شماره فصل", "عنوان", "تاریخ تحویل", "وضعیت", "عملیات"]
    )
    layout.addWidget(self.chapters_table)

    progress_layout = QHBoxLayout()
    progress_layout.addWidget(QLabel("درصد پیشرفت (دستی):"))
    self.progress_spinbox = QDoubleSpinBox()
    self.progress_spinbox.setRange(0, 100)
    self.progress_spinbox.setSingleStep(1)
    self.progress_spinbox.setSuffix(" %")
    progress_layout.addWidget(self.progress_spinbox)

    save_progress_btn = QPushButton("ذخیره پیشرفت")
    save_progress_btn.clicked.connect(lambda: self.save_manual_progress(thesis))
    progress_layout.addWidget(save_progress_btn)
    layout.addLayout(progress_layout)

    self.progress_bar = QProgressBar()
    layout.addWidget(QLabel("پیشرفت فعلی:"))
    layout.addWidget(self.progress_bar)

    add_chapter_btn = QPushButton("افزودن فصل (دانشجو)")
    layout.addWidget(add_chapter_btn)

    tab.setLayout(layout)
    self.load_chapters(thesis)

def patched_load_chapters(self, thesis):
    session = get_session()
    chapters = session.query(Chapter).filter_by(thesis_id=thesis.id).all()
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

    current_progress = thesis.progress_percent or 0.0
    self.progress_bar.setValue(int(current_progress))
    self.progress_spinbox.setValue(current_progress)

def patched_show_student_detail(self, thesis):
    if self.tabs.count() == 4:
        self.tabs.addTab(self.detail_tab, "جزئیات دانشجو")
    self.tabs.setCurrentWidget(self.detail_tab)

    if self.detail_tab.layout():
        while self.detail_tab.layout().count():
            child = self.detail_tab.layout().takeAt(0)
            if child.widget():
                child.widget().deleteLater()

    layout = QVBoxLayout()
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

    thesis_group = QGroupBox("پایان‌نامه")
    t_layout = QFormLayout()
    self.thesis_title_edit = QLineEdit(thesis.title or "")
    t_layout.addRow("عنوان:", self.thesis_title_edit)
    save_title_btn = QPushButton("ذخیره عنوان")
    save_title_btn.clicked.connect(lambda: self.save_thesis_title(thesis))
    t_layout.addRow(save_title_btn)
    self.proposal_date_edit = QDateEdit()
    if thesis.proposal_approval_date:
        self.proposal_date_edit.setDate(QDate(thesis.proposal_approval_date))
    t_layout.addRow("تاریخ تصویب پروپوزال:", self.proposal_date_edit)
    thesis_group.setLayout(t_layout)
    layout.addWidget(thesis_group)

    subtabs = QTabWidget()
    self.detail_subtabs = subtabs

    meetings_tab = QWidget()
    self.setup_meetings_tab(meetings_tab, thesis)
    subtabs.addTab(meetings_tab, "جلسات")

    chapters_tab = QWidget()
    self.setup_chapters_tab(chapters_tab, thesis)
    subtabs.addTab(chapters_tab, "فصل‌ها و پیشرفت")

    files_tab = QWidget()
    self.setup_files_tab(files_tab, thesis)
    subtabs.addTab(files_tab, "فایل‌ها")

    notes_tab = QWidget()
    self.setup_notes_tab(notes_tab, thesis)
    subtabs.addTab(notes_tab, "یادداشت‌ها")

    messages_tab = QWidget()
    self.setup_messages_tab(messages_tab, thesis)
    subtabs.addTab(messages_tab, "پیام‌ها")

    grade_tab = QWidget()
    self.setup_grade_tab(grade_tab, thesis)
    subtabs.addTab(grade_tab, "نمره و دفاع")

    layout.addWidget(subtabs)
    self.detail_tab.setLayout(layout)

# ============================================================
# ۲. متدهای اصلاح‌شده برای دانشجو (نمودار زیبا)
# ============================================================
def patched_load_progress(self):
    stu = self.current_student
    thesis = stu.theses
    if thesis:
        progress_value = thesis.progress_percent if thesis.progress_percent else 0.0
        self.progress_bar_student.setValue(int(progress_value))

        self.progress_plot.clear()

        steps = 5
        x = list(range(steps))
        y = [ (i / (steps-1)) * progress_value for i in range(steps) ]

        pen = pg.mkPen(color=QColor(0, 150, 200), width=4)
        line = self.progress_plot.plot(x, y, pen=pen, symbol='o', symbolSize=12,
                                       symbolBrush=QColor(0, 200, 150), symbolPen='w')

        fill_brush = pg.mkBrush(color=QColor(0, 150, 200, 50))
        self.progress_plot.addItem(pg.FillBetweenItem(line, pg.PlotDataItem(x, [0]*len(x)), brush=fill_brush))

        target_pen = pg.mkPen(color=QColor(200, 50, 50), width=2, style=Qt.PenStyle.DashLine)
        self.progress_plot.plot(x, [100]*len(x), pen=target_pen)

        self.progress_plot.setLabel('left', 'درصد پیشرفت', color='#555', size='12pt')
        self.progress_plot.setLabel('bottom', 'مراحل', color='#555', size='12pt')
        self.progress_plot.setTitle('روند پیشرفت پایان‌نامه', color='#333', size='14pt')
        self.progress_plot.showGrid(x=True, y=True, alpha=0.3)
        self.progress_plot.setYRange(0, 110)
        self.progress_plot.setXRange(-0.5, steps-0.5)
        if self.main.current_theme == 'dark':
            self.progress_plot.setBackground('#2d2d44')
        else:
            self.progress_plot.setBackground('white')
    else:
        self.progress_bar_student.setValue(0)
        self.progress_plot.clear()
        self.progress_plot.setTitle('هیچ پایان‌نامه‌ای تعریف نشده')

# ============================================================
# ۳. متدهای اصلاح‌شده برای ادمین (بایگانی کامل)
# ============================================================

# ۳-۱) اصلاح تابع بارگذاری بایگانی (نمایش دفاع‌شده‌ها)
def patched_load_archive(self):
    session = get_session()
    # نمایش پایان‌نامه‌هایی که بایگانی شده‌اند یا وضعیت دفاع شده دارند
    query = session.query(Thesis).filter(
        or_(Thesis.is_archived == True, Thesis.project_status == "دفاع شده")
    )
    search = self.archive_search.text().strip()
    if search:
        query = query.join(Student).filter(
            (Student.first_name.contains(search)) | (Student.last_name.contains(search)) |
            (Student.student_id.contains(search)) | (Thesis.title.contains(search))
        )
    year = self.archive_year_combo.currentText()
    if year != "همه سال‌ها":
        try:
            y = int(year)
            query = query.filter(Thesis.defense_date != None).filter(
                Thesis.defense_date >= QDate(y, 1, 1).toPyDate(),
                Thesis.defense_date <= QDate(y, 12, 31).toPyDate()
            )
        except:
            pass

    theses = query.all()
    self.archive_table.setRowCount(len(theses))
    self.archive_table.setColumnCount(7)
    self.archive_table.setHorizontalHeaderLabels(
        ["دانشجو", "شماره دانشجویی", "عنوان", "تاریخ دفاع", "نمره", "تعداد فایل‌ها", "عملیات"]
    )
    for i, thesis in enumerate(theses):
        student = thesis.student
        self.archive_table.setItem(i, 0, QTableWidgetItem(f"{student.first_name} {student.last_name}"))
        self.archive_table.setItem(i, 1, QTableWidgetItem(student.student_id))
        self.archive_table.setItem(i, 2, QTableWidgetItem(thesis.title or ""))
        self.archive_table.setItem(i, 3, QTableWidgetItem(str(thesis.defense_date) if thesis.defense_date else ""))
        self.archive_table.setItem(i, 4, QTableWidgetItem(str(thesis.final_grade) if thesis.final_grade else "ثبت نشده"))
        file_count = session.query(UploadedFile).filter_by(thesis_id=thesis.id).count()
        self.archive_table.setItem(i, 5, QTableWidgetItem(str(file_count)))
        btn = QPushButton("دانلود بایگانی")
        btn.clicked.connect(lambda _, t=thesis: self.download_archive_item(t))
        self.archive_table.setCellWidget(i, 6, btn)
    self.archive_table.resizeColumnsToContents()

# ۳-۲) متد دانلود بایگانی (بدون تغییر)
def patched_download_archive_item(self, thesis):
    student = thesis.student
    folder_name = f"{student.student_id}_{thesis.title or 'پایان‌نامه'}"
    folder_name = re.sub(r'[\\/*?:"<>|]', "", folder_name)
    dest_dir = ARCHIVE_DIR / folder_name
    dest_dir.mkdir(exist_ok=True)

    files = get_session().query(UploadedFile).filter_by(thesis_id=thesis.id).all()
    if not files:
        QMessageBox.warning(self, "توجه", "هیچ فایلی برای این پایان‌نامه وجود ندارد.")
        return

    copied = 0
    for file in files:
        if os.path.exists(file.stored_path):
            shutil.copy(file.stored_path, dest_dir / file.original_name)
            copied += 1

    meta = {
        "student": {
            "first_name": student.first_name,
            "last_name": student.last_name,
            "student_id": student.student_id,
            "field": student.field_of_study
        },
        "supervisor": f"{thesis.supervisor.first_name} {thesis.supervisor.last_name}" if thesis.supervisor else "",
        "defense_date": str(thesis.defense_date) if thesis.defense_date else "",
        "final_grade": thesis.final_grade,
        "meetings": [{"date": str(m.meeting_date), "topic": m.topic} for m in thesis.meetings],
        "files": [{"name": f.original_name, "type": f.file_type} for f in files]
    }
    with open(dest_dir / "metadata.json", "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

    QMessageBox.information(self, "موفقیت", f"بایگانی در {dest_dir} ایجاد شد.\nتعداد فایل‌های کپی‌شده: {copied}")

# ۳-۳) متد جدید برای بایگانی دستی
def patched_archive_selected_thesis(self):
    row = self.archive_table.currentRow()
    if row < 0:
        QMessageBox.warning(self, "خطا", "لطفاً یک پایان‌نامه انتخاب کنید.")
        return
    # دریافت thesis از ردیف انتخاب شده
    title = self.archive_table.item(row, 2).text()
    session = get_session()
    thesis = session.query(Thesis).filter_by(title=title).first()
    if not thesis:
        QMessageBox.warning(self, "خطا", "پایان‌نامه یافت نشد.")
        return
    if thesis.is_archived:
        QMessageBox.information(self, "توجه", "این پایان‌نامه قبلاً بایگانی شده است.")
        return
    thesis.is_archived = True
    session.commit()
    QMessageBox.information(self, "موفقیت", "پایان‌نامه با موفقیت بایگانی شد.")
    self.load_archive()

# ۳-۴) پچ متد setup_archive_tab برای اضافه کردن دکمه بایگانی دستی
def patched_setup_archive_tab(self):
    layout = QVBoxLayout()
    filter_layout = QHBoxLayout()
    self.archive_search = QLineEdit()
    self.archive_search.setPlaceholderText("جستجو بر اساس نام، شماره دانشجویی...")
    self.archive_search.textChanged.connect(self.load_archive)
    self.archive_year_combo = QComboBox()
    self.archive_year_combo.addItem("همه سال‌ها")
    for y in range(2020, 2031):
        self.archive_year_combo.addItem(str(y))
    self.archive_year_combo.currentTextChanged.connect(self.load_archive)
    filter_layout.addWidget(QLabel("جستجو:"))
    filter_layout.addWidget(self.archive_search)
    filter_layout.addWidget(QLabel("سال دفاع:"))
    filter_layout.addWidget(self.archive_year_combo)

    self.archive_table = QTableWidget()
    self.archive_table.setColumnCount(7)
    self.archive_table.setHorizontalHeaderLabels(
        ["دانشجو", "شماره دانشجویی", "عنوان", "تاریخ دفاع", "نمره", "تعداد فایل‌ها", "عملیات"]
    )

    btn_layout = QHBoxLayout()
    download_btn = QPushButton("دریافت بایگانی")
    download_btn.clicked.connect(self.download_archive)
    excel_btn = QPushButton("خروجی اکسل")
    excel_btn.clicked.connect(self.export_archive_excel)
    archive_btn = QPushButton("بایگانی کردن")  # دکمه جدید
    archive_btn.clicked.connect(self.archive_selected_thesis)
    btn_layout.addWidget(download_btn)
    btn_layout.addWidget(excel_btn)
    btn_layout.addWidget(archive_btn)

    layout.addLayout(filter_layout)
    layout.addWidget(self.archive_table)
    layout.addLayout(btn_layout)
    self.archive_tab.setLayout(layout)

# ============================================================
# ۴. پچ متد save_evaluation در DefenseEvaluationDialog برای بایگانی خودکار
# ============================================================
def patched_save_evaluation(self, final):
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
    final_grade = round(avg * 5, 2)

    session = get_session()
    existing = session.query(main_module.DefenseEvaluation).filter_by(thesis_id=self.thesis.id).first()
    if not existing:
        existing = main_module.DefenseEvaluation(thesis_id=self.thesis.id, professor_id=self.thesis.supervisor_id)
        session.add(existing)
    existing.evaluation_data_json = eval_data
    existing.final_score = final_grade
    existing.is_final = final
    if final:
        self.thesis.final_grade = final_grade
        self.thesis.project_status = "دفاع شده"
        self.thesis.is_archived = True   # بایگانی خودکار
    session.commit()
    session.close()
    QMessageBox.information(self, "موفقیت", "ارزیابی ذخیره شد." if not final else "نمره نهایی ثبت و دفاع پایان یافت. پایان‌نامه بایگانی شد.")
    if final:
        self.accept()
    else:
        self.accept()

# ============================================================
# اعمال پچ‌ها روی کلاس‌ها
# ============================================================
ProfessorPanel.save_manual_progress = save_manual_progress
ProfessorPanel.setup_chapters_tab = patched_setup_chapters_tab
ProfessorPanel.load_chapters = patched_load_chapters
ProfessorPanel.show_student_detail = patched_show_student_detail

StudentPanel.load_progress = patched_load_progress

AdminPanel.load_archive = patched_load_archive
AdminPanel.download_archive_item = patched_download_archive_item
AdminPanel.archive_selected_thesis = patched_archive_selected_thesis
AdminPanel.setup_archive_tab = patched_setup_archive_tab   # اضافه شد

DefenseEvaluationDialog.save_evaluation = patched_save_evaluation   # پچ متد دفاع

# ============================================================
# اجرای اصلی برنامه
# ============================================================
if __name__ == '__main__':
    create_default_admin()
    app = QApplication(sys.argv)
    font = QFont("Vazirmatn", 10)
    app.setFont(font)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec())