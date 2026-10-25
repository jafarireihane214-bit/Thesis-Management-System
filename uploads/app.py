# thesis_app.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import os
import shutil
from datetime import datetime

# ==================== راه‌اندازی اولیه ====================
DB_NAME = "thesis.db"
PDF_FOLDER = "pdfs"
os.makedirs(PDF_FOLDER, exist_ok=True)

# ==================== دیتابیس ====================
def init_db():
    conn = sqlite3.connect(DB_NAME)
    conn.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY,
        username TEXT UNIQUE,
        password TEXT,
        full_name TEXT,
        role TEXT,
        is_active INTEGER DEFAULT 1
    )''')
    conn.execute('''CREATE TABLE IF NOT EXISTS theses (
        id INTEGER PRIMARY KEY,
        student_id INTEGER,
        title TEXT,
        pdf_path TEXT,
        status TEXT DEFAULT 'pending'
    )''')
    conn.execute('''CREATE TABLE IF NOT EXISTS messages (
        id INTEGER PRIMARY KEY,
        sender_id INTEGER,
        receiver_id INTEGER,
        message TEXT,
        date TEXT
    )''')
    
    # استاد پیش‌فرض
    conn.execute("DELETE FROM users WHERE role='professor'")
    conn.execute("INSERT INTO users (username, password, full_name, role) VALUES (?,?,?,?)",
                 ('admin', '1234', 'دکتر احمدی', 'professor'))
    conn.commit()
    conn.close()

def login_student(username, password):
    conn = sqlite3.connect(DB_NAME)
    row = conn.execute("SELECT id, full_name FROM users WHERE username=? AND password=? AND role='student'", (username, password)).fetchone()
    conn.close()
    return row

def login_professor(username, password):
    conn = sqlite3.connect(DB_NAME)
    row = conn.execute("SELECT id, full_name FROM users WHERE username=? AND password=? AND role='professor'", (username, password)).fetchone()
    conn.close()
    return row

def register_student(full_name, username, password):
    conn = sqlite3.connect(DB_NAME)
    try:
        conn.execute("INSERT INTO users (username, password, full_name, role) VALUES (?,?,?,?)",
                     (username, password, full_name, 'student'))
        conn.commit()
        return True, "ثبت نام موفق"
    except:
        return False, "نام کاربری تکراری است"
    finally:
        conn.close()

def save_thesis(student_id, title, pdf_path):
    conn = sqlite3.connect(DB_NAME)
    conn.execute("DELETE FROM theses WHERE student_id=?", (student_id,))
    conn.execute("INSERT INTO theses (student_id, title, pdf_path) VALUES (?,?,?)", (student_id, title, pdf_path))
    conn.commit()
    conn.close()

def get_thesis(student_id):
    conn = sqlite3.connect(DB_NAME)
    row = conn.execute("SELECT title, pdf_path, status FROM theses WHERE student_id=?", (student_id,)).fetchone()
    conn.close()
    return row

def approve_thesis(student_id):
    conn = sqlite3.connect(DB_NAME)
    conn.execute("UPDATE theses SET status='approved' WHERE student_id=?", (student_id,))
    conn.commit()
    conn.close()

def get_all_students():
    conn = sqlite3.connect(DB_NAME)
    rows = conn.execute("SELECT id, full_name, username FROM users WHERE role='student'").fetchall()
    conn.close()
    return rows

def delete_student(student_id):
    conn = sqlite3.connect(DB_NAME)
    conn.execute("DELETE FROM users WHERE id=?", (student_id,))
    conn.execute("DELETE FROM theses WHERE student_id=?", (student_id,))
    conn.commit()
    conn.close()

def send_message(sender, receiver, msg):
    conn = sqlite3.connect(DB_NAME)
    conn.execute("INSERT INTO messages (sender_id, receiver_id, message, date) VALUES (?,?,?, datetime('now'))",
                 (sender, receiver, msg))
    conn.commit()
    conn.close()

def get_messages(user1, user2):
    conn = sqlite3.connect(DB_NAME)
    rows = conn.execute('''SELECT message, date, sender_id FROM messages
                           WHERE (sender_id=? AND receiver_id=?) OR (sender_id=? AND receiver_id=?)
                           ORDER BY id''', (user1, user2, user2, user1)).fetchall()
    conn.close()
    return rows

def save_pdf(file_path, student_id):
    if not file_path:
        return None
    ext = os.path.splitext(file_path)[1]
    new_name = f"student_{student_id}_{datetime.now().strftime('%Y%m%d%H%M%S')}{ext}"
    new_path = os.path.join(PDF_FOLDER, new_name)
    shutil.copy(file_path, new_path)
    return new_path

# ==================== صفحه لاگین ====================
class LoginApp:
    def __init__(self, root):
        self.root = root
        self.root.title("سامانه مدیریت پایان نامه")
        self.root.geometry("400x400")
        self.root.configure(bg="#f0f2f5")
        self.role = "student"
        self.create_widgets()
    
    def create_widgets(self):
        frame = tk.Frame(self.root, bg="#f0f2f5")
        frame.place(relx=0.5, rely=0.4, anchor="center")
        
        tk.Label(frame, text="سامانه مدیریت پایان نامه", font=("Segoe UI", 16, "bold"), bg="#f0f2f5").pack(pady=20)
        
        role_frame = tk.Frame(frame, bg="#f0f2f5")
        role_frame.pack()
        tk.Button(role_frame, text="دانشجو", command=lambda: self.set_role("student"), bg="#2b5b84", fg="white", width=10).pack(side="left", padx=5)
        tk.Button(role_frame, text="استاد", command=lambda: self.set_role("professor"), bg="#2b5b84", fg="white", width=10).pack(side="left", padx=5)
        
        tk.Label(frame, text="نام کاربری:", bg="#f0f2f5").pack(pady=(15,0))
        self.entry_user = tk.Entry(frame)
        self.entry_user.pack()
        
        tk.Label(frame, text="رمز عبور:", bg="#f0f2f5").pack()
        self.entry_pass = tk.Entry(frame, show="*")
        self.entry_pass.pack()
        
        tk.Button(frame, text="ورود", command=self.login, bg="#27ae60", fg="white", width=15).pack(pady=15)
        tk.Button(frame, text="ثبت نام دانشجو", command=self.register, bg="#e67e22", fg="white", width=15).pack()
        
        tk.Label(frame, text="استاد: admin / 1234", font=("Segoe UI", 9), fg="gray", bg="#f0f2f5").pack(pady=15)
    
    def set_role(self, role):
        self.role = role
    
    def login(self):
        u = self.entry_user.get().strip()
        p = self.entry_pass.get().strip()
        if self.role == "student":
            user = login_student(u, p)
        else:
            user = login_professor(u, p)
        
        if user:
            self.root.destroy()
            new_root = tk.Tk()
            if self.role == "student":
                StudentApp(new_root, user)
            else:
                ProfessorApp(new_root, user)
            new_root.mainloop()
        else:
            messagebox.showerror("خطا", "اطلاعات صحیح نیست")
    
    def register(self):
        win = tk.Toplevel(self.root)
        win.title("ثبت نام دانشجو")
        win.geometry("350x350")
        win.configure(bg="#f0f2f5")
        
        tk.Label(win, text="ثبت نام دانشجو", font=("Segoe UI", 14, "bold"), bg="#f0f2f5").pack(pady=10)
        
        tk.Label(win, text="نام کامل:", bg="#f0f2f5").pack()
        name_entry = tk.Entry(win)
        name_entry.pack()
        
        tk.Label(win, text="نام کاربری:", bg="#f0f2f5").pack()
        user_entry = tk.Entry(win)
        user_entry.pack()
        
        tk.Label(win, text="رمز عبور:", bg="#f0f2f5").pack()
        pass_entry = tk.Entry(win, show="*")
        pass_entry.pack()
        
        def do_register():
            if not name_entry.get() or not user_entry.get() or not pass_entry.get():
                messagebox.showerror("خطا", "همه فیلدها را پر کنید")
                return
            ok, msg = register_student(name_entry.get(), user_entry.get(), pass_entry.get())
            messagebox.showinfo("نتیجه", msg)
            if ok:
                win.destroy()
        
        tk.Button(win, text="ثبت نام", command=do_register, bg="#27ae60", fg="white").pack(pady=20)

# ==================== صفحه استاد ====================
class ProfessorApp:
    def __init__(self, root, user):
        self.root = root
        self.user_id = user[0]
        self.user_name = user[1]
        self.selected_student = None
        self.root.title(f"پایان نامه - استاد: {self.user_name}")
        self.root.geometry("900x600")
        self.root.configure(bg="#f0f2f5")
        
        self.setup_sidebar()
        self.content = tk.Frame(self.root, bg="#f0f2f5")
        self.content.pack(side="right", fill="both", expand=True)
        self.show_students()
    
    def setup_sidebar(self):
        sidebar = tk.Frame(self.root, bg="#1e3a5f", width=180)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        tk.Label(sidebar, text="📚", font=("Segoe UI", 24), bg="#1e3a5f", fg="white").pack(pady=20)
        tk.Label(sidebar, text=self.user_name, bg="#1e3a5f", fg="white").pack()
        
        for text, cmd in [("👨‍🎓 دانشجویان", self.show_students), ("💬 پیام‌ها", self.show_messages), ("🚪 خروج", self.logout)]:
            tk.Button(sidebar, text=text, bg="#1e3a5f", fg="white", bd=0, anchor="w", padx=20, pady=10, command=cmd).pack(fill="x")
    
    def clear(self):
        for w in self.content.winfo_children():
            w.destroy()
    
    def show_students(self):
        self.clear()
        tk.Label(self.content, text="لیست دانشجویان", font=("Segoe UI", 18, "bold"), bg="#f0f2f5").pack(pady=10)
        
        self.tree = ttk.Treeview(self.content, columns=("نام", "نام کاربری"), show="headings", height=15)
        self.tree.heading("نام", text="نام")
        self.tree.heading("نام کاربری", text="نام کاربری")
        self.tree.pack(fill="both", expand=True, padx=20, pady=10)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        
        btn_frame = tk.Frame(self.content, bg="#f0f2f5")
        btn_frame.pack(fill="x", padx=20, pady=5)
        tk.Button(btn_frame, text="✅ تأیید پایان نامه", command=self.approve, bg="#27ae60", fg="white").pack(side="left", padx=5)
        tk.Button(btn_frame, text="❌ حذف دانشجو", command=self.delete_student, bg="#e74c3c", fg="white").pack(side="left", padx=5)
        
        self.load_students()
    
    def load_students(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for s in get_all_students():
            self.tree.insert("", "end", values=(s[1], s[2]), iid=s[0])
    
    def on_select(self, event):
        sel = self.tree.selection()
        if sel:
            self.selected_student = int(sel[0])
    
    def approve(self):
        if not self.selected_student:
            messagebox.showerror("خطا", "دانشجویی انتخاب کنید")
            return
        approve_thesis(self.selected_student)
        messagebox.showinfo("موفق", "پایان نامه تأیید شد")
    
    def delete_student(self):
        if not self.selected_student:
            return
        if messagebox.askyesno("حذف", "مطمئنید؟"):
            delete_student(self.selected_student)
            self.selected_student = None
            self.load_students()
    
    def show_messages(self):
        self.clear()
        tk.Label(self.content, text="پیام با دانشجو", font=("Segoe UI", 18, "bold"), bg="#f0f2f5").pack(pady=10)
        
        students = get_all_students()
        left = tk.Frame(self.content, bg="#f0f2f5")
        left.pack(side="left", fill="y", padx=10)
        tk.Label(left, text="انتخاب دانشجو:", bg="#f0f2f5").pack()
        student_list = tk.Listbox(left, height=15)
        student_list.pack()
        for s in students:
            student_list.insert(tk.END, s[1], s[0])
        
        right = tk.Frame(self.content, bg="#f0f2f5")
        right.pack(side="right", fill="both", expand=True)
        msg_text = tk.Text(right, wrap="word", height=15, state="disabled")
        msg_text.pack(fill="both", expand=True)
        entry = tk.Text(right, height=4)
        entry.pack(fill="x")
        
        def load_msgs(other_id):
            msg_text.config(state="normal")
            msg_text.delete("1.0", tk.END)
            for msg, date, sender in get_messages(self.user_id, other_id):
                sender_name = "من" if sender == self.user_id else "دانشجو"
                msg_text.insert(tk.END, f"{sender_name} ({date}):\n{msg}\n\n")
            msg_text.config(state="disabled")
        
        def on_select(ev):
            sel = student_list.curselection()
            if sel:
                other_id = student_list.get(sel[1])
                load_msgs(other_id)
                entry.other_id = other_id
        
        def send():
            if hasattr(entry, 'other_id'):
                msg = entry.get("1.0", tk.END).strip()
                if msg:
                    send_message(self.user_id, entry.other_id, msg)
                    entry.delete("1.0", tk.END)
                    load_msgs(entry.other_id)
        
        student_list.bind("<<ListboxSelect>>", on_select)
        tk.Button(right, text="ارسال", command=send, bg="#27ae60", fg="white").pack()
    
    def logout(self):
        self.root.destroy()
        new = tk.Tk()
        LoginApp(new)
        new.mainloop()

# ==================== صفحه دانشجو ====================
class StudentApp:
    def __init__(self, root, user):
        self.root = root
        self.user_id = user[0]
        self.user_name = user[1]
        self.pdf_path = None
        self.root.title(f"پایان نامه - دانشجو: {self.user_name}")
        self.root.geometry("700x500")
        self.root.configure(bg="#f0f2f5")
        
        self.setup_sidebar()
        self.content = tk.Frame(self.root, bg="#f0f2f5")
        self.content.pack(side="right", fill="both", expand=True)
        self.show_thesis()
    
    def setup_sidebar(self):
        sidebar = tk.Frame(self.root, bg="#1e3a5f", width=180)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        tk.Label(sidebar, text="📚", font=("Segoe UI", 24), bg="#1e3a5f", fg="white").pack(pady=20)
        tk.Label(sidebar, text=self.user_name, bg="#1e3a5f", fg="white").pack()
        
        for text, cmd in [("📄 پایان نامه", self.show_thesis), ("💬 پیام به استاد", self.show_messages), ("🚪 خروج", self.logout)]:
            tk.Button(sidebar, text=text, bg="#1e3a5f", fg="white", bd=0, anchor="w", padx=20, pady=10, command=cmd).pack(fill="x")
    
    def clear(self):
        for w in self.content.winfo_children():
            w.destroy()
    
    def show_thesis(self):
        self.clear()
        tk.Label(self.content, text="پایان نامه من", font=("Segoe UI", 18, "bold"), bg="#f0f2f5").pack(pady=15)
        
        thesis = get_thesis(self.user_id)
        frame = tk.Frame(self.content, bg="white", padx=20, pady=20, relief="ridge")
        frame.pack(pady=20, padx=50, fill="x")
        
        tk.Label(frame, text="عنوان:", bg="white").pack(anchor="w")
        self.title_entry = tk.Entry(frame, width=60)
        self.title_entry.pack(fill="x", pady=5)
        if thesis:
            self.title_entry.insert(0, thesis[0] or "")
        
        self.file_label = tk.Label(frame, text="هیچ فایلی انتخاب نشده", fg="gray", bg="white")
        self.file_label.pack()
        
        def select():
            path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
            if path:
                self.pdf_path = save_pdf(path, self.user_id)
                self.file_label.config(text=os.path.basename(path), fg="green")
        
        tk.Button(frame, text="انتخاب فایل", command=select, bg="#2b5b84", fg="white").pack(pady=5)
        
        def save():
            if self.title_entry.get():
                save_thesis(self.user_id, self.title_entry.get(), self.pdf_path)
                messagebox.showinfo("موفق", "ذخیره شد")
        tk.Button(frame, text="ذخیره", command=save, bg="#27ae60", fg="white").pack(pady=15)
        
        if thesis and thesis[2] == "approved":
            tk.Label(self.content, text="✅ پایان نامه شما تأیید شده است", fg="green", bg="#f0f2f5").pack()
    
    def show_messages(self):
        self.clear()
        tk.Label(self.content, text="ارسال پیام به استاد", font=("Segoe UI", 18, "bold"), bg="#f0f2f5").pack(pady=15)
        
        frame = tk.Frame(self.content, bg="white", padx=20, pady=20, relief="ridge")
        frame.pack(pady=20, padx=50, fill="both", expand=True)
        
        tk.Label(frame, text="پیام‌های قبلی:", bg="white").pack(anchor="w")
        msg_text = tk.Text(frame, wrap="word", height=10, state="disabled")
        msg_text.pack(fill="both", expand=True, pady=5)
        
        tk.Label(frame, text="پیام جدید:", bg="white").pack(anchor="w")
        entry = tk.Text(frame, height=4)
        entry.pack(fill="x", pady=5)
        
        def load():
            msg_text.config(state="normal")
            msg_text.delete("1.0", tk.END)
            for msg, date, sender in get_messages(self.user_id, 1):  # 1 = استاد
                sender_name = "من" if sender == self.user_id else "استاد"
                msg_text.insert(tk.END, f"{sender_name} ({date}):\n{msg}\n\n")
            msg_text.config(state="disabled")
        
        def send():
            msg = entry.get("1.0", tk.END).strip()
            if msg:
                send_message(self.user_id, 1, msg)
                entry.delete("1.0", tk.END)
                load()
        
        tk.Button(frame, text="ارسال", command=send, bg="#27ae60", fg="white").pack()
        load()
    
    def logout(self):
        self.root.destroy()
        new = tk.Tk()
        LoginApp(new)
        new.mainloop()

# ==================== اجرا ====================
if __name__ == "__main__":
    init_db()
    root = tk.Tk()
    LoginApp(root)
    root.mainloop()