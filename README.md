# Comprehensive University Thesis Management System

A full-featured desktop application developed with Python and PyQt6 to streamline and automate the complex workflows of university thesis management. The system is designed to facilitate seamless interactions between students, supervisors, and educational administrators.

## 🚀 Key Features
- **Role-Based Access Control:** Dedicated panels for Admin, Professor, and Student roles with specific permissions.
- **Student Registration & Approval:** Secure registration with document/photo upload, followed by admin verification and automated account creation.
- **Supervision Requests:** Students can request specific professors, and upon approval, a thesis project is automatically created.
- **Meeting & File Management:** Track consultation sessions, assign tasks, and manage various thesis file uploads (Proposals, Chapters, etc.).
- **Defense Evaluation:** A multi-criteria grading system with automated final score calculation.
- **Reporting & Archiving:** Export data to Excel, generate analytical reports, and automatically archive defended theses.
- **Advanced Utilities:** Automated email notifications (smtplib), Jalali (Persian) date support (jdatetime), and automated database backup.

## 🛠️ Technologies Used
- **Language:** Python 3.x
- **GUI Framework:** PyQt6
- **Database & ORM:** SQLAlchemy, SQLite
- **Testing:** Pytest (100% test coverage on core modules)
- **Libraries:** pandas, openpyxl, jdatetime, smtplib

## 📂 Architecture & Structure
The project follows a modular, layered architecture (Service-Repository pattern):
- `domain/`: Core data models and business logic.
- `engines/`: Core processing and computation logic.
- `infrastructure/`: Database connection and repository pattern implementation.
- `services/`: Independent services (Email, Backup, Report generation).
- `ui/`: PyQt6 graphical user interface and styling.

## 👨‍💻 Developer
**Reyhaneh Jafari** - Computer Software Engineering
- [LinkedIn](YOUR_LINKEDIN_URL) | [Email](mailto:jafarireyhane214@gmail.com)
