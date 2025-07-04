# Smart Student Dashboard

A comprehensive, automated student management dashboard built in Google Sheets, powered by Google Apps Script.  
This project streamlines attendance, analytics, permissions, and logging for educational environments.

---

## 🚀 Features

- **Real-time Progress Tracking**
  - Students can update their own progress; mentors can verify and validate entries.
- **Insightful Leaderboards & Analytics**
  - Dynamic leaderboards and ranking sheets for instant insights into top performers and engagement.
  - Visual analytics and charts for progress trends, badge distributions, and mentor evaluations.
- **Automated Attendance Management**
  - Daily attendance columns generated automatically with easy dropdowns and conditional formatting.
  - Automated attendance percentage calculation and email notifications.
- **Multi-Level Access Control**
  - Four roles: Student, Mentor, Admin, Master Admin, each with distinct permissions.
  - Fine-grained, row-level and column-level protections enforced programmatically.
  - Automated editor/viewer management for privacy and security.
  - No student can make changes to anything except their own progress.
- **Logging & Audit**
  - Tracks edits, row changes, and logs user actions for transparency.
  - Daily performance logging and average calculations.

---

## 🛠️ How It Works

### 1. Progress Tracking & Insightful Leaderboards
- Students update their progress in real time, but can only edit their own row; mentors verify and validate these entries.
- The system auto-generates leaderboards and ranking sheets based on progress data, providing instant insights into top performers and engagement.
- Visual analytics and charts are created from the data to highlight trends, badge distributions, and mentor evaluations.

### 2. Attendance Automation
- The script automatically syncs student lists and generates new attendance columns for each day.
- Dropdowns and conditional formatting are applied for easy marking of Present/Absent.
- Attendance reports are generated and email notifications are sent automatically.

### 3. Multi-Level Permissions & Security
- **Student:** Can only edit their own row/progress; all other data is protected.
- **Mentor:** Can update and review progress for their assigned students.
- **Admin:** Can manage attendance, analytics, mentor assignments, and attendance reports.
- **Master Admin:** Has unrestricted access to all sheets and settings.
- All permissions are enforced via Google Apps Script protections and sharing settings.
- No student can alter any data except their own progress, ensuring data integrity and privacy.

### 4. Logging & Reporting
- All edits and changes are logged for audit and transparency.
- Daily and top-10 averages are calculated and displayed for performance review.

---

## 📂 Project Structure

- `Attendance.gs` – Attendance automation and email notifications.
- `Analytics.gs` – Chart generation and analytics.
- `Code.gs` – Configure permissions, manage progress dashboards & leaderboards.
- `logger.gs` – Logging, reporting, and audit trails.

---

## 🧑‍💻 Usage

1. **Copy the scripts** into your Google Sheet via Extensions > Apps Script.
2. **Set up sheet tabs** as referenced in the scripts (e.g., `Attendance`, `Analytics`, `Updates-Students`, `Admins`, etc.).
3. **Configure admin, mentor, and student emails** in the `Admins` sheet.
4. **Run the setup and automation functions** as needed.

> _See inline comments in each script for customization and advanced usage._

---

## 🔒 Privacy & Security

- No student data is shared publicly.
- All permissions are managed programmatically for maximum privacy.
- Only code is open-source; the actual dashboard is private.

---

## 📜 License

MIT License

---

## 🤝 Contributing

Pull requests and suggestions are welcome!

---

## 📧 Contact

For questions or demo requests, contact [2023mcb1294@gmail.com](mailto:2023mcb1294@gmail.com).

---