A **Python-based Attendance Tracking System** using **QR code scanning** and **Excel as the database**. This system is designed for **schools, offices, and events** to automate attendance recording, prevent errors, and make history tracking simple.

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](https://opensource.org/licenses/MIT) ![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)

## 📦 Installation

To set up the project, follow these steps:

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/attendance-tracking-system.git
   cd attendance-tracking-system
   ```
2. Install Python dependencies:
   ```bash
   pip install pillow pyzbar opencv-python openpyxl
   ```

## ✨ Features

- **QR Code Scanning** – Uses your webcam to read QR codes for quick check-in/out.
- **Multiple Modes** – AM Time In, AM Time Out, PM Time In, PM Time Out.
- **Excel Database Storage** – Stores attendance data locally in an Excel file (`attendance.xlsx`).
- **Color-Coded Export** – Attendance history exports with color codes for Present, Absent, and Cutting Class.
- **Duplicate Prevention** – Prevents the same QR from being scanned repeatedly in a short time.
- **Password-Protected Admin Panel** – Allows viewing history, exporting data, and resetting records.
- **Real-Time Interface** – Live camera feed and instant confirmation messages.
- **Automatic Status Tracking** – Marks students as Present, Absent, or Cutting Class based on scan records.

## 📖 Usage/Examples

To use the attendance tracking system, run the application:
```bash
python main.py
```

Follow the on-screen instructions to scan QR codes and record attendance.

## 📂 Database Information

This system uses **Excel (`.xlsx`)** as the database, powered by the [`openpyxl`](https://openpyxl.readthedocs.io/) library.

- **File Name:** `attendance.xlsx`
- **Location:** Saved inside the **Documents** folder of the current user.
- **Structure:**

| Name           | AM Time In       | AM Time Out       | PM Time In       | PM Time Out       | Status              |
|----------------|------------------|-------------------|------------------|-------------------|---------------------|
| Jochen Galera  | 2025-08-15 07:58 | 2025-08-15 12:01  | 2025-08-15 13:02 | 2025-08-15 16:05  | AM Present, PM Present |

- **Status Rules:**
  - **AM Present** → Has both AM Time In & AM Time Out.
  - **AM Cutting Class** → Has AM Time In but no AM Time Out.
  - **AM Absent** → No AM Time In.
  - **PM Present / Cutting Class / Absent** → Same rules applied for PM session.

- **Export Feature:**
  - Exports to `attendance_history_YYYY-MM-DD_HH-MM-SS.xlsx`.
  - Color codes rows:
    - 🟩 Green – Full Present
    - 🟧 Orange – Partially Present
    - 🟥 Red – Cutting Class
    - 🟨 Yellow – Full Absent

## 🛠 Requirements

Install Python dependencies using:

```bash
pip install pillow pyzbar opencv-python openpyxl
```

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
