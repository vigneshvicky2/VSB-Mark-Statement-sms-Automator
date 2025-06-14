# 📤 Marks Statements SMS Automation

This Python-based desktop application automates the process of sending personalized marks statements to students via **Google Messages Web**. Built using `CustomTkinter` for a modern GUI and `Selenium` for browser automation, the tool ensures screenshots are captured and compiled into a Word document for documentation.

---

## 🚀 Features

* ✅ Upload **Excel file** containing student contacts
* ✅ Upload **Word document** with individual marks messages
* ✅ Automatically sends personalized messages using **Google Messages Web**
* ✅ Captures screenshots of each sent message
* ✅ Compiles screenshots into a **single Word document**
* ✅ Clean, user-friendly GUI with **CustomTkinter**
* ✅ Multi-threaded process to keep the UI responsive

---

## 🛠️ Tech Stack

* Python 3.x
* CustomTkinter
* Selenium WebDriver
* pandas
* python-docx
* Pillow (ImageGrab)
* tkinter (native GUI library)

---

## 📂 Project Structure

```
project/
│
├── main.py               # Main GUI and logic script
├── README.md             # Project documentation
```

---

## 📋 Prerequisites

* Python 3.x installed
* Google Chrome browser
* Required Python packages (install below)

---

## 📦 Installation

1. Clone this repository:

   ```bash
   git clone https://github.com/yourusername/marks-sms-automation.git
   cd marks-sms-automation
   ```

2. Install required packages:

   ```bash
   pip install -r requirements.txt
   ```

   Or manually install:

   ```bash
   pip install customtkinter pandas python-docx pillow selenium
   ```

3. Download ChromeDriver compatible with your Chrome version from:
   [https://chromedriver.chromium.org/downloads](https://chromedriver.chromium.org/downloads)

4. Place `chromedriver.exe` in the same folder or ensure it is in your system's PATH.

---

## 🧑‍💻 How to Use

1. **Run the application**:

   ```bash
   python main.py
   ```

2. **Upload Files**:

   * Excel File (contacts must be in a column named `Contact`)
   * Word Document (each page/message separated by an empty line)

3. **Choose screenshot saving directory**

4. **Click “Start Process”**

   * The app will open Google Messages Web
   * Scan the QR code to link your device
   * Messages will be sent, screenshots saved, and compiled

5. **Result**:

   * A Word document `Messages_Screenshots.docx` will be generated in your selected directory

---

## ⚠️ Notes

* The app assumes messages in the Word file are separated by **empty lines**.
* Contact numbers in Excel must be in a column titled **`Contact`**.
* Google Messages Web must remain open and connected during the automation.

---

## 📸 Screenshots

*(Add screenshots of your UI and generated document here if needed)*

---

## 👨‍💻 Author

**Vignesh K**
Engineering Student | Passionate about Automation and Software Solutions

---

## 📃 License

This project is open source and free to use. Feel free to contribute or customize it!

---

Let me know if you'd like the README in PDF format, or translated into Tamil or any other language!
