🧪 Cold Test Logger
A Python-based GUI tool for logging hardware test results efficiently into Excel

📌 Overview
Cold Test Logger is a desktop application built using Python that automates the process of recording and managing cold test measurements for hardware components. It is designed for use in manufacturing/testing environments where large volumes of impedance data need to be organized, validated, and stored systematically.

✅ Originally developed during my internship at SFO Technologies, Bangalore, this tool aims to reduce manual Excel entry, prevent human error, and generate quick test reports.

🎯 Features
🔢 Input validation using Regex

🧮 Automatic PASS/FAIL result generation

📊 Dynamic Excel file creation using OpenPyXL

📁 Save test data in structured rows/columns

📄 Export reports in .txt format

🖥️ Built with Tkinter for a simple, intuitive GUI

💾 Converted to .exe for standalone use with PyInstaller

🚀 Tech Stack
Python- Core programming language
Tkinter- GUI framework
OpenPyXL- Excel read/write functionality
Regex- Input format validation
PyInstaller	- Script-to-EXE conversion

🛠 How It Works

User enters:
    Part Number (validated)
    Serial Number (validated)
    Number of test points

For each point, inputs:
    Supply/Component
    Test Point
    Expected Impedance
    Measured Value
    Tester Name

On completion:
    Excel sheet is generated with all entries
    PASS/FAIL status calculated and highlighted
    Optional .txt report is generated for summary