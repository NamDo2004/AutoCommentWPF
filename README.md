# 📚 Student Evaluation Tool - WPF Application
A Windows desktop application built with WPF that allows users to import student scores, evaluate performance, and generate automatic comments based on predefined criteria.

# 🖼️ Application Overview
![Image](https://github.com/user-attachments/assets/6b11e70b-9c63-4257-a41b-ca6bcb6949e0)

# ⚙️ Features
## ✅ Import from Excel
- Load student data (name, speaking, listening, reading-writing scores) from an Excel file.
- Automatically calculates total score and performance level (T for good, others for weak).
![Image](https://github.com/user-attachments/assets/b00014dc-7f67-4db2-828c-c155ab88a53e)

# ✅ Export to Excel
- Export the updated student data including generated evaluations and comments to a new Excel file.
![image](https://github.com/user-attachments/assets/f5e4e292-dc66-4aef-bc10-702ceeced3e4)

# ✅ Auto Generate Comments
- Generates comments for each student based on their performance level using a nhan_xet.json file.
- Uses different comment pools for "Good" and "Needs Improvement".
- ![Image](https://github.com/user-attachments/assets/a917aafc-68e2-415b-ae64-3a4bbdf31d5e)

# 🧠 Technologies Used
- C#
- WPF (Windows Presentation Foundation)
- Newtonsoft.Json (for JSON parsing)
- Microsoft.Office.Interop.Excel (or ClosedXML for Excel interaction)

# 📌 How to Use
- Click "Import" to load student data from Excel.
- Click "Auto nhận xét" to generate comments based on scores.
- Click "Export" to save the updated data.
