# 📄 Invitation Letter Generator

A Python-based **GUI application** to generate personalized PDF invitations for multiple companies in bulk, using Excel data and Word templates.

## ✨ Features
- 📊 Reads HR details from an Excel file
- 📝 Supports multiple invitation templates (`2026 Batch`, `BPharma`)
- 🔄 Automatically replaces placeholders in `.docx` templates
- 📂 Creates structured company folders for output
- 📄 Converts `.docx` invitations to PDF
- 🎨 User-friendly Tkinter GUI with dark mode
- 👋 Automatic salutation detection (Sir/Ma'am)

## 🛠️ Requirements
Install dependencies using:
```bash
pip install pandas python-docx docx2pdf openpyxl
📁 Invitation-Generator
 ├── template.xlsx                     # Excel file with HR details
 ├── 2026_Invitation_XYZ.docx        # Word template for 2026 batch
 ├── Bpharma_Invitation_XYZ.docx     # Word template for BPharma
 ├── output_pdfs/                       # Auto-created output folder
 ├── invitation_generator.py           # Main script
 └── README.md
📊 Excel Format
Ensure your template.xlsx contains these columns:

HR Name	Designation	Company Name	Gender
Mr. John Doe	HR Manager	ABC Pvt. Ltd.	Male
Ms. Jane Roe	Recruiter	XYZ Ltd.	Female

🚀 How to Use
Place template.xlsx and Word templates in the same folder as the script.

Run the script:

bash
Copy
Edit
python invitation_generator.py
In the GUI:

Select the desired invitation template

Click "Generate Invitations"

PDFs will be saved in output_pdfs with company-specific subfolders.

🖼️ Example
After running, your output might look like:

Copy
Edit
output_pdfs/
 ├── ABC Pvt. Ltd/
 │    └── 1/
 │         └── ABC Pvt. Ltd_Invitation_XYZ.pdf
 ├── XYZ Ltd/
 │    └── 1/
 │         └── XYZ Ltd_Invitation_XYZ.pdf
📜 License
This project is for educational and organizational purposes.
Feel free to modify it for your needs.

