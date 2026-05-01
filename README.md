# ONSkills Certificate Automation System

Yeh ek automated system hai jo Excel data se certificates generate karta hai.

## 🚀 Quick Setup (Naye System Ke Liye)

Agar aap kisi naye computer par ise setup kar rahe hain, toh niche diye gaye steps follow karein:

### 1. Python Install Karein
Sabse pehle system mein Python install hona chahiye. Aap [python.org](https://www.python.org/) se download kar sakte hain.

### 2. Libraries Install Karein
Terminal ya Command Prompt (CMD) open karein aur ye command chalaein:
```bash
pip install Pillow openpyxl
```

### 3. Folder Structure
Ensure karein ki aapke folder ka structure aisa ho:
```
Certificates/
├── template/
│   └── certificate_template.jpg     # Aapka certificate ka blank image
├── input/
│   └── (Yahan aapki Excel file hogi)
├── output/
│   └── (Yahan generated certificates save honge)
├── generate_certificates.py          # Main Python script
└── requirements.txt
```

---

## 📊 Excel File Ka Format

Aapki Excel file mein niche diye gaye **Headers** (Line 1) hone chahiye. Spelling ka dhyan rakhein:

1. `Certificate No`
2. `Name`
3. `Title` (Mr. / Ms. / Mrs.)
4. `Father/Spouse Name`
5. `Relation` (S/o / D/o / W/o)
6. `Aadhar No`
7. `Job Role`
8. `Duration`
9. `Training Center`
10. `Sponsored By`
11. `District`
12. `State`
13. `Place of Issue`
14. `Date of Issue`

---

## 🛠️ Commands Kaise Chalayein?

### A. Certificates Generate Karein
Input folder mein Excel file rakhein aur ye command chalaein:
```bash
python generate_certificates.py
```

### B. Text Position Check Karein (Calibration)
Agar aapko dekhna hai ki text sahi jagah aa raha hai ya nahi:
```bash
python generate_certificates.py --calibrate
```
*Iski output `output/_calibration_preview.jpg` mein dikhegi.*

### C. Sample Excel Banayein
Test karne ke liye dummy data ki Excel banani ho:
```bash
python generate_certificates.py --sample
```

---

## 💡 Zaroori Baatein (Important Notes)

1. **Date Format**: Excel mein Date of Issue ka format kuch bhi ho, script usey automatic `20 Apr 2026` jaisa bana degi.
2. **Quality**: Script automatic **2.0x Upscaling** karti hai taaki PDF 100% zoom par clear aur bade size mein dikhe.
3. **Folders**: Certificates automatic **Date of Issue** ke naam se bane folders mein save honge.
4. **Coordinates**: Agar text ki position thodi upar-neeche karni ho, toh `generate_certificates.py` mein `FIELD_POSITIONS` section mein (x, y) values badal sakte hain.
   - **x** badhane se text **Right** jayega.
   - **y** badhane se text **Neeche** jayega.
5. **MS Paint**: Sahi (x, y) coordinates nikalne ke liye template image ko MS Paint mein kholein aur mouse pointer ki location niche-left mein dekhein.

---

**Developed for:** Cee Vision Technologies Pvt. Ltd.
