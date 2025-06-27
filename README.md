# 🧪 CSP Validator & Magento Builder — Zero Excuses Edition

A PyQt6-powered desktop application to validate Customer Specific Pricing (CSP) updates and generate Magento import-ready files. Designed for maximum speed, clarity, and zero manual intervention.

---

## 🚀 Features

- Load Magento CSP export file (`.xlsx`) with fixed headers
- Load flexible-format Supplier CSP update file
- Map Supplier columns dynamically: SKU, Price, Rebate, and Customer Group
- Validation output:
  - ✅ No Change
  - 🔁 Updated Existing
  - 🆕 New
  - 📉 Disabled
- Automatic normalization of % rebate formats
- Export validation report to Excel (`Summary`, `No Change`, `Disabled`, `Updated Existing`, `New`)
- Generate Magento import-ready Excel file with:
  - `Disabled_Profile`
  - `Updated_Profile`
  - `New_Profile`
- Real-time verbose logging to GUI
- Dark mode enabled

---

## 🛠 Installation

### 1. Clone the Repo
```bash
git clone https://github.com/your-org/csp-validator.git
cd csp-validator
```

### 2. Set Up Python Environment
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

---

## 🏗 Build Executable

Compile into a standalone `.exe`:
```bash
.\.venv\Scripts\pyinstaller.exe --noconfirm --onefile --windowed --clean --name "CSP_Validator" main.py
```

Result: `./dist/CSP_Validator.exe`

---

## 🧠 How to Use

### Part 1: Validate Pricing
1. **Launch the app**
2. **Load Magento CSP export** (with fixed headers)
3. **Load Supplier CSP update file**
4. **Map**:
   - SKU → part number
   - Price → fixed or contract price
   - Rebate → % off
   - Customer Group → customer tier/group
5. **Click Run Validation**
6. Output Excel saved in your Documents folder (`CSP_Validation_YYYYMMDD_HHMMSS.xlsx`)

### Part 2: Generate Magento Files
1. After validation completes, click **Generate Magento Upload Files**
2. Output file (`Magento_Upload_YYYYMMDD_HHMMSS.xlsx`) will contain:
   - `Disabled_Profile`
   - `Updated_Profile`
   - `New_Profile`

---

## 📄 Magento Required Columns

Magento CSP file must contain:
```
SupplierPartnumber | fixedprice | Rebate | CustomerGroup
```

---

## 📂 Output Location

All `.xlsx` output files are saved in:
```
C:\Users\<YourUsername>\Documents\
```

---

## 📜 License

MIT — Automate or die trying.

---

## 👨‍💻 Author

**Joshua Taitt**  
Automation Specialist @ Neta Scientific
