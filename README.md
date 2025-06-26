# 🧪 CSP Validator — Zero Excuses Edition

A PyQt6-powered desktop app that validates Customer Specific Pricing (CSP) updates from suppliers against your Magento CSP export. Built for speed, clarity, and zero manual work.

---

## 🚀 Features

- Load Magento CSP file (`.xlsx`)
- Load Supplier CSP update file (flexible headers)
- Column mapping for SKU, Price, and Rebate
- Handles:
  - ✅ No Change
  - 🔁 Updated Existing
  - 🆕 New Additions
  - 📉 Disabled (missing in supplier file)
- Normalizes rebate format (e.g., 0.1 → 10%)
- Exports summary and categorized sheets to Excel
- Logs everything in real-time
- Runs in dark mode for retinal sanity

---

## 📦 Installation

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

## 🛠 Build Executable

To create a standalone `.exe`:

```bash
.\.venv\Scripts\pyinstaller.exe --noconfirm --onefile --windowed --clean --name "CSP_Validator" main.py
```

Output: `./dist/CSP_Validator.exe`

---

## 📂 Output

On run, the tool saves an Excel report in your **Documents** folder:
```
C:\Users\<You>\Documents\CSP_Validation_YYYYMMDD_HHMMSS.xlsx
```

Sheets:
- `Summary`
- `No Change`
- `Updated Existing`
- `Disabled`
- `New`

---

## ❓ How to Use

1. **Launch the app** (via script or `CSP_Validator.exe`)
2. **Load Files**  
   - Load Magento file (exported `.xlsx` with fixed headers)  
   - Load Supplier file (update file with dynamic headers)
3. **Map Columns**  
   - SKU → supplier’s part number column  
   - Price → supplier’s contract/fixed price column  
   - Rebate → % rebate or markdown column
4. **Run Validation**
5. **Get Output** in your Documents folder

---

