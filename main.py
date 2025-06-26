import sys, os, logging, pandas as pd
from datetime import datetime
from PyQt6.QtWidgets import (QApplication, QFileDialog, QMainWindow, QWidget, QLabel,
                             QPushButton, QComboBox, QTextEdit, QVBoxLayout, QHBoxLayout,
                             QMessageBox)
from PyQt6.QtCore import pyqtSignal, QObject
from PyQt6.QtGui import QPalette, QColor


class LogSignal(QObject):
    message = pyqtSignal(str)


class QTextEditHandler(logging.Handler):
    def __init__(self, sig):
        super().__init__()
        self.sig = sig

    def emit(self, rec):
        self.sig.message.emit(self.format(rec))


class CSPValidator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CSP Validator — Zero Excuses")
        self.resize(820, 520)
        self.mag_df = self.sup_df = None
        self.mag_cols = []
        self.build_ui()
        self.dark_mode()
        self.bind_logs()

    def build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        self.mag_label = QLabel("Magento File: ✖ Not Loaded")
        self.sup_label = QLabel("Supplier File: ✖ Not Loaded")
        load_mag = QPushButton("Load Magento File")
        load_sup = QPushButton("Load Supplier File")
        load_mag.clicked.connect(self.load_mag)
        load_sup.clicked.connect(self.load_sup)

        self.sku_box = QComboBox()
        self.price_box = QComboBox()
        self.rebate_box = QComboBox()
        for b in (self.sku_box, self.price_box, self.rebate_box):
            b.setEnabled(False)
            b.setToolTip("Map this to the corresponding column in the Supplier File.")

        run_btn = QPushButton("Run Validation")
        run_btn.clicked.connect(self.run_validation)

        about_btn = QPushButton("ℹ️ About / How to Use")
        about_btn.clicked.connect(self.show_about)

        self.log = QTextEdit()
        self.log.setReadOnly(True)

        lay = QVBoxLayout(central)
        file_grid = QVBoxLayout()
        for btn, lbl in ((load_mag, self.mag_label), (load_sup, self.sup_label)):
            row = QHBoxLayout()
            row.addWidget(btn)
            row.addWidget(lbl, 1)
            file_grid.addLayout(row)
        lay.addLayout(file_grid)

        map_row = QHBoxLayout()
        map_row.addWidget(QLabel("Supplier SKU Column:"))
        map_row.addWidget(self.sku_box)
        map_row.addWidget(QLabel("Supplier Price Column:"))
        map_row.addWidget(self.price_box)
        map_row.addWidget(QLabel("Supplier Rebate Column:"))
        map_row.addWidget(self.rebate_box)
        lay.addLayout(map_row)

        lay.addWidget(about_btn)
        lay.addWidget(run_btn)
        lay.addWidget(QLabel("Log:"))
        lay.addWidget(self.log)

    def show_about(self):
        text = (
            "🐾 **CSP Validator — How to Use** 🐾\n\n"
            "This tool compares Customer Specific Pricing (CSP) between your Magento export "
            "and an updated CSP file from a supplier.\n\n"
            "📂 Step 1: Load Files\n"
            "    - Click 'Load Magento File' and select your export from Magento.\n"
            "    - Click 'Load Supplier File' and select the CSP update file from the supplier.\n\n"
            "🧩 Step 2: Map Columns\n"
            "    - Map the supplier's SKU, Price, and Rebate columns to match Magento's format.\n"
            "    - Use the dropdowns to select which column matches what.\n\n"
            "🚀 Step 3: Run Validation\n"
            "    - Click 'Run Validation' to process the comparison.\n"
            "    - A summary and 4 categorized sheets will be exported:\n"
            "        ✅ No Change — CSPs that match exactly.\n"
            "        📉 Disabled — Items removed from the new file.\n"
            "        🆕 New — Items newly added by supplier.\n"
            "        🔁 Updated Existing — CSPs that changed in price or rebate.\n\n"
            "📄 Output:\n"
            "    - File saved in your Documents folder.\n"
            "    - First sheet is a summary count of all categories.\n\n"
                    )
        QMessageBox.information(self, "About / How to Use", text)

    def dark_mode(self):
        p = self.palette()
        p.setColor(QPalette.ColorRole.Window, QColor(30, 30, 30))
        p.setColor(QPalette.ColorRole.WindowText, QColor(220, 220, 220))
        p.setColor(QPalette.ColorRole.Base, QColor(25, 25, 25))
        p.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
        p.setColor(QPalette.ColorRole.ToolTipBase, QColor(220, 220, 220))
        p.setColor(QPalette.ColorRole.ToolTipText, QColor(30, 30, 30))
        p.setColor(QPalette.ColorRole.Text, QColor(220, 220, 220))
        p.setColor(QPalette.ColorRole.Button, QColor(45, 45, 45))
        p.setColor(QPalette.ColorRole.ButtonText, QColor(220, 220, 220))
        p.setColor(QPalette.ColorRole.BrightText, QColor(255, 0, 0))
        self.setPalette(p)

    def bind_logs(self):
        log_signal = LogSignal()
        handler = QTextEditHandler(log_signal)
        formatter = logging.Formatter("%(asctime)s — %(message)s", "%H:%M:%S")
        handler.setFormatter(formatter)
        logging.basicConfig(level=logging.INFO, handlers=[handler])
        log_signal.message.connect(self.log.append)

    def load_mag(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Magento CSP Export", "", "Excel Files (*.xlsx *.xls)")
        if not path: return
        try:
            self.mag_df = pd.read_excel(path, dtype=str).fillna("")
            self.mag_cols = self.mag_df.columns.tolist()
            self.mag_label.setText(f"Magento: ✔ {os.path.basename(path)}")
            logging.info("Magento file loaded with %d rows and %d columns.", len(self.mag_df), len(self.mag_cols))
        except Exception as e:
            QMessageBox.critical(self, "Error Loading File", f"Could not load Magento file:\n{e}")
            logging.error("Failed to load Magento file: %s", e)

    def load_sup(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Supplier CSP Update", "", "Excel Files (*.xlsx *.xls)")
        if not path: return
        try:
            self.sup_df = pd.read_excel(path, dtype=str).fillna("")
            self.sup_label.setText(f"Supplier: ✔ {os.path.basename(path)}")
            sup_cols = self.sup_df.columns.tolist()
            for box in (self.sku_box, self.price_box, self.rebate_box):
                box.clear()
                box.addItems([""] + sup_cols)
                box.setEnabled(True)
            logging.info("Supplier file loaded with %d rows and %d columns.", len(self.sup_df), len(sup_cols))
        except Exception as e:
            QMessageBox.critical(self, "Error Loading File", f"Could not load Supplier file:\n{e}")
            logging.error("Failed to load Supplier file: %s", e)

    @staticmethod
    def norm_rebate(series):
        return pd.to_numeric(series.astype(str).str.replace("%", "", regex=False), errors='coerce').fillna(0)

    def run_validation(self):
        if self.mag_df is None or self.sup_df is None:
            QMessageBox.critical(self, "Error", "Please load both the Magento and Supplier files first.")
            return

        mappings = {
            "sku": self.sku_box.currentText(),
            "price": self.price_box.currentText(),
            "rebate": self.rebate_box.currentText()
        }

        if any(not value for value in mappings.values()):
            QMessageBox.critical(self, "Error", "Please map all three supplier columns (SKU, Price, Rebate).")
            return

        try:
            self.validate(mappings)
        except Exception as e:
            logging.exception("A critical error occurred during validation: %s", e)
            QMessageBox.critical(self, "Validation Crash", f"An unexpected error occurred:\n{e}")

    def validate(self, mappings):
        logging.info("Starting validation process...")
        mag, sup = self.mag_df.copy(), self.sup_df.copy()

        mag["__SKU__"] = mag["SupplierPartnumber"].astype(str).str.strip().str.upper()
        sup["__SKU__"] = sup[mappings["sku"]].astype(str).str.strip().str.upper()
        mag = mag[mag["__SKU__"] != ""]
        sup = sup[sup["__SKU__"] != ""]

        mag["__PRICE__"] = pd.to_numeric(mag["fixedprice"], errors='coerce').fillna(0)
        mag["__REBATE__"] = self.norm_rebate(mag["Rebate"]).round(2)
        sup["__PRICE__"] = pd.to_numeric(sup[mappings["price"]], errors='coerce').fillna(0)
        sup["__REBATE__"] = self.norm_rebate(sup[mappings["rebate"]])
        sup["__REBATE__"] = sup["__REBATE__"].apply(lambda x: x * 100 if 0 < abs(x) < 1 else x).round(2)

        mag_skus = set(mag["__SKU__"])
        sup_skus = set(sup["__SKU__"])

        disabled_skus = mag_skus - sup_skus
        disabled = mag[mag["__SKU__"].isin(disabled_skus)]

        new_skus = sup_skus - mag_skus
        new = sup[sup["__SKU__"].isin(new_skus)]

        common_skus = mag_skus & sup_skus
        mag_common = mag[mag["__SKU__"].isin(common_skus)]
        sup_common = sup[sup["__SKU__"].isin(common_skus)]

        if not mag_common.empty:
            sup_sub = sup_common[["__SKU__", "__PRICE__", "__REBATE__"]].rename(
                columns={"__PRICE__": "Updated Fixed Price", "__REBATE__": "Updated Rebate"}
            )
            merged = mag_common.merge(sup_sub, on="__SKU__", how="left")
            merged["__IS_SAME__"] = (
                merged["__PRICE__"].round(2) == merged["Updated Fixed Price"].round(2)
            ) & (
                merged["__REBATE__"].round(2) == merged["Updated Rebate"].round(2)
            )
            updated = merged[~merged["__IS_SAME__"]].copy()
            no_change = merged[merged["__IS_SAME__"]].copy()
            updated_skus = set(updated["__SKU__"])
            no_change = no_change[~no_change["__SKU__"].isin(updated_skus)]
            no_change = no_change[self.mag_cols]
            updated = updated[self.mag_cols + ["Updated Fixed Price", "Updated Rebate"]]
        else:
            no_change = pd.DataFrame(columns=self.mag_cols)
            updated = pd.DataFrame(columns=self.mag_cols + ["Updated Fixed Price", "Updated Rebate"])

        disabled = disabled[self.mag_cols]
        new = new[self.sup_df.columns]

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(os.path.expanduser("~"), "Documents", f"CSP_Validation_{ts}.xlsx")
        os.makedirs(os.path.dirname(out_path), exist_ok=True)

        summary_df = pd.DataFrame({
            "Category": ["Disabled", "New", "No Change", "Updated Existing"],
            "Count": [len(disabled), len(new), len(no_change), len(updated)]
        })

        with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
            summary_df.to_excel(writer, "Summary", index=False)
            no_change.to_excel(writer, "No Change", index=False)
            disabled.to_excel(writer, "Disabled", index=False)
            updated.to_excel(writer, "Updated Existing", index=False)
            new.to_excel(writer, "New", index=False)

        logging.info("Validation complete! Output file created.")
        QMessageBox.information(self, "Done", f"Validation finished successfully.\n\nOutput saved to:\n{out_path}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = CSPValidator()
    win.show()
    sys.exit(app.exec())
