import sys, os, logging, pandas as pd
from datetime import datetime
from PyQt6.QtWidgets import (QApplication, QFileDialog, QMainWindow, QWidget, QLabel,
                             QPushButton, QComboBox, QTextEdit, QVBoxLayout, QHBoxLayout,
                             QMessageBox, QGridLayout)
from PyQt6.QtCore import pyqtSignal, QObject
from PyQt6.QtGui import QPalette, QColor


class LogSignal(QObject):
    """Signal for logging messages to the UI."""
    message = pyqtSignal(str)


class QTextEditHandler(logging.Handler):
    """A logging handler that emits a PyQt signal."""

    def __init__(self, sig):
        super().__init__()
        self.sig = sig

    def emit(self, rec):
        self.sig.message.emit(self.format(rec))


class CSPValidator(QMainWindow):
    """Main application window for CSP validation and Magento file generation."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("CSP Validator & Magento Builder — Zero Excuses")
        self.resize(820, 620)  # Increased height for new controls
        self.mag_df = self.sup_df = None
        self.disabled_df = self.updated_df = self.new_df = self.no_change_df = None
        self.mag_cols = []
        self.build_ui()
        self.dark_mode()
        self.bind_logs()

    def build_ui(self):
        """Constructs the user interface."""
        central = QWidget()
        self.setCentralWidget(central)
        self.mag_label = QLabel("Magento File: ✖ Not Loaded")
        self.sup_label = QLabel("Supplier File: ✖ Not Loaded")
        load_mag = QPushButton("Load Magento File")
        load_sup = QPushButton("Load Supplier File")
        load_mag.clicked.connect(self.load_mag)
        load_sup.clicked.connect(self.load_sup)

        # Combo boxes for mapping supplier file columns
        self.sku_box = QComboBox()
        self.price_box = QComboBox()
        self.rebate_box = QComboBox()
        self.group_box = QComboBox()  # New: Customer Group mapping

        for box in (self.sku_box, self.price_box, self.rebate_box, self.group_box):
            box.setEnabled(False)
            box.setToolTip("Map this to the corresponding column in the Supplier File.")

        run_btn = QPushButton("Run Validation")
        run_btn.clicked.connect(self.run_validation)

        # New: Button to build Magento upload files
        self.build_btn = QPushButton("Generate Magento Upload Files")
        self.build_btn.clicked.connect(self.build_upload_files)
        self.build_btn.setEnabled(False)
        self.build_btn.setToolTip("Run validation first to enable this.")

        about_btn = QPushButton("ℹ️ About / How to Use")
        about_btn.clicked.connect(self.show_about)

        self.log = QTextEdit()
        self.log.setReadOnly(True)

        # Main layout
        lay = QVBoxLayout(central)
        file_grid = QVBoxLayout()
        for btn, lbl in ((load_mag, self.mag_label), (load_sup, self.sup_label)):
            row = QHBoxLayout()
            row.addWidget(btn)
            row.addWidget(lbl, 1)
            file_grid.addLayout(row)
        lay.addLayout(file_grid)

        # Grid layout for cleaner mapping controls
        map_layout = QGridLayout()
        map_layout.addWidget(QLabel("Supplier SKU Column:"), 0, 0)
        map_layout.addWidget(self.sku_box, 0, 1)
        map_layout.addWidget(QLabel("Supplier Price Column:"), 0, 2)
        map_layout.addWidget(self.price_box, 0, 3)
        map_layout.addWidget(QLabel("Supplier Rebate Column:"), 1, 0)
        map_layout.addWidget(self.rebate_box, 1, 1)
        map_layout.addWidget(QLabel("Supplier Customer Group Column:"), 1, 2)
        map_layout.addWidget(self.group_box, 1, 3)
        lay.addLayout(map_layout)

        lay.addWidget(about_btn)

        # Action buttons in a horizontal row
        action_row = QHBoxLayout()
        action_row.addWidget(run_btn, 1)
        action_row.addWidget(self.build_btn, 1)
        lay.addLayout(action_row)

        lay.addWidget(QLabel("Log:"))
        lay.addWidget(self.log, 1)  # Allow log to expand

    def show_about(self):
        """Displays a help message box."""
        text = (
            "🐾 **CSP Validator & Magento Builder — How to Use** 🐾\n\n"
            "This tool compares Customer Specific Pricing (CSP) and builds upload files for Magento.\n\n"
            "**Part 1: Validation**\n\n"
            "📂 Step 1: Load Files\n"
            "    - Click 'Load Magento File' and select your current CSP export from Magento.\n"
            "    - Click 'Load Supplier File' and select the new/updated CSP file from the supplier.\n\n"
            "🧩 Step 2: Map Columns\n"
            "    - For the supplier file, use the dropdowns to match their column names to our data points (SKU, Price, Rebate, Customer Group).\n\n"
            "🚀 Step 3: Run Validation\n"
            "    - Click 'Run Validation'.\n"
            "    - The tool generates an Excel file categorizing all changes (No Change, Disabled, New, Updated).\n\n"
            "**Part 2: Magento File Generation**\n\n"
            "🔨 Step 4: Generate Upload Files\n"
            "    - After a successful validation, the 'Generate Magento Upload Files' button becomes active.\n"
            "    - Click it to create a new Excel file containing three sheets, formatted for Magento import:\n"
            "        - `Disabled_Profile`: Deactivates old CSPs.\n"
            "        - `Updated_Profile`: Updates existing CSPs with new prices/rebates.\n"
            "        - `New_Profile`: Creates new CSPs.\n\n"
            "📄 Output:\n"
            "    - All generated files are saved in your computer's Documents folder."
        )
        QMessageBox.information(self, "About / How to Use", text)

    def dark_mode(self):
        """Applies a dark color scheme to the application."""
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
        """Redirects Python's logging output to the text box in the UI."""
        log_signal = LogSignal()
        handler = QTextEditHandler(log_signal)
        formatter = logging.Formatter("%(asctime)s — %(message)s", "%H:%M:%S")
        handler.setFormatter(formatter)
        logging.basicConfig(level=logging.INFO, handlers=[handler])
        log_signal.message.connect(self.log.append)

    def load_mag(self):
        """Loads the Magento CSP export file."""
        path, _ = QFileDialog.getOpenFileName(self, "Select Magento CSP Export", "", "Excel Files (*.xlsx *.xls)")
        if not path: return
        try:
            self.mag_df = pd.read_excel(path, dtype=str).fillna("")
            required_mag_cols = {"SupplierPartnumber", "fixedprice", "Rebate", "CustomerGroup"}
            if not required_mag_cols.issubset(self.mag_df.columns):
                missing = required_mag_cols - set(self.mag_df.columns)
                raise ValueError(f"Magento file is missing required columns: {', '.join(missing)}")
            self.mag_cols = self.mag_df.columns.tolist()
            self.mag_label.setText(f"Magento: ✔ {os.path.basename(path)}")
            logging.info("Magento file loaded with %d rows and %d columns.", len(self.mag_df), len(self.mag_cols))
        except Exception as e:
            QMessageBox.critical(self, "Error Loading File", f"Could not load Magento file:\n{e}")
            logging.error("Failed to load Magento file: %s", e)
            self.mag_df = None
            self.mag_label.setText("Magento File: ✖ Not Loaded")

    def load_sup(self):
        """Loads the Supplier CSP update file."""
        path, _ = QFileDialog.getOpenFileName(self, "Select Supplier CSP Update", "", "Excel Files (*.xlsx *.xls)")
        if not path: return
        try:
            self.sup_df = pd.read_excel(path, dtype=str).fillna("")
            self.sup_label.setText(f"Supplier: ✔ {os.path.basename(path)}")
            sup_cols = self.sup_df.columns.tolist()
            for box in (self.sku_box, self.price_box, self.rebate_box, self.group_box):
                box.clear()
                box.addItems([""] + sup_cols)
                box.setEnabled(True)
            logging.info("Supplier file loaded with %d rows and %d columns.", len(self.sup_df), len(sup_cols))
        except Exception as e:
            QMessageBox.critical(self, "Error Loading File", f"Could not load Supplier file:\n{e}")
            logging.error("Failed to load Supplier file: %s", e)
            self.sup_df = None
            self.sup_label.setText("Supplier File: ✖ Not Loaded")

    @staticmethod
    def norm_rebate(series):
        """Converts a rebate column to a numeric type, handling percentages."""
        numeric_series = pd.to_numeric(series.astype(str).str.replace("%", "", regex=False), errors='coerce').fillna(0)
        return numeric_series

    def run_validation(self):
        """Entry point for the validation process, triggered by a button click."""
        if self.mag_df is None or self.sup_df is None:
            QMessageBox.critical(self, "Error", "Please load both the Magento and Supplier files first.")
            return

        # Reset from any previous run
        self.build_btn.setEnabled(False)
        self.disabled_df = self.updated_df = self.new_df = self.no_change_df = None

        mappings = {
            "sku": self.sku_box.currentText(),
            "price": self.price_box.currentText(),
            "rebate": self.rebate_box.currentText(),
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
        """Performs the core logic of comparing the two files."""
        logging.info("Starting validation process...")
        mag, sup = self.mag_df.copy(), self.sup_df.copy()

        # --- Data Cleaning and Preparation ---
        mag["__SKU__"] = mag["SupplierPartnumber"].astype(str).str.strip().str.upper()
        sup["__SKU__"] = sup[mappings["sku"]].astype(str).str.strip().str.upper()
        mag = mag[mag["__SKU__"] != ""].dropna(subset=['__SKU__'])
        sup = sup[sup["__SKU__"] != ""].dropna(subset=['__SKU__'])

        mag["__PRICE__"] = pd.to_numeric(mag["fixedprice"], errors='coerce').fillna(0).round(2)
        mag["__REBATE__"] = self.norm_rebate(mag["Rebate"])
        sup["__PRICE__"] = pd.to_numeric(sup[mappings["price"]], errors='coerce').fillna(0).round(2)
        sup["__REBATE__"] = self.norm_rebate(sup[mappings["rebate"]])
        sup["__REBATE__"] = sup["__REBATE__"].apply(lambda x: x * 100 if 0 < abs(x) < 1 else x)

        # --- Categorization Logic ---
        mag_skus = set(mag["__SKU__"])
        sup_skus = set(sup["__SKU__"])

        self.disabled_df = mag[mag["__SKU__"].isin(mag_skus - sup_skus)][self.mag_cols].copy()
        self.new_df = sup[sup["__SKU__"].isin(sup_skus - mag_skus)][self.sup_df.columns].copy()

        common_skus = mag_skus & sup_skus
        mag_common = mag[mag["__SKU__"].isin(common_skus)]
        sup_common = sup[sup["__SKU__"].isin(common_skus)]

        self.no_change_df = pd.DataFrame(columns=self.mag_cols)
        self.updated_df = pd.DataFrame(columns=self.mag_cols + ["Updated Fixed Price", "Updated Rebate"])

        if not mag_common.empty:
            sup_sub = sup_common[["__SKU__", "__PRICE__", "__REBATE__"]].rename(
                columns={"__PRICE__": "Updated Fixed Price", "__REBATE__": "Updated Rebate"}
            )
            merged = mag_common.merge(sup_sub, on="__SKU__", how="left")

            price_is_same = (merged["__PRICE__"] == merged["Updated Fixed Price"])
            rebate_is_same = (merged["__REBATE__"].round(2) == merged["Updated Rebate"].round(2))
            merged["__IS_SAME__"] = price_is_same & rebate_is_same

            self.updated_df = merged[~merged["__IS_SAME__"]][
                self.mag_cols + ["Updated Fixed Price", "Updated Rebate"]].copy()
            self.no_change_df = merged[merged["__IS_SAME__"]][self.mag_cols].copy()

        # --- Save Validation Report ---
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(os.path.expanduser("~"), "Documents", f"CSP_Validation_{ts}.xlsx")
        os.makedirs(os.path.dirname(out_path), exist_ok=True)

        summary_df = pd.DataFrame({
            "Category": ["Disabled", "New", "No Change", "Updated Existing"],
            "Count": [len(self.disabled_df), len(self.new_df), len(self.no_change_df), len(self.updated_df)]
        })

        with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
            summary_df.to_excel(writer, "Summary", index=False)
            self.no_change_df.to_excel(writer, "No Change", index=False)
            self.disabled_df.to_excel(writer, "Disabled", index=False)
            self.updated_df.to_excel(writer, "Updated Existing", index=False)
            self.new_df.to_excel(writer, "New", index=False)

        logging.info("Validation complete! Output file created.")
        self.build_btn.setEnabled(True)  # Enable the build button
        QMessageBox.information(self, "Done", f"Validation finished successfully.\n\n"
                                              f"You can now generate the Magento upload files.\n\n"
                                              f"Output saved to:\n{out_path}")

    def build_upload_files(self):
        """Generates the Magento-ready upload file with Disabled, Updated, and New profiles."""
        if self.disabled_df is None:  # Check if validation has been run
            QMessageBox.critical(self, "Error", "Please run the validation process before generating upload files.")
            return

        group_mapping = self.group_box.currentText()
        if not group_mapping and not self.new_df.empty:
            QMessageBox.critical(self, "Mapping Missing",
                                 "Please map the 'Supplier Customer Group' column to build the 'New Profile' sheet.")
            return

        logging.info("Starting Magento upload file build...")
        try:
            current_date = datetime.now().strftime("%m/%d/%Y")

            # --- 1. Disabled Profile Build ---
            disabled_upload = pd.DataFrame()
            if not self.disabled_df.empty:
                df = self.disabled_df
                disabled_upload = pd.DataFrame({
                    "Part Number": df["SupplierPartnumber"], "Customer Group": df["CustomerGroup"],
                    "Start Date": current_date, "End Date": "",
                    "Fixed Price": pd.to_numeric(df["fixedprice"], errors='coerce').fillna(0).round(2),
                    "Active": "No", "Contract": "Y", "Rebate": df["Rebate"]
                })

            # --- 2. Updated Existing Profile Build ---
            updated_upload = pd.DataFrame()
            if not self.updated_df.empty:
                df = self.updated_df
                updated_upload = pd.DataFrame({
                    "Part Number": df["SupplierPartnumber"], "Customer Group": df["CustomerGroup"],
                    "Start Date": current_date, "End Date": "",
                    "Fixed Price": pd.to_numeric(df["Updated Fixed Price"], errors='coerce').fillna(0).round(2),
                    "Active": "Yes", "Contract": "Y", "Rebate": df["Updated Rebate"]
                })

            # --- 3. New Profile Build ---
            new_upload = pd.DataFrame()
            if not self.new_df.empty:
                df = self.new_df
                mappings = {"sku": self.sku_box.currentText(), "price": self.price_box.currentText(),
                            "rebate": self.rebate_box.currentText(), "group": self.group_box.currentText()}

                rebate_series = self.norm_rebate(df[mappings["rebate"]])
                rebate_series = rebate_series.apply(lambda x: x * 100 if 0 < abs(x) < 1 else x)
                rebate_values = rebate_series.round(2)

                new_upload = pd.DataFrame({
                    "Part Number": df[mappings["sku"]], "Customer Group": df[mappings["group"]],
                    "Start Date": current_date, "End Date": "",
                    "Fixed Price": pd.to_numeric(df[mappings["price"]], errors='coerce').fillna(0).round(2),
                    "Active": "Yes", "Contract": "Y", "Rebate": rebate_values
                })

            # --- Save Magento Upload File ---
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = os.path.join(os.path.expanduser("~"), "Documents", f"Magento_Upload_{ts}.xlsx")

            with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
                if not disabled_upload.empty:
                    disabled_upload.to_excel(writer, "Disabled_Profile", index=False)
                if not updated_upload.empty:
                    updated_upload.to_excel(writer, "Updated_Profile", index=False)
                if not new_upload.empty:
                    new_upload.to_excel(writer, "New_Profile", index=False)

            logging.info("Magento upload file created at %s", out_path)
            QMessageBox.information(self, "Build Complete",
                                    f"Magento upload file has been created successfully.\n\nOutput saved to:\n{out_path}")

        except KeyError as e:
            logging.error("A mapped column was not found: %s", e)
            QMessageBox.critical(self, "Mapping Error",
                                 f"A specified column could not be found in the supplier file: {e}.\nPlease check your column mappings.")
        except Exception as e:
            logging.exception("Failed to build Magento upload files: %s", e)
            QMessageBox.critical(self, "Build Crash", f"An unexpected error occurred during file generation:\n{e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = CSPValidator()
    win.show()
    sys.exit(app.exec())
