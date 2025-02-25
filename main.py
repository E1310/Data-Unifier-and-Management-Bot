import os
import sys
import csv
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QHBoxLayout,
    QComboBox, QDoubleSpinBox, QMessageBox, QCompleter, QScrollArea, QDialog,
    QDialogButtonBox, QSizePolicy
)
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtCore import QSize, Qt

def resource_path(relative_path):
    """Get absolute path to resource, works for development and for PyInstaller bundle."""
    try:
        # PyInstaller creates a temporary folder and stores the path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class AttributeMapper(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("D.U.M.B - Data Unifier & Management Bot")
        self.setGeometry(100, 100, 1000, 700)

        # ------------------ APPLICATION ICON ------------------
        self.setWindowIcon(QIcon(resource_path("assets/favicon.png")))
        
        # Main layout
        self.layout = QVBoxLayout(self)

        # ------------------ TOP-RIGHT INFO BUTTON ------------------
        header_layout = QHBoxLayout()
        header_spacer = QWidget()
        header_spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        header_layout.addWidget(header_spacer)
        
        self.info_button = QPushButton("i")
        self.info_button.setFixedSize(30, 30)
        self.info_button.setStyleSheet("border-radius: 15px; background-color: #0078d7; color: white; font-weight: bold;")
        self.info_button.clicked.connect(self.show_info)
        header_layout.addWidget(self.info_button)
        
        self.layout.addLayout(header_layout)

        # ------------------ STEP 1: LOAD PRODUCT VARIANT FILE ------------------
        self.add_segment_header(
            "Step 1: Load Product Variant File", 
            "Make sure to export from the product variant view in Odoo (not the product view).\nFilter with Variant Values is set. \nWhen exporting, select 'I want to update data' you will need the Product template ID, Name, & Variant Values",
            "assets/help_variant.png"
        )
        file_variant_layout = QHBoxLayout()
        self.load_variant_file_btn = QPushButton("Load Product Variant File")
        self.load_variant_file_btn.clicked.connect(self.load_variant_file)
        file_variant_layout.addWidget(self.load_variant_file_btn)
        self.variant_file_label = QLabel("No file loaded")
        file_variant_layout.addWidget(self.variant_file_label)
        self.layout.addLayout(file_variant_layout)
        
        # ------------------ STEP 2: LOAD COMPONENTS FILE ------------------
        self.add_segment_header(
            "Step 2: Load Components File", 
            "Make sure to export from the product variant view in Odoo (not the product view).\nFilter out the items you don't need (e.g. Raw material category, no finished products ...). \nWhen exporting, select 'I want to update data' you will need the Name, & Unit of Measure",
            "assets/help_components.png"
        )
        file_components_layout = QHBoxLayout()
        self.load_product_file_btn = QPushButton("Load Components File")
        self.load_product_file_btn.clicked.connect(self.load_product_file)
        file_components_layout.addWidget(self.load_product_file_btn)
        self.product_file_label = QLabel("No file loaded")
        file_components_layout.addWidget(self.product_file_label)
        self.layout.addLayout(file_components_layout)
        
        # ------------------ STEP 3: SELECT PRODUCT TEMPLATE ------------------
        self.add_segment_header(
            "Step 3: Select Product Template", 
            "You can only proccess one product template at a time. \nSelect the product template to assign components to.",
            "assets/help_template.png"
        )
        self.product_group_combo = QComboBox()
        self.product_group_combo.currentIndexChanged.connect(self.populate_attributes)
        self.layout.addWidget(self.product_group_combo)
        
        # ------------------ STEP 4: ASSIGN COMPONENTS TO YOUR RECIPE ------------------
        self.add_segment_header(
            "Step 4: Assign Components to Your Recipe", 
            "For each attribute value, assign one or more components for the BoM.",
            "assets/help_assign.png"
        )
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_area.setWidget(self.scroll_content)
        self.layout.addWidget(self.scroll_area)
        
        # ------------------ STEP 5: GENERATE CSV FILE ------------------
        self.add_segment_header(
            "Step 5: Generate CSV File", 
            "Generate the CSV file to import BoMs into Odoo. Select the BoM type and the quantity of products to be produced.",
            "assets/help_generate.png"
        )
        self.generate_btn = QPushButton("Generate Output")
        self.generate_btn.clicked.connect(self.open_generate_dialog)
        self.layout.addWidget(self.generate_btn)
        
        # ------------------ DATA & MAPPING STRUCTURES ------------------
        self.variant_data = None
        self.product_data = None
        
        # For quick product lookups: product_name -> { 'id': ..., 'uom_id': ... }
        self.product_info_map = {}
        
        # Expected columns in the Excel files
        self.variant_column = "product_template_variant_value_ids"
        self.product_group_column = "product_tmpl_id/id"
        self.product_group_name_column = "product_tmpl_id/name"
        self.variant_id_column = "id"  # The variant's product ID in Odoo
        
        self.product_column = "name"
        self.product_id_column = "id"
        self.uom_column = "uom_id"
        
        # Holds attribute names and their possible values
        self.attributes = {}
        
        # For storing user-assigned components per (attribute, value)
        # Key: (attribute, value)
        # Value: list of dicts { 'combo': QComboBox, 'qty': QDoubleSpinBox, 'uom_label': QLabel, 'widget': QWidget }
        self.components_map = {}

    # ------------------ SEGMENT HEADERS & HELP ------------------
    
    def add_segment_header(self, title, help_text, help_image):
        """
        Adds a horizontal layout containing:
          - The step title
          - A "?" help button immediately next to it
        """
        header_layout = QHBoxLayout()
        
        # Layout for label + "?"
        label_button_layout = QHBoxLayout()
        title_label = QLabel(title)
        title_label.setStyleSheet("font-weight: bold; font-size: 16px;")
        label_button_layout.addWidget(title_label)
        
        help_btn = QPushButton("?")
        help_btn.setFixedSize(25, 25)
        help_btn.setStyleSheet("border-radius: 12px; background-color: #0078d7; color: white; font-weight: bold;")
        help_btn.clicked.connect(lambda: self.show_help_popup(help_text, help_image))
        label_button_layout.addWidget(help_btn)
        
        # Put label+button layout in the main header layout
        header_layout.addLayout(label_button_layout)
        header_layout.addStretch()
        
        self.layout.addLayout(header_layout)

    def show_help_popup(self, help_text, help_image):
        dialog = QDialog(self)
        dialog.setWindowTitle("Help")
        d_layout = QVBoxLayout(dialog)
        
        text_label = QLabel(help_text)
        text_label.setWordWrap(True)
        d_layout.addWidget(text_label)
        
        pixmap = QPixmap(resource_path(help_image))
        if not pixmap.isNull():
            image_label = QLabel()
            image_label.setPixmap(pixmap)
            d_layout.addWidget(image_label)
        
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        button_box.accepted.connect(dialog.accept)
        d_layout.addWidget(button_box)
        
        dialog.exec()

    def show_info(self):
        """
        Shows an About dialog with author, version, and open-source info.
        """
        info_text = (
            "D.U.M.B - Data Unifier & Management Bot\n"
            "Author: Daher Zaidan\n"
            "Version: 1.0\n"
            "Open Source\n\n"
            "This software is open source. Feel free to contribute or modify."
        )
        QMessageBox.information(self, "About", info_text)

    # ------------------ FILE LOADING ------------------
    
    def load_variant_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Product Variant File", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.variant_data = pd.read_excel(file_path)
            self.variant_file_label.setText(file_path.split("/")[-1])
            # Validate columns
            if self.variant_column not in self.variant_data.columns:
                self.show_error(f"Column '{self.variant_column}' not found in the variant file")
                self.variant_data = None
                return
            if self.variant_id_column not in self.variant_data.columns:
                self.show_error(f"Column '{self.variant_id_column}' (Variant ID) not found in the variant file")
                self.variant_data = None
                return
            self.populate_group_dropdown()
    
    def load_product_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Components File", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.product_data = pd.read_excel(file_path)
            self.product_file_label.setText(file_path.split("/")[-1])
            # Build product info map
            self.product_info_map.clear()
            for idx, row in self.product_data.iterrows():
                product_name = str(row.get(self.product_column, ""))
                product_id = str(row.get(self.product_id_column, ""))
                uom = str(row.get(self.uom_column, ""))
                self.product_info_map[product_name] = {'id': product_id, 'uom_id': uom}
    
    def refresh_files(self):
        self.load_variant_file()
        self.load_product_file()
    
    # ------------------ GROUP & ATTRIBUTES ------------------
    
    def populate_group_dropdown(self):
        if self.variant_data is None:
            return
        self.product_group_combo.clear()
        groups = self.variant_data[[self.product_group_column, self.product_group_name_column]].drop_duplicates()
        for _, row in groups.iterrows():
            self.product_group_combo.addItem(str(row[self.product_group_name_column]), row[self.product_group_column])
    
    def populate_attributes(self):
        """
        Displays attributes in the scroll area. Each (attribute, value) pair can have multiple
        component lines assigned.
        """
        if self.variant_data is None:
            return

        # Clear existing items
        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        selected_group = self.product_group_combo.currentData()
        if not selected_group:
            return

        group_variants = self.variant_data[self.variant_data[self.product_group_column] == selected_group]
        self.attributes.clear()
        self.components_map.clear()

        for _, row in group_variants.iterrows():
            raw_variants = str(row[self.variant_column]).split(",")
            for item in raw_variants:
                item = item.strip()
                if not item:
                    continue
                if ": " in item:
                    attribute, value = item.split(": ", 1)
                else:
                    attribute, value = "Unknown", item
                if attribute not in self.attributes:
                    self.attributes[attribute] = []
                if value not in self.attributes[attribute]:
                    self.attributes[attribute].append(value)

        # Build the UI for each attribute/value
        for attribute in sorted(self.attributes.keys()):
            attribute_label = QLabel(attribute)
            attribute_label.setStyleSheet("font-weight: bold; font-size: 14px;")
            self.scroll_layout.addWidget(attribute_label)

            for value in self.attributes[attribute]:
                variant_widget = QWidget()
                variant_layout = QVBoxLayout(variant_widget)

                value_label = QLabel(f"- {value}")
                variant_layout.addWidget(value_label)

                components_container = QVBoxLayout()
                variant_layout.addLayout(components_container)

                # Track components for (attribute, value)
                self.components_map.setdefault((attribute, value), [])

                add_component_btn = QPushButton("Add Component")
                add_component_btn.clicked.connect(
                    lambda _, c=components_container, attr=attribute, val=value: self.add_component_line(c, attr, val)
                )
                variant_layout.addWidget(add_component_btn)

                self.scroll_layout.addWidget(variant_widget)
    
    # ------------------ ADD COMPONENT LINE ------------------
    
    def add_component_line(self, container_layout, attribute, value):
        """
        Adds a new row with:
          - Product combo box
          - Quantity spin box
          - Read-only UoM label
          - Remove button
        """
        component_widget = QWidget()
        layout = QHBoxLayout(component_widget)

        # Product combo
        combo = QComboBox()
        combo.setEditable(True)
        if self.product_data is not None:
            products = self.product_data[self.product_column].dropna().astype(str).tolist()
            combo.addItems(products)
            completer = QCompleter(products)
            completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
            combo.setCompleter(completer)
        layout.addWidget(combo)

        # Quantity
        qty_input = QDoubleSpinBox()
        qty_input.setDecimals(2)
        qty_input.setRange(0.01, 10000)
        qty_input.setValue(1.00)
        layout.addWidget(qty_input)

        # UoM label
        uom_label = QLabel("")
        uom_label.setStyleSheet("background-color: transparent; padding: 2px;")
        uom_label.setMinimumWidth(50)
        layout.addWidget(uom_label)

        # Update UoM whenever the combo changes
        combo.currentIndexChanged.connect(lambda idx, c=combo, l=uom_label: self.update_uom_label(c, l))
        if combo.count() > 0:
            self.update_uom_label(combo, uom_label)

        # Remove button
        remove_btn = QPushButton("X")
        remove_btn.setFixedSize(QSize(30, 30))
        remove_btn.clicked.connect(lambda _, w=component_widget: w.setParent(None))
        layout.addWidget(remove_btn)

        container_layout.addWidget(component_widget)

        self.components_map[(attribute, value)].append({
            'combo': combo,
            'qty': qty_input,
            'uom_label': uom_label,
            'widget': component_widget
        })

    def update_uom_label(self, combo, label):
        selected_product = combo.currentText()
        info = self.product_info_map.get(selected_product, {})
        label.setText(info.get('uom_id', ""))

    # ------------------ GENERATE OUTPUT ------------------
    
    def open_generate_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Generate BoM Output")
        layout = QVBoxLayout(dialog)

        qty_label = QLabel("Qty to be produced:")
        layout.addWidget(qty_label)
        qty_spin = QDoubleSpinBox()
        qty_spin.setDecimals(2)
        qty_spin.setRange(0.01, 100000)
        qty_spin.setValue(1.0)
        layout.addWidget(qty_spin)

        bom_type_label = QLabel("BoM Type:")
        layout.addWidget(bom_type_label)
        bom_type_combo = QComboBox()
        bom_type_combo.addItem("Manufacture this product")
        bom_type_combo.addItem("Kit")
        layout.addWidget(bom_type_combo)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        layout.addWidget(buttons)

        def on_ok():
            selected_qty = qty_spin.value()
            selected_bom_type = bom_type_combo.currentText()
            dialog.accept()
            self.generate_csv(selected_qty, selected_bom_type)

        def on_cancel():
            dialog.reject()

        buttons.accepted.connect(on_ok)
        buttons.rejected.connect(on_cancel)
        dialog.exec()

    def generate_csv(self, product_qty, bom_type):
        """
        For each variant (in the selected product group), aggregate all components and
        write a CSV file. The first component row includes BOM header info; subsequent
        component rows have empty header columns.
        """
        if self.variant_data is None:
            self.show_error("No variant data loaded.")
            return
        selected_group = self.product_group_combo.currentData()
        if not selected_group:
            self.show_error("No product group selected.")
            return

        file_path, _ = QFileDialog.getSaveFileName(self, "Save BoM CSV", "", "CSV Files (*.csv)")
        if not file_path:
            return

        group_variants = self.variant_data[self.variant_data[self.product_group_column] == selected_group]
        fieldnames = [
            "product_tmpl_id/id",
            "product_id/id",
            "type",
            "product_qty",
            "bom_line_ids/product_id/id",
            "bom_line_ids/product_qty",
            "bom_line_ids/product_uom_id"
        ]

        try:
            with open(file_path, mode="w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()

                for _, row in group_variants.iterrows():
                    tmpl_id = str(row.get(self.product_group_column, ""))
                    variant_id = str(row.get(self.variant_id_column, ""))

                    # Parse attribute-value pairs
                    raw_variants = str(row[self.variant_column]).split(",")
                    av_pairs = []
                    for item in raw_variants:
                        item = item.strip()
                        if not item:
                            continue
                        if ": " in item:
                            attr, val = item.split(": ", 1)
                        else:
                            attr, val = "Unknown", item
                        av_pairs.append((attr, val))

                    # Gather all components
                    all_components = []
                    for (attr, val) in av_pairs:
                        if (attr, val) in self.components_map:
                            all_components.extend(self.components_map[(attr, val)])
                    
                    if not all_components:
                        continue

                    first_line = True
                    for comp in all_components:
                        selected_product = comp['combo'].currentText()
                        component_qty = comp['qty'].value()
                        product_info = self.product_info_map.get(selected_product, {})
                        comp_product_id = product_info.get('id', "")
                        comp_uom_id = product_info.get('uom_id', "")

                        if first_line:
                            writer.writerow({
                                "product_tmpl_id/id": tmpl_id,
                                "product_id/id": variant_id,
                                "type": bom_type,
                                "product_qty": product_qty,
                                "bom_line_ids/product_id/id": comp_product_id,
                                "bom_line_ids/product_qty": component_qty,
                                "bom_line_ids/product_uom_id": comp_uom_id
                            })
                            first_line = False
                        else:
                            writer.writerow({
                                "product_tmpl_id/id": "",
                                "product_id/id": "",
                                "type": "",
                                "product_qty": "",
                                "bom_line_ids/product_id/id": comp_product_id,
                                "bom_line_ids/product_qty": component_qty,
                                "bom_line_ids/product_uom_id": comp_uom_id
                            })

            QMessageBox.information(self, "Success", f"CSV file generated:\n{file_path}")
        except Exception as e:
            self.show_error(f"Error writing CSV: {e}")

    # ------------------ ERROR POPUP ------------------
    
    def show_error(self, message):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Critical)
        msg.setWindowTitle("Error")
        msg.setText(message)
        msg.exec()

# ------------------ MAIN LAUNCH ------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AttributeMapper()
    window.show()
    sys.exit(app.exec())

# ------------------ END or FIN ------------------