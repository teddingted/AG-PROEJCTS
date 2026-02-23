import sys
import os
import cv2
import numpy as np
import csv

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
    QFileDialog, QMessageBox, QGroupBox, QRadioButton, QSlider, QScrollArea,
    QTableWidget, QTableWidgetItem, QHeaderView, QLineEdit, QCheckBox, 
    QColorDialog, QToolButton, QSpinBox, QAbstractSpinBox, QFormLayout,
    QTreeWidget, QTreeWidgetItem, QInputDialog
)
from PySide6.QtCore import Qt, Slot
from PySide6.QtGui import QColor, QImage, QPixmap, QIcon

from core.project import Project
from core.series import Series
from gui.image_canvas import ImageCanvas
from core.processor import ImageProcessor

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("AutoPlotDigitizer V2 - Windows Port")
        self.setGeometry(100, 100, 1400, 900)
        
        self.project_model = Project()
        self.project_model.add_observer(self.on_project_updated)
        
        # Central Widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout()
        
        # 1. Left Panel (Controls)
        left_layout = QVBoxLayout()
        left_layout.setSpacing(10)
        left_layout.setAlignment(Qt.AlignTop)
        
        # 1.1 Load Image / Project
        gb_load = QGroupBox("1. Project / Image")
        load_layout = QVBoxLayout()
        
        hbox_proj = QHBoxLayout()
        btn_load_proj = QPushButton("Load Project")
        btn_load_proj.clicked.connect(self.load_project_ui)
        btn_save_proj = QPushButton("Save Project")
        btn_save_proj.clicked.connect(self.save_project_ui)
        hbox_proj.addWidget(btn_load_proj)
        hbox_proj.addWidget(btn_save_proj)
        
        btn_load_img = QPushButton("Load Image")
        btn_load_img.clicked.connect(self.load_image_ui)
        
        load_layout.addLayout(hbox_proj)
        load_layout.addWidget(btn_load_img)
        gb_load.setLayout(load_layout)
        left_layout.addWidget(gb_load)
        
        # 2. Calibration
        gb_calib = QGroupBox("2. Calibration")
        calib_layout = QVBoxLayout()
        
        self.btn_calib_mode = QPushButton("Set Calibration Points (0/4)")
        self.btn_calib_mode.setCheckable(True)
        self.btn_calib_mode.clicked.connect(self.toggle_calibration_mode)
        
        form_calib = QFormLayout()
        
        self.inp_x1 = QLineEdit("0.0")
        self.inp_x2 = QLineEdit("1.0")
        self.inp_y1 = QLineEdit("0.0")
        self.inp_y2 = QLineEdit("1.0")
        
        self.chk_log_x = QCheckBox("Log Scale X")
        self.chk_log_y = QCheckBox("Log Scale Y")
        
        self.btn_reset_calib = QPushButton("Reset Points")
        self.btn_reset_calib.clicked.connect(self.reset_calibration)
        
        form_calib.addRow("X1 Value:", self.inp_x1)
        form_calib.addRow("X2 Value:", self.inp_x2)
        form_calib.addRow("Y1 Value:", self.inp_y1)
        form_calib.addRow("Y2 Value:", self.inp_y2)
        form_calib.addRow(self.chk_log_x)
        form_calib.addRow(self.chk_log_y)
        form_calib.addRow(self.btn_reset_calib)
        
        calib_layout.addWidget(self.btn_calib_mode)
        calib_layout.addLayout(form_calib)
        gb_calib.setLayout(calib_layout)
        left_layout.addWidget(gb_calib)
        
        # 3. Masking & Color
        gb_mask = QGroupBox("3. Extraction Mask & Color")
        mask_layout = QVBoxLayout()
        mask_layout.setSpacing(2)
        mask_layout.setContentsMargins(5,2,5,2)
        mask_layout.setAlignment(Qt.AlignTop) 
        
        mask_row1 = QHBoxLayout()
        self.btn_mask_mode = QPushButton("Draw Mask (Pen)")
        self.btn_mask_mode.setCheckable(True)
        self.btn_mask_mode.clicked.connect(self.toggle_mask_mode)
        
        self.btn_color = QPushButton("Pick Target Color (Default: Dark)")
        self.btn_color.clicked.connect(self.pick_color)
        self.target_color_hsv = None 
        self.current_series_color = Qt.red 
        
        mask_row1.addWidget(self.btn_mask_mode)
        mask_row1.addWidget(self.btn_color)
        mask_layout.addLayout(mask_row1)
        
        # Line Mode (Solid vs Dotted)
        self.gb_line_mode = QGroupBox("Line Type")
        line_mode_layout = QVBoxLayout()
        line_mode_layout.setSpacing(5)
        line_mode_layout.setContentsMargins(5, 5, 5, 5)
        
        self.radio_solid = QRadioButton("Solid Line")
        self.radio_dotted = QRadioButton("Dotted/Dashed Line (Gap Fill)")
        self.radio_solid.setChecked(True)
        self.radio_dotted.toggled.connect(self.toggle_gap_slider)
        
        line_mode_layout.addWidget(self.radio_solid)
        
        layout_dotted_row = QHBoxLayout()
        layout_dotted_row.addWidget(self.radio_dotted)
        
        self.btn_help_dotted = QToolButton()
        self.btn_help_dotted.setText("?")
        self.btn_help_dotted.setFixedSize(20, 20)
        self.btn_help_dotted.setToolTip("Use for graphs with grid lines or dashed curves.\nAutomatically bridges gaps and filters small noise.\nUse 'Gap Fill' slider to adjust connection strength.")
        self.btn_help_dotted.setStyleSheet("border-radius: 10px; background-color: #ddd; font-weight: bold;")
        layout_dotted_row.addWidget(self.btn_help_dotted)
        layout_dotted_row.addStretch()
        
        line_mode_layout.addLayout(layout_dotted_row)
        
        self.container_gap = QWidget()
        gap_layout = QVBoxLayout(self.container_gap)
        gap_layout.setSpacing(2)
        gap_layout.setContentsMargins(0,0,0,0)
        
        self.lbl_gap = QLabel("Gap Fill: 3px")
        self.slider_gap = QSlider(Qt.Horizontal)
        self.slider_gap.setRange(1, 20)
        self.slider_gap.setValue(3)
        self.slider_gap.valueChanged.connect(self.update_gap_label)
        
        gap_layout.addWidget(self.lbl_gap)
        gap_layout.addWidget(self.slider_gap)
        
        line_mode_layout.addWidget(self.container_gap)
        self.container_gap.setVisible(False) 
        
        self.gb_line_mode.setLayout(line_mode_layout)
        mask_layout.addWidget(self.gb_line_mode)
        
        mask_btn_layout = QHBoxLayout()
        mask_btn_layout.setSpacing(5)
        self.btn_undo_mask = QPushButton("Undo Stroke")
        self.btn_undo_mask.clicked.connect(self.undo_mask)
        
        self.btn_mask_eraser = QPushButton("Eraser Mode")
        self.btn_mask_eraser.setCheckable(True)
        self.btn_mask_eraser.clicked.connect(self.toggle_mask_eraser)
        
        self.btn_clear_mask = QPushButton("Clear/Delete Mask")
        self.btn_clear_mask.clicked.connect(self.clear_mask)
        
        mask_btn_layout.addWidget(self.btn_undo_mask)
        mask_btn_layout.addWidget(self.btn_mask_eraser)
        mask_btn_layout.addWidget(self.btn_clear_mask)
        mask_layout.addLayout(mask_btn_layout)
        mask_layout.addStretch()
        
        gb_mask.setLayout(mask_layout)
        left_layout.addWidget(gb_mask)
        
        # 4. Extract
        gb_extract = QGroupBox("4. Add Series")
        extract_layout = QVBoxLayout()
        extract_layout.setSpacing(2)
        extract_layout.setContentsMargins(5,2,5,2)
        extract_layout.setAlignment(Qt.AlignTop)
        self.btn_extract = QPushButton("Extract & Add Series")
        self.btn_extract.clicked.connect(self.extract_data)
        
        self.inp_series_name = QLineEdit("Series 1")
        
        extract_layout.addWidget(QLabel("Series Name:"))
        extract_layout.addWidget(self.inp_series_name)
        extract_layout.addWidget(self.btn_extract)
        
        # Sampling Control
        gb_sampling = QGroupBox("Sampling")
        sampling_layout = QVBoxLayout()
        
        self.radio_raw = QRadioButton("All Points (Raw)")
        self.radio_fixed = QRadioButton("Fixed Count (2~2000)")
        self.radio_key = QRadioButton("Key Points (Douglas-Peucker)")
        self.radio_raw.setChecked(True)
        
        class WarningSpinBox(QSpinBox):
            def validate(self, text, pos):
                return super().validate(text, pos)

            def fixup(self, text):
                try:
                    val = int(text)
                    if val < self.minimum():
                         QMessageBox.warning(self, "Invalid Input", f"Value {val} is too small. Setting to minimum {self.minimum()}.")
                    elif val > self.maximum():
                         QMessageBox.warning(self, "Invalid Input", f"Value {val} is too large. Setting to maximum {self.maximum()}.")
                except:
                    pass
                return super().fixup(text)
        
        self.spin_count = WarningSpinBox()
        self.spin_count.setRange(2, 2000)
        self.spin_count.setValue(50)
        self.spin_count.setCorrectionMode(QAbstractSpinBox.CorrectToNearestValue)
        
        self.slider_epsilon = QSlider(Qt.Horizontal)
        self.slider_epsilon.setRange(1, 200) 
        self.slider_epsilon.setValue(100) 
        
        gb_ext_mode = QGroupBox("Extraction Mode")
        ext_mode_layout = QVBoxLayout()
        self.radio_segmented = QRadioButton("Segmented (Broken Lines)")
        self.radio_continuous = QRadioButton("Continuous (Smooth)")
        self.radio_segmented.setChecked(True)
        self.radio_segmented.setToolTip("Best for lines interrupted by grid or gaps. Captures multiple segments.")
        self.radio_continuous.setToolTip("Best for single smooth lines. Filters noise aggressively.")
        
        # Link dotted to segmented (Step 1146/1154)
        self.radio_dotted.toggled.connect(self.on_line_mode_changed)

        ext_mode_layout.addWidget(self.radio_segmented)
        ext_mode_layout.addWidget(self.radio_continuous)
        gb_ext_mode.setLayout(ext_mode_layout)
        extract_layout.addWidget(gb_ext_mode)
        
        self.slider_epsilon.setToolTip("Drag Left for Lower Detail (Smoother), Right for Higher Detail (More Points)")
        
        sampling_layout.addWidget(self.radio_raw)
        
        layout_fixed = QHBoxLayout()
        layout_fixed.addWidget(self.radio_fixed)
        layout_fixed.addWidget(self.spin_count)
        sampling_layout.addLayout(layout_fixed)
        
        layout_key = QHBoxLayout()
        self.radio_key.setText("Key Points")
        
        self.btn_help_key = QToolButton()
        self.btn_help_key.setText("?")
        self.btn_help_key.setFixedSize(20, 20)
        self.btn_help_key.setToolTip("Adjust detail level. Left = Low Detail, Right = High Detail.")
        self.btn_help_key.setStyleSheet("border-radius: 10px; background-color: #ddd; font-weight: bold;")
        
        layout_key.addWidget(self.radio_key)
        layout_key.addWidget(self.btn_help_key)
        layout_key.addWidget(QLabel("Detail:"))
        layout_key.addWidget(self.slider_epsilon)
        sampling_layout.addLayout(layout_key)
        
        gb_sampling.setLayout(sampling_layout)
        extract_layout.addWidget(gb_sampling)
        
        gb_extract.setLayout(extract_layout)
        left_layout.addWidget(gb_extract)

        # Range Limit
        gb_range = QGroupBox("Data Range (Data Coordinates)")
        range_layout = QVBoxLayout()
        
        rx_layout = QHBoxLayout()
        self.chk_range_x = QCheckBox("Limit X")
        self.chk_range_x.toggled.connect(self.toggle_range_inputs)
        self.inp_min_x_val = QLineEdit()
        self.inp_min_x_val.setPlaceholderText("Min X")
        self.inp_max_x_val = QLineEdit()
        self.inp_max_x_val.setPlaceholderText("Max X")
        rx_layout.addWidget(self.chk_range_x)
        rx_layout.addWidget(self.inp_min_x_val)
        rx_layout.addWidget(self.inp_max_x_val)
        
        ry_layout = QHBoxLayout()
        self.chk_range_y = QCheckBox("Limit Y")
        self.chk_range_y.toggled.connect(self.toggle_range_inputs)
        self.inp_min_y_val = QLineEdit()
        self.inp_min_y_val.setPlaceholderText("Min Y")
        self.inp_max_y_val = QLineEdit()
        self.inp_max_y_val.setPlaceholderText("Max Y")
        ry_layout.addWidget(self.chk_range_y)
        ry_layout.addWidget(self.inp_min_y_val)
        ry_layout.addWidget(self.inp_max_y_val)
        
        range_layout.addLayout(rx_layout)
        range_layout.addLayout(ry_layout)
        gb_range.setLayout(range_layout)
        left_layout.addWidget(gb_range)
        
        left_layout.addStretch()
        
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_widget.setLayout(left_layout)
        scroll_area.setWidget(scroll_widget)
        scroll_area.setFixedWidth(380)
        
        main_layout.addWidget(scroll_area)
        
        # 2. Right Layout (Canvas + Data Table)
        right_layout = QVBoxLayout()
        
        # 2.1 Canvas
        self.canvas = ImageCanvas()
        self.canvas.calibration_point_added.connect(self.on_calibration_point_added)
        self.canvas.point_picked.connect(self.on_canvas_point_picked)
        right_layout.addWidget(self.canvas, stretch=2)
        
        # 2.2 Data Table (Bottom)
        gb_table = QGroupBox("Extracted Data")
        table_layout = QVBoxLayout()
        
        self.table = QTableWidget()
        self.table.setColumnCount(3) # Initial default
        self.table.setHorizontalHeaderLabels(["Index", "Series 1 X", "Series 1 Y"])
        self.table.horizontalHeader().sectionDoubleClicked.connect(self.on_header_double_clicked)
        table_layout.addWidget(self.table)
        
        series_btn_layout = QHBoxLayout()
        btn_delete_series = QPushButton("Delete Selected Series")
        btn_delete_series.clicked.connect(self.delete_selected_series)
        btn_clear_series = QPushButton("Clear All")
        btn_clear_series.clicked.connect(self.clear_all_series)
        series_btn_layout.addWidget(btn_delete_series)
        series_btn_layout.addWidget(btn_clear_series)
        table_layout.addLayout(series_btn_layout)
        
        btn_export = QPushButton("Export to CSV")
        btn_export.clicked.connect(self.export_csv)
        table_layout.addWidget(btn_export)
        gb_table.setLayout(table_layout)
        
        right_layout.addWidget(gb_table, stretch=1)
        
        main_layout.addLayout(right_layout, stretch=1)
        central_widget.setLayout(main_layout)
        
        self.lbl_status = QLabel("Ready")
        self.statusBar().addWidget(self.lbl_status)
        
        self.toggle_range_inputs()
    
    # Methods
    def load_image_ui(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Image", "", "Images (*.png *.jpg *.jpeg *.bmp)")
        if file_path:
            self.project_model.set_image(file_path)
            self.canvas.load_image(file_path)
            self.lbl_status.setText(f"Loaded: {file_path}")

    def on_project_updated(self):
        # Only clear extracted point items - NOT scene.clear() which would erase mask too
        self.canvas.clear_extracted_points()
        
        for series in self.project_model.series_list:
             gradients = series.calculate_instant_gradients()
             self.canvas.draw_extracted_points(series.raw_pixels, series.data_points, gradients=gradients, color=series.color)
        
        next_idx = len(self.project_model.series_list) + 1
        self.inp_series_name.setText(f"Series {next_idx}")
        self.update_table()

    def toggle_calibration_mode(self):
        if self.btn_calib_mode.isChecked():
            self.canvas.set_mode(ImageCanvas.MODE_CALIBRATE)
        else:
            self.canvas.set_mode(ImageCanvas.MODE_VIEW)

    def toggle_range_inputs(self):
        x_enabled = self.chk_range_x.isChecked()
        self.inp_min_x_val.setEnabled(x_enabled)
        self.inp_max_x_val.setEnabled(x_enabled)
        
        y_enabled = self.chk_range_y.isChecked()
        self.inp_min_y_val.setEnabled(y_enabled)
        self.inp_max_y_val.setEnabled(y_enabled)

    def toggle_gap_slider(self):
        if self.radio_dotted.isChecked():
            self.container_gap.setVisible(True)
        else:
            self.container_gap.setVisible(False)
            
    def on_line_mode_changed(self):
        if self.radio_dotted.isChecked():
            self.radio_segmented.setChecked(True)
            self.radio_continuous.setEnabled(False)
            self.radio_segmented.setEnabled(True) 
        else:
            self.radio_continuous.setEnabled(True)
            self.radio_segmented.setEnabled(True)
            
    def update_gap_label(self):
        val = self.slider_gap.value()
        self.lbl_gap.setText(f"Gap Fill: {val}px")

    def toggle_mask_mode(self):
        if self.btn_mask_mode.isChecked():
            if self.btn_mask_eraser.isChecked():
                self.btn_mask_eraser.setChecked(False)
            self.canvas.set_mode(ImageCanvas.MODE_MASK)
        else:
            self.canvas.set_mode(ImageCanvas.MODE_VIEW)
            
    def toggle_mask_eraser(self):
        if self.btn_mask_eraser.isChecked():
            if self.btn_mask_mode.isChecked():
                self.btn_mask_mode.setChecked(False)
            self.canvas.set_mode(ImageCanvas.MODE_MASK_ERASER)
        else:
            self.canvas.set_mode(ImageCanvas.MODE_VIEW)
            
    def undo_mask(self):
        self.canvas.undo_last_mask()
        
    def clear_mask(self):
        self.canvas.clear_mask()
        
    def pick_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            h, s, v, _ = color.getHsv()
            cv_h = int(h / 2)   # Qt 0-359 -> OpenCV 0-179
            cv_s = int(s)
            cv_v = int(v)
            # Build HSV range (lower, upper) that process_images expects
            tolerance_h = 12
            tolerance_s = 60
            tolerance_v = 60
            lower = np.array([max(0, cv_h - tolerance_h),
                              max(0, cv_s - tolerance_s),
                              max(0, cv_v - tolerance_v)])
            upper = np.array([min(179, cv_h + tolerance_h),
                              min(255, cv_s + tolerance_s),
                              min(255, cv_v + tolerance_v)])
            self.target_color_hsv = (lower, upper)
            self.btn_color.setText(f"Color: HSV({cv_h},{cv_s},{cv_v})")
            self.btn_color.setStyleSheet(f"background-color: {color.name()}; color: {'white' if v < 128 else 'black'}")
            self.current_series_color = color
            

    def update_calib_instructions(self):
         self.lbl_status.setText("Standard: Click Axis Points X1, X2, Y1, Y2")

    def reset_calibration(self):
        self.canvas.clear_calibration_points()
        self.btn_calib_mode.setText("Set Calibration Points Points (0/4)")
        self.btn_calib_mode.setChecked(False)
        self.project_model.calibration.pixel_points = []
        self.canvas.set_mode(ImageCanvas.MODE_VIEW)
        self.lbl_status.setText("Calibration Reset. Please set 4 points.")

    def on_calibration_point_added(self, idx, x, y):
        count = idx + 1
        self.btn_calib_mode.setText(f"Set Calibration Points ({count}/4)")
        if count == 4:
            self.btn_calib_mode.setChecked(False)
            self.canvas.set_mode(ImageCanvas.MODE_VIEW)
            
    def on_canvas_point_picked(self, x, y):
        # Implement if needed, was in original?
        pass

    def extract_data(self):
        series_name_input = self.inp_series_name.text()
        
        if not self.project_model.image_path:
            QMessageBox.warning(self, "Error", "Please load an image first.")
            return
            
        if len(self.canvas.calibration_points) != 4:
            QMessageBox.warning(self, "Error", "Please set 4 calibration points.")
            return

        try:
            x1 = float(self.inp_x1.text())
            x2 = float(self.inp_x2.text())
            y1 = float(self.inp_y1.text())
            y2 = float(self.inp_y2.text())
            is_log_x = self.chk_log_x.isChecked()
            is_log_y = self.chk_log_y.isChecked()
        except ValueError:
            QMessageBox.warning(self, "Error", "Invalid calibration values.")
            return
            
        calib_points_px = []
        for item in self.canvas.calibration_points:
            r = item.rect()
            calib_points_px.append((r.center().x(), r.center().y()))

        try:
            calib_values = [(x1, 0), (x2, 0), (0, y1), (0, y2)] 
            self.project_model.update_calibration(calib_points_px, calib_values, is_log_x, is_log_y)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Calibration failed: {str(e)}")
            return

        # Get Mask
        try:
            mask_qimage = self.canvas.get_mask_image()
            if mask_qimage is None:
                 QMessageBox.warning(self, "Error", "Failed to get mask.")
                 return
            
            ptr = mask_qimage.constBits()
            width = mask_qimage.width()
            height = mask_qimage.height()
            bytes_per_line = mask_qimage.bytesPerLine()

            raw_arr = np.array(ptr, copy=True).reshape(height, bytes_per_line)
            # Strip row padding: keep only actual pixel columns
            mask_cv = raw_arr[:, :width].copy()
            
            # Original Image
            qimg_orig = self.canvas.current_image.toImage().convertToFormat(QImage.Format_RGB888)
            width_o = qimg_orig.width()
            height_o = qimg_orig.height()
            bytes_o = qimg_orig.bytesPerLine()
            
            orig_arr = np.array(qimg_orig.constBits(), copy=True).reshape(height_o, bytes_o)
            # Slice off row padding: bytesPerLine may be larger than width*3
            orig_cv = orig_arr[:, :width_o * 3].reshape(height_o, width_o, 3)
            orig_cv = cv2.cvtColor(orig_cv, cv2.COLOR_RGB2BGR)

            processor = ImageProcessor()
            
            line_type = 'auto'
            gap_fill = None
            
            if self.radio_dotted.isChecked():
                line_type = 'dotted'
                gap_val = self.slider_gap.value()
                gap_fill = gap_val
            else:
                line_type = 'solid'
                
            ext_mode = 'segmented'
            if self.radio_continuous.isChecked():
                ext_mode = 'continuous'

            points, _ = processor.process_images(
                orig_cv, 
                mask_cv, 
                hsv_range=self.target_color_hsv,
                line_type=line_type, 
                gap_fill=gap_fill,
                extraction_mode=ext_mode
            )
            if points and len(points) > 0:
                valid_points = []
                for p in points:
                    try:
                        dx, dy = self.project_model.calibration.map_to_data(p[0], p[1])
                        valid_points.append((dx, dy, p[0], p[1]))
                    except Exception as e:
                        pass
                
                # Filter Range
                final_series_data = []
                final_series_px = []
                
                # Check range enabled
                limit_x = self.chk_range_x.isChecked()
                limit_y = self.chk_range_y.isChecked()
                min_x = float(self.inp_min_x_val.text()) if limit_x and self.inp_min_x_val.text() else -float('inf')
                max_x = float(self.inp_max_x_val.text()) if limit_x and self.inp_max_x_val.text() else float('inf')
                min_y = float(self.inp_min_y_val.text()) if limit_y and self.inp_min_y_val.text() else -float('inf')
                max_y = float(self.inp_max_y_val.text()) if limit_y and self.inp_max_y_val.text() else float('inf')
                
                # Sampling logic
                if self.radio_raw.isChecked():
                    for d in valid_points:
                         if min_x <= d[0] <= max_x and min_y <= d[1] <= max_y:
                             final_series_data.append((d[0], d[1]))
                             final_series_px.append((d[2], d[3]))
                             
                elif self.radio_fixed.isChecked():
                    target_count = self.spin_count.value()
                    # Filter first
                    filtered = []
                    for d in valid_points:
                         if min_x <= d[0] <= max_x and min_y <= d[1] <= max_y:
                             filtered.append(d)
                    
                    if len(filtered) > target_count:
                         indices = np.linspace(0, len(filtered)-1, target_count, dtype=int)
                         for idx in indices:
                             d = filtered[idx]
                             final_series_data.append((d[0], d[1]))
                             final_series_px.append((d[2], d[3]))
                    else:
                        for d in filtered:
                             final_series_data.append((d[0], d[1]))
                             final_series_px.append((d[2], d[3]))
                             
                elif self.radio_key.isChecked():
                    # Douglas-Peucker
                    filtered = []
                    for d in valid_points:
                         if min_x <= d[0] <= max_x and min_y <= d[1] <= max_y:
                             filtered.append(d)
                             
                    if len(filtered) < 3:
                        for d in filtered:
                             final_series_data.append((d[0], d[1]))
                             final_series_px.append((d[2], d[3]))
                    else:
                        data_pts = np.array([ [[p[0], p[1]]] for p in filtered ], dtype=np.float32)
                        
                        print(f"DEBUG: Key Points: Filtered to {len(filtered)} before approx")
                        
                        slider_val = self.slider_epsilon.value()
                        min_eps = 0.001
                        max_eps = 1.0 # Reduced from 5.0
                        
                        t = (slider_val - 1) / 199.0 # 0.0 to 1.0
                        epsilon = max_eps - t * (max_eps - min_eps)
                        
                        approx_curve = cv2.approxPolyDP(data_pts, epsilon, False)
                        
                        # Ensure native float
                        resamp_data = [(float(p[0][0]), float(p[0][1])) for p in approx_curve]
                        final_series_data = resamp_data
                         
                        start_search_idx = 0
                        
                        x_check_val = 0.0
                        try:
                            x_check_val = float(self.inp_x2.text())
                        except:
                            x_check_val = 100.0 # Default fallback
                        
                        for rdx, rdy in final_series_data:
                            best_idx = start_search_idx
                            min_err = float('inf')
                            
                            for k in range(start_search_idx, len(filtered)):
                                d = filtered[k]
                                err = (d[0]-rdx)**2 + (d[1]-rdy)**2
                                if err < min_err:
                                    min_err = err
                                    best_idx = k
                                if (d[0] - rdx) > abs(x_check_val) * 0.1: 
                                    break
                                    
                            start_search_idx = best_idx
                            final_series_px.append((filtered[best_idx][2], filtered[best_idx][3]))
                            
                # Add Series
                vis_color = self.current_series_color if self.target_color_hsv else QColor(np.random.randint(0, 255), np.random.randint(0, 255), np.random.randint(0, 255))
        
                new_series = Series(series_name_input, vis_color)
                new_series.set_data(final_series_px, final_series_data)
                new_series.line_type = line_type
                if gap_fill: new_series.gap_fill = gap_fill
                
                self.project_model.add_series(new_series)
                QMessageBox.information(self, "Success", f"Added Series '{series_name_input}' with {len(final_series_data)} points.")
            else:
                 QMessageBox.warning(self, "No Data", "No points extracted based on current mask/color/grid settings.")

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Processing failed: {str(e)}")
            import traceback
            traceback.print_exc()
            return
    
    def delete_selected_series(self):
        # Determine series from selected column
        col = self.table.currentColumn()
        if col <= 0:
            QMessageBox.warning(self, "Error", "Please select a series column (X or Y) to delete.")
            return
            
        series_idx = (col - 1) // 2
        if series_idx < 0 or series_idx >= len(self.project_model.series_list):
            return
            
        series = self.project_model.series_list[series_idx]
        
        # Confirm deletion
        reply = QMessageBox.question(self, "Confirm", f"Delete series '{series.name}'?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply != QMessageBox.Yes:
            return

        self.project_model.remove_series(series_idx)
        self.canvas.clear_extracted_points()
        
        # Re-draw remaining logic
        for s in self.project_model.series_list:
            gradients = s.calculate_instant_gradients()
            self.canvas.draw_extracted_points(s.raw_pixels, s.data_points, gradients=gradients, color=s.color)
        
        self.update_table()

    def clear_all_series(self):
        reply = QMessageBox.question(self, "Confirm", "Clear all series?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.project_model.clear_data()
            self.canvas.clear_extracted_points()
            self.update_table()

    def update_table(self):
        # Side-by-Side Table Logic
        self.table.clear()
        
        series_list = self.project_model.series_list
        if not series_list:
            self.table.setColumnCount(0)
            self.table.setRowCount(0)
            return
            
        # 1. Setup Columns
        # Column 0: Index
        # Column 1,2: S1_X, S1_Y, ...
        col_count = 1 + 2 * len(series_list)
        self.table.setColumnCount(col_count)
        
        headers = ["Index"]
        for s in series_list:
            headers.append(f"{s.name} X")
            headers.append(f"{s.name} Y")
        self.table.setHorizontalHeaderLabels(headers)
        
        # 2. Determine Max Rows
        max_rows = 0
        for s in series_list:
            if len(s.data_points) > max_rows:
                max_rows = len(s.data_points)
        self.table.setRowCount(max_rows)
        
        # 3. Populate Index Column
        for i in range(max_rows):
            self.table.setItem(i, 0, QTableWidgetItem(str(i+1)))
            
        # 4. Populate Data
        for s_idx, s in enumerate(series_list):
            col_x = 1 + s_idx * 2
            col_y = col_x + 1
            
            for i, (x, y) in enumerate(s.data_points):
                self.table.setItem(i, col_x, QTableWidgetItem(f"{x:.4f}"))
                self.table.setItem(i, col_y, QTableWidgetItem(f"{y:.4f}"))

    def on_header_double_clicked(self, index):
        if index <= 0:
            return
            
        series_idx = (index - 1) // 2
        if series_idx < 0 or series_idx >= len(self.project_model.series_list):
            return
            
        s = self.project_model.series_list[series_idx]
        new_name, ok = QInputDialog.getText(self, "Rename Series", "Enter new name:", text=s.name)
        
        if ok and new_name:
            s.name = new_name
            self.update_table()
                
    def export_csv(self):
        if not self.project_model.series_list:
            return
            
        file_path, _ = QFileDialog.getSaveFileName(self, "Save CSV", "", "CSV Files (*.csv)")
        if file_path:
            try:
                header, rows = self.project_model.get_csv_data()
                with open(file_path, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(header)
                    writer.writerows(rows)
                QMessageBox.information(self, "Saved", f"Data saved to {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to save: {str(e)}")

    def save_project_ui(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Project", "", "JSON Files (*.json)")
        if file_path:
            try:
                self.project_model.save_project(file_path)
                QMessageBox.information(self, "Success", f"Project saved to {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to save project: {str(e)}")

    def load_project_ui(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Project", "", "JSON Files (*.json)")
        if file_path:
            try:
                self.project_model.load_project(file_path)
                if self.project_model.image_path:
                     self.canvas.load_image(self.project_model.image_path)
                     self.lbl_status.setText(f"Loaded Project: {self.project_model.image_path.split('/')[-1]}")
                QMessageBox.information(self, "Success", f"Project loaded from {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to load project: {str(e)}")
