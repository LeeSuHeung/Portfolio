import sys
import os
import shutil
import pandas as pd
import xlwings as xw
import tempfile
import time
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QTableWidget, QTableWidgetItem, 
                             QSplitter, QMessageBox, QComboBox, QFileDialog, 
                             QCheckBox, QProgressBar, QFrame, QAbstractItemView, 
                             QStyleFactory, QSpinBox, QShortcut, QHeaderView)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QColor, QDragEnterEvent, QDropEvent, QKeySequence

# =========================================================
# [UI 컴포넌트] 드래그 앤 드롭 파일 로더
# =========================================================
class FileDropZone(QLabel):
    fileDropped = pyqtSignal(str)

    def __init__(self, title, color_theme):
        super().__init__()
        self.setText(f"\n📂 {title}\n(Drag & Drop)")
        self.setAlignment(Qt.AlignCenter)
        self.setWordWrap(True)
        self.default_style = f"""
            QLabel {{
                border: 2px dashed {color_theme};
                border-radius: 10px;
                background-color: #f8f9fa;
                color: #555;
                font-weight: bold;
                font-size: 13px;
                padding: 5px;
            }}
        """
        self.setStyleSheet(self.default_style)
        self.setAcceptDrops(True)
        self.setFixedHeight(90)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls(): event.accept()
        else: event.ignore()

    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files:
            path = files[0]
            if path.lower().endswith(('.xlsx', '.xls', '.xlsm')):
                filename = os.path.basename(path)
                self.setText(f"✅ {filename}\n📁 {path}")
                self.setStyleSheet(self.default_style.replace("dashed", "solid").replace("#f8f9fa", "#e8f5e9"))
                self.fileDropped.emit(path)

# =========================================================
# [Core] 백그라운드 비교 로직
# =========================================================
class AnalyzerThread(QThread):
    finished = pyqtSignal(object, object, list, list, list) 

    def __init__(self, path_a, sheet_a, path_b, sheet_b, key_col, header_idx):
        super().__init__()
        self.path_a, self.sheet_a = path_a, sheet_a
        self.path_b, self.sheet_b = path_b, sheet_b
        self.key_col = key_col
        self.header_idx = header_idx

    def run(self):
        try:
            df_a = pd.read_excel(self.path_a, sheet_name=self.sheet_a, header=self.header_idx).fillna("")
            df_b = pd.read_excel(self.path_b, sheet_name=self.sheet_b, header=self.header_idx).fillna("")

            df_a = df_a.astype(str)
            df_b = df_b.astype(str)

            diff_cells = []       
            missing_in_b = []     
            missing_in_a = []     

            if self.key_col and self.key_col in df_a.columns and self.key_col in df_b.columns:
                df_a_idx = df_a.set_index(self.key_col)
                df_b_idx = df_b.set_index(self.key_col)
                
                keys_only_a = df_a_idx.index.difference(df_b_idx.index)
                keys_only_b = df_b_idx.index.difference(df_a_idx.index)
                common_keys = df_a_idx.index.intersection(df_b_idx.index)

                for k in keys_only_a:
                    missing_in_b.append(df_a[df_a[self.key_col] == k].index[0])

                for k in keys_only_b:
                    missing_in_a.append(df_b[df_b[self.key_col] == k].index[0])

                common_cols = df_a_idx.columns.intersection(df_b_idx.columns)
                for k in common_keys:
                    row_a = df_a_idx.loc[k]
                    row_b = df_b_idx.loc[k]
                    
                    if isinstance(row_a, pd.DataFrame): row_a = row_a.iloc[0]
                    if isinstance(row_b, pd.DataFrame): row_b = row_b.iloc[0]

                    ui_row_a = df_a[df_a[self.key_col] == k].index[0]
                    ui_row_b = df_b[df_b[self.key_col] == k].index[0]

                    for col in common_cols:
                        val_a = row_a[col]
                        val_b = row_b[col]
                        if val_a != val_b:
                            diff_cells.append({
                                'key': k,
                                'r_a': ui_row_a, 'c_a': df_a.columns.get_loc(col), 'val_a': val_a,
                                'r_b': ui_row_b, 'c_b': df_b.columns.get_loc(col), 'val_b': val_b
                            })
            else:
                min_rows = min(len(df_a), len(df_b))
                min_cols = min(len(df_a.columns), len(df_b.columns))
                
                for r in range(min_rows):
                    for c in range(min_cols):
                        val_a = df_a.iloc[r, c]
                        val_b = df_b.iloc[r, c]
                        if val_a != val_b:
                            diff_cells.append({
                                'key': f"Row {r}",
                                'r_a': r, 'c_a': c, 'val_a': val_a,
                                'r_b': r, 'c_b': c, 'val_b': val_b
                            })
                
                if len(df_a) > len(df_b):
                    missing_in_b = list(range(len(df_b), len(df_a)))
                elif len(df_b) > len(df_a):
                    missing_in_a = list(range(len(df_a), len(df_b)))

            self.finished.emit(df_a, df_b, diff_cells, missing_in_b, missing_in_a)

        except Exception as e:
            print(f"Analysis Error: {e}")
            self.finished.emit(None, None, [], [], [])

# =========================================================
# [Main App] 메인 윈도우
# =========================================================
class ExcelSyncPro(QWidget):
    def __init__(self, file_a=None, file_b=None):
        super().__init__()
        self.path_a = ""
        self.path_b = ""
        self.df_a = None
        self.df_b = None
        self.diff_data = []      
        self.missing_in_b = []   
        self.missing_in_a = []   
        self.current_sheet = None
        
        self.error_targets = []
        self.current_error_idx = -1
        
        self.undo_stack = []
        
        self.initUI()

        if file_a and file_b:
            self.load_file(file_a, 'A')
            self.load_file(file_b, 'B')
            self.drop_a.setText(f"✅ {os.path.basename(file_a)}\n📁 {file_a}")
            self.drop_a.setStyleSheet(self.drop_a.default_style.replace("dashed", "solid").replace("#f8f9fa", "#e8f5e9"))
            self.drop_b.setText(f"✅ {os.path.basename(file_b)}\n📁 {file_b}")
            self.drop_b.setStyleSheet(self.drop_b.default_style.replace("dashed", "solid").replace("#f8f9fa", "#e8f5e9"))
            
    def initUI(self):
        self.setWindowTitle('Excel Sync Master Pro v8 (Anti-Crash & Undo)')
        self.resize(1600, 1000)
        self.setStyle(QStyleFactory.create('Fusion'))

        main_layout = QVBoxLayout()

        top_layout = QHBoxLayout()
        self.drop_a = FileDropZone("기준 파일 (Source A)", "#3f51b5")
        self.drop_a.fileDropped.connect(lambda p: self.load_file(p, 'A'))
        self.drop_b = FileDropZone("비교 파일 (Target B)", "#f44336")
        self.drop_b.fileDropped.connect(lambda p: self.load_file(p, 'B'))
        
        top_layout.addWidget(self.drop_a)
        top_layout.addWidget(self.drop_b)
        main_layout.addLayout(top_layout)

        ctrl_panel = QFrame()
        ctrl_panel.setStyleSheet("background: #f5f5f5; border-radius: 5px; border: 1px solid #ddd;")
        ctrl_layout = QHBoxLayout(ctrl_panel)

        ctrl_layout.addWidget(QLabel("<b>📌 시트:</b>"))
        self.combo_sheet = QComboBox()
        self.combo_sheet.currentIndexChanged.connect(self.on_sheet_change)
        ctrl_layout.addWidget(self.combo_sheet)

        ctrl_layout.addSpacing(15)
        
        ctrl_layout.addWidget(QLabel("<b>🔢 헤더 행(Row):</b>"))
        self.spin_header = QSpinBox()
        self.spin_header.setRange(1, 100)
        self.spin_header.setValue(4) 
        self.spin_header.setToolTip("데이터의 제목(컬럼명)이 있는 행 번호를 지정하세요.")
        self.spin_header.valueChanged.connect(self.on_sheet_change) 
        ctrl_layout.addWidget(self.spin_header)

        ctrl_layout.addSpacing(15)
        ctrl_layout.addWidget(QLabel("<b>🔑 기준 Key:</b>"))
        self.combo_key = QComboBox()
        self.combo_key.setToolTip("고유 ID 컬럼 선택 시 행이 섞여도 비교 가능")
        ctrl_layout.addWidget(self.combo_key)

        self.btn_analyze = QPushButton("🔍 분석 실행 (F5)")
        self.btn_analyze.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; padding: 6px 15px;")
        self.btn_analyze.clicked.connect(self.run_analysis)
        ctrl_layout.addWidget(self.btn_analyze)

        self.btn_undo = QPushButton("↩ 되돌리기 (Ctrl+Z)")
        self.btn_undo.setStyleSheet("background-color: #757575; color: white; font-weight: bold; padding: 6px 15px;")
        self.btn_undo.setEnabled(False)
        self.btn_undo.clicked.connect(self.undo_action)
        ctrl_layout.addWidget(self.btn_undo)

        ctrl_layout.addStretch()
        
        self.chk_autosave = QCheckBox("자동 저장 (Auto-Save)")
        self.chk_autosave.setChecked(True) 
        self.chk_autosave.setStyleSheet("color: #d32f2f; font-weight: bold;")
        ctrl_layout.addWidget(self.chk_autosave)
        
        main_layout.addWidget(ctrl_panel)

        self.main_v_splitter = QSplitter(Qt.Vertical)

        top_tables_widget = QWidget()
        top_tables_layout = QVBoxLayout(top_tables_widget)
        top_tables_layout.setContentsMargins(0, 0, 0, 0)

        self.splitter = QSplitter(Qt.Horizontal)
        
        wid_a = QWidget()
        v_a = QVBoxLayout(wid_a)
        v_a.setContentsMargins(0,0,0,0)
        self.lbl_status_a = QLabel("File A")
        self.lbl_status_a.setStyleSheet("font-weight: bold; color: #303f9f;")
        self.table_a = self.create_table()
        
        btn_box_a = QHBoxLayout()
        self.btn_save_a = QPushButton("💾 A 저장")
        self.btn_save_a.clicked.connect(lambda: self.save_file('A'))
        
        self.btn_b_to_a_row_sel = QPushButton("➕ B 선택행 가져오기")
        self.btn_b_to_a_row_sel.clicked.connect(lambda: self.sync_rows_missing('A', only_selected=True))
        self.btn_b_to_a_row_sel.setStyleSheet("color: #E65100;")
        
        self.btn_b_to_a_row_all = QPushButton("➕ B 전체행 가져오기")
        self.btn_b_to_a_row_all.clicked.connect(lambda: self.sync_rows_missing('A', only_selected=False))
        self.btn_b_to_a_row_all.setStyleSheet("color: #E65100; font-weight: bold;")
        
        self.btn_b_to_a_val_sel = QPushButton("⚡ B 선택값 덮기")
        self.btn_b_to_a_val_sel.clicked.connect(lambda: self.sync_values('A', only_selected=True))
        self.btn_b_to_a_val_sel.setStyleSheet("color: #D32F2F;")
        
        self.btn_b_to_a_val_all = QPushButton("⚡ B 전체값 덮기")
        self.btn_b_to_a_val_all.clicked.connect(lambda: self.sync_values('A', only_selected=False))
        self.btn_b_to_a_val_all.setStyleSheet("color: #D32F2F; font-weight: bold;")

        btn_box_a.addWidget(self.btn_save_a)
        btn_box_a.addStretch()
        btn_box_a.addWidget(self.btn_b_to_a_row_sel)
        btn_box_a.addWidget(self.btn_b_to_a_row_all)
        btn_box_a.addWidget(self.btn_b_to_a_val_sel)
        btn_box_a.addWidget(self.btn_b_to_a_val_all)
        
        v_a.addWidget(self.lbl_status_a)
        v_a.addWidget(self.table_a)
        v_a.addLayout(btn_box_a)
        
        wid_b = QWidget()
        v_b = QVBoxLayout(wid_b)
        v_b.setContentsMargins(0,0,0,0)
        self.lbl_status_b = QLabel("File B")
        self.lbl_status_b.setStyleSheet("font-weight: bold; color: #d32f2f;")
        self.table_b = self.create_table()

        btn_box_b = QHBoxLayout()
        self.btn_a_to_b_val_sel = QPushButton("⚡ A 선택값 덮기")
        self.btn_a_to_b_val_sel.clicked.connect(lambda: self.sync_values('B', only_selected=True))
        self.btn_a_to_b_val_sel.setStyleSheet("color: #1976D2;")

        self.btn_a_to_b_val_all = QPushButton("⚡ A 전체값 덮기")
        self.btn_a_to_b_val_all.clicked.connect(lambda: self.sync_values('B', only_selected=False))
        self.btn_a_to_b_val_all.setStyleSheet("color: #1976D2; font-weight: bold;")

        self.btn_a_to_b_row_sel = QPushButton("➕ A 선택행 가져오기")
        self.btn_a_to_b_row_sel.clicked.connect(lambda: self.sync_rows_missing('B', only_selected=True))
        self.btn_a_to_b_row_sel.setStyleSheet("color: #388E3C;")

        self.btn_a_to_b_row_all = QPushButton("➕ A 전체행 가져오기")
        self.btn_a_to_b_row_all.clicked.connect(lambda: self.sync_rows_missing('B', only_selected=False))
        self.btn_a_to_b_row_all.setStyleSheet("color: #388E3C; font-weight: bold;")
        
        self.btn_save_b = QPushButton("💾 B 저장")
        self.btn_save_b.clicked.connect(lambda: self.save_file('B'))

        btn_box_b.addWidget(self.btn_a_to_b_val_sel)
        btn_box_b.addWidget(self.btn_a_to_b_val_all)
        btn_box_b.addWidget(self.btn_a_to_b_row_sel)
        btn_box_b.addWidget(self.btn_a_to_b_row_all)
        btn_box_b.addStretch()
        btn_box_b.addWidget(self.btn_save_b)

        v_b.addWidget(self.lbl_status_b)
        v_b.addWidget(self.table_b)
        v_b.addLayout(btn_box_b)

        self.splitter.addWidget(wid_a)
        self.splitter.addWidget(wid_b)
        self.splitter.setSizes([800, 800])
        
        top_tables_layout.addWidget(self.splitter)

        bottom_summary_widget = QWidget()
        diff_summary_layout = QVBoxLayout(bottom_summary_widget)
        diff_summary_layout.setContentsMargins(0, 5, 0, 0)
        
        diff_label = QLabel("<b>📊 상세 차이점 요약 (리스트 더블 클릭 시 해당 위치로 즉시 이동 / 경계선을 드래그하여 크기 조절)</b>")
        diff_label.setStyleSheet("color: #424242;")
        
        self.diff_summary_table = QTableWidget()
        self.diff_summary_table.setColumnCount(4)
        self.diff_summary_table.setHorizontalHeaderLabels(["위치 (행 또는 Key)", "컬럼명", "File A 값", "File B 값"])
        self.diff_summary_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.diff_summary_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.diff_summary_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.diff_summary_table.setStyleSheet("gridline-color: #e0e0e0; selection-background-color: #ffcc80;")
        self.diff_summary_table.itemDoubleClicked.connect(self.on_diff_item_clicked) 

        diff_summary_layout.addWidget(diff_label)
        diff_summary_layout.addWidget(self.diff_summary_table)

        self.main_v_splitter.addWidget(top_tables_widget)
        self.main_v_splitter.addWidget(bottom_summary_widget)
        self.main_v_splitter.setSizes([650, 250]) 

        main_layout.addWidget(self.main_v_splitter)

        self.setup_sync_scrolling()
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        main_layout.addWidget(self.progress)
        self.setLayout(main_layout)

        QShortcut(QKeySequence("F5"), self).activated.connect(self.run_analysis)
        QShortcut(QKeySequence("F7"), self).activated.connect(lambda: self.navigate_error(-1))
        QShortcut(QKeySequence("F8"), self).activated.connect(lambda: self.navigate_error(1))
        QShortcut(QKeySequence("Ctrl+Z"), self).activated.connect(self.undo_action)

    def create_table(self):
        table = QTableWidget()
        table.setAlternatingRowColors(True)
        table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        table.setStyleSheet("gridline-color: #e0e0e0; selection-background-color: #bbdefb; selection-color: black;")
        return table

    def setup_sync_scrolling(self):
        self.table_a.verticalScrollBar().valueChanged.connect(self.table_b.verticalScrollBar().setValue)
        self.table_b.verticalScrollBar().valueChanged.connect(self.table_a.verticalScrollBar().setValue)
        self.table_a.horizontalScrollBar().valueChanged.connect(self.table_b.horizontalScrollBar().setValue)
        self.table_b.horizontalScrollBar().valueChanged.connect(self.table_a.horizontalScrollBar().setValue)

    def load_file(self, path, side):
        if side == 'A': self.path_a = path
        else: self.path_b = path

        self.undo_stack.clear()
        self.update_undo_btn()

        if self.path_a and self.path_b:
            try:
                xl_a = pd.ExcelFile(self.path_a)
                xl_b = pd.ExcelFile(self.path_b)
                common = list(set(xl_a.sheet_names) & set(xl_b.sheet_names))
                common.sort()
                
                self.combo_sheet.blockSignals(True)
                self.combo_sheet.clear()
                self.combo_sheet.addItems(common)
                self.combo_sheet.blockSignals(False)
                
                if common: self.on_sheet_change()
            except Exception as e:
                QMessageBox.critical(self, "오류", f"파일 읽기 실패: {e}")

    def on_sheet_change(self):
        sheet = self.combo_sheet.currentText()
        if not sheet: return
        self.current_sheet = sheet

        h_idx = self.spin_header.value() - 1
        
        try:
            df = pd.read_excel(self.path_a, sheet_name=sheet, header=h_idx, nrows=1)
            self.combo_key.blockSignals(True)
            self.combo_key.clear()
            self.combo_key.addItem("선택 안 함 (행 번호 기준)")
            self.combo_key.addItems([str(c) for c in df.columns])
            self.combo_key.blockSignals(False)
            self.run_analysis()
        except Exception:
            pass 

    def run_analysis(self):
        if not self.path_a or not self.path_b: return
        key_text = self.combo_key.currentText()
        key_col = key_text if "선택 안 함" not in key_text else None
        
        h_idx = self.spin_header.value() - 1

        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.btn_analyze.setEnabled(False)

        self.worker = AnalyzerThread(self.path_a, self.current_sheet, self.path_b, self.current_sheet, key_col, h_idx)
        self.worker.finished.connect(self.on_analysis_done)
        self.worker.start()

    def on_analysis_done(self, df_a, df_b, diff_cells, missing_b, missing_a):
        self.progress.setVisible(False)
        self.btn_analyze.setEnabled(True)
        if df_a is None: return

        self.df_a, self.df_b = df_a, df_b
        self.diff_data = diff_cells
        self.missing_in_b = missing_b 
        self.missing_in_a = missing_a 

        self.render_table(self.table_a, df_a)
        self.render_table(self.table_b, df_b)

        diff_color = QColor(255, 235, 59)
        text_color = QColor(255, 0, 0)
        
        self.diff_summary_table.setRowCount(len(diff_cells))

        for i, item in enumerate(diff_cells):
            it_a = self.table_a.item(item['r_a'], item['c_a'])
            if it_a: it_a.setBackground(diff_color)
            it_b = self.table_b.item(item['r_b'], item['c_b'])
            if it_b: 
                it_b.setBackground(diff_color)
                it_b.setForeground(text_color)
                it_b.setToolTip(f"A값: {item['val_a']}")

            col_name = df_a.columns[item['c_a']] if item['c_a'] < len(df_a.columns) else "Unknown"
            
            loc_item = QTableWidgetItem(str(item['key']))
            loc_item.setData(Qt.UserRole, (item['r_a'], item['r_b'], item['c_a'], item['c_b']))
            
            self.diff_summary_table.setItem(i, 0, loc_item)
            self.diff_summary_table.setItem(i, 1, QTableWidgetItem(str(col_name)))
            self.diff_summary_table.setItem(i, 2, QTableWidgetItem(str(item['val_a'])))
            self.diff_summary_table.setItem(i, 3, QTableWidgetItem(str(item['val_b'])))

        color_only_a = QColor(200, 230, 201) 
        color_only_b = QColor(255, 224, 178) 
        for r in missing_b:
            for c in range(self.table_a.columnCount()):
                it = self.table_a.item(r, c)
                if it: it.setBackground(color_only_a)
        for r in missing_a:
            for c in range(self.table_b.columnCount()):
                it = self.table_b.item(r, c)
                if it: it.setBackground(color_only_b)

        self.lbl_status_a.setText(f"File A (B에 없는 행: {len(missing_b)}건)")
        self.lbl_status_b.setText(f"File B (A에 없는 행: {len(missing_a)}건)")
        
        has_missing_b = len(missing_b) > 0
        self.btn_a_to_b_row_all.setEnabled(has_missing_b)
        self.btn_a_to_b_row_sel.setEnabled(has_missing_b)
        self.btn_a_to_b_row_all.setText(f"➕ A 전체행 가져오기 ({len(missing_b)})")
        
        has_missing_a = len(missing_a) > 0
        self.btn_b_to_a_row_all.setEnabled(has_missing_a)
        self.btn_b_to_a_row_sel.setEnabled(has_missing_a)
        self.btn_b_to_a_row_all.setText(f"➕ B 전체행 가져오기 ({len(missing_a)})")
        
        has_diffs = len(diff_cells) > 0
        self.btn_a_to_b_val_all.setEnabled(has_diffs)
        self.btn_a_to_b_val_sel.setEnabled(has_diffs)
        self.btn_b_to_a_val_all.setEnabled(has_diffs)
        self.btn_b_to_a_val_sel.setEnabled(has_diffs)

        targets = set()
        for diff in diff_cells:
            targets.add((diff['r_a'], diff['r_b']))
        for r in missing_b:
            targets.add((r, r))
        for r in missing_a:
            targets.add((r, r))
            
        self.error_targets = sorted(list(targets), key=lambda x: x[0])
        self.current_error_idx = -1 
        
    def render_table(self, table, df):
        table.blockSignals(True)
        table.clear()
        table.setRowCount(df.shape[0])
        table.setColumnCount(df.shape[1])
        table.setHorizontalHeaderLabels([str(c) for c in df.columns])
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                table.setItem(r, c, QTableWidgetItem(str(df.iloc[r, c])))
        table.blockSignals(False)

    def on_diff_item_clicked(self, item):
        row = item.row()
        loc_item = self.diff_summary_table.item(row, 0) 
        if loc_item:
            data = loc_item.data(Qt.UserRole)
            if data:
                r_a, r_b, c_a, c_b = data
                
                if r_a < self.table_a.rowCount():
                    self.table_a.clearSelection()
                    it_a = self.table_a.item(r_a, c_a)
                    if it_a:
                        self.table_a.scrollToItem(it_a, QAbstractItemView.PositionAtCenter)
                        it_a.setSelected(True)
                        
                if r_b < self.table_b.rowCount():
                    self.table_b.clearSelection()
                    it_b = self.table_b.item(r_b, c_b)
                    if it_b:
                        self.table_b.scrollToItem(it_b, QAbstractItemView.PositionAtCenter)
                        it_b.setSelected(True)

    def get_xw_sheet(self, path, sheet_name):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            wb = app.books.open(path)
            ws = wb.sheets[sheet_name]
            return app, wb, ws
        except:
            return None, None, None

    def update_undo_btn(self):
        if self.undo_stack:
            self.btn_undo.setEnabled(True)
            self.btn_undo.setText(f"↩ 되돌리기 ({len(self.undo_stack)})")
            self.btn_undo.setStyleSheet("background-color: #9C27B0; color: white; font-weight: bold; padding: 6px 15px;")
        else:
            self.btn_undo.setEnabled(False)
            self.btn_undo.setText("↩ 되돌리기")
            self.btn_undo.setStyleSheet("background-color: #757575; color: white; font-weight: bold; padding: 6px 15px;")

    def undo_action(self):
        if not self.undo_stack: return
        
        last_action = self.undo_stack.pop()
        try:
            shutil.copy2(last_action['backup_path'], last_action['original_path'])
            QMessageBox.information(self, "되돌리기 완료", f"다음 작업을 성공적으로 취소했습니다:\n[{last_action['desc']}]")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"되돌리기 실패: {e}")
        finally:
            self.update_undo_btn()
            self.run_analysis() 

    def sync_values(self, target, only_selected=False):
        if not self.diff_data: return
        
        diffs_to_process = self.diff_data
        
        if only_selected:
            selected_a = set((it.row(), it.column()) for it in self.table_a.selectedItems())
            selected_b = set((it.row(), it.column()) for it in self.table_b.selectedItems())
            
            diffs_to_process = [
                d for d in self.diff_data 
                if (d['r_a'], d['c_a']) in selected_a or (d['r_b'], d['c_b']) in selected_b
            ]
            
            if not diffs_to_process:
                QMessageBox.warning(self, "경고", "선택된 차이점 셀이 없습니다.\n표에서 변경할 노란색 셀을 먼저 선택해주세요.")
                return
        
        if target == 'A':
            target_path, s_char, t_char = self.path_a, 'B', 'A'
        else:
            target_path, s_char, t_char = self.path_b, 'A', 'B'

        msg = f"{s_char}의 선택 값 -> {t_char} 덮어쓰기 ({len(diffs_to_process)}건)" if only_selected else f"{s_char}의 전체 값 -> {t_char} 덮어쓰기 ({len(diffs_to_process)}건)"
        if QMessageBox.question(self, "확인", msg, QMessageBox.Yes | QMessageBox.No) == QMessageBox.No: return

        bak_name = f"sync_bak_{int(time.time())}.xlsx"
        backup_path = os.path.join(tempfile.gettempdir(), bak_name)
        shutil.copy2(target_path, backup_path)
        self.undo_stack.append({
            'target': target,
            'original_path': target_path,
            'backup_path': backup_path,
            'desc': msg
        })
        self.update_undo_btn()

        app, wb, ws = self.get_xw_sheet(target_path, self.current_sheet)
        if not ws: return

        try:
            # [수정] 엑셀 자동계산을 일시 중단하여 속도 향상 및 오류 방지
            app.calculation = 'manual' 
            
            self.progress.setVisible(True)
            self.progress.setRange(0, len(diffs_to_process))
            
            header_offset = self.spin_header.value() + 1 

            for i, diff in enumerate(diffs_to_process):
                if target == 'A':
                    r, c = diff['r_a'], diff['c_a']
                    val = diff['val_b']
                else:
                    r, c = diff['r_b'], diff['c_b']
                    val = diff['val_a']
                
                ws.range((r + header_offset, c + 1)).value = val 
                self.progress.setValue(i+1)
                
                # [핵심 수정] 30건마다 프로그램이 숨을 고르도록 처리 (RPC 튕김 현상 방지)
                if i % 30 == 0:
                    QApplication.processEvents()

            # [수정] 원상복구
            app.calculation = 'automatic' 
            wb.save()
            
            if not self.chk_autosave.isChecked():
                QMessageBox.information(self, "완료", f"{t_char} 파일 수정 완료.")
            wb.close()
        except Exception as e:
            QMessageBox.critical(self, "오류", str(e))
        finally:
            app.quit()
            self.progress.setVisible(False)
            self.run_analysis()

    def sync_rows_missing(self, target, only_selected=False):
        if target == 'A':
            indices, s_df, t_path, t_char = self.missing_in_a, self.df_b, self.path_a, 'A'
            source_table = self.table_b 
        else:
            indices, s_df, t_path, t_char = self.missing_in_b, self.df_a, self.path_b, 'B'
            source_table = self.table_a 

        if not indices: return
        
        indices_to_process = indices
        
        if only_selected:
            selected_rows = set(it.row() for it in source_table.selectedItems())
            indices_to_process = [r for r in indices if r in selected_rows]
            
            if not indices_to_process:
                s_char = 'B' if target == 'A' else 'A'
                QMessageBox.warning(self, "경고", f"선택된 누락 행이 없습니다.\n{s_char} 표에서 추가할 색칠된 행을 먼저 선택해주세요.")
                return

        msg = f"선택된 {len(indices_to_process)}행 추가 -> {t_char} 파일" if only_selected else f"전체 {len(indices_to_process)}행 추가 -> {t_char} 파일"
        if QMessageBox.question(self, "확인", msg, QMessageBox.Yes | QMessageBox.No) == QMessageBox.No: return

        bak_name = f"sync_bak_{int(time.time())}.xlsx"
        backup_path = os.path.join(tempfile.gettempdir(), bak_name)
        shutil.copy2(t_path, backup_path)
        self.undo_stack.append({
            'target': target,
            'original_path': t_path,
            'backup_path': backup_path,
            'desc': msg
        })
        self.update_undo_btn()

        app, wb, ws = self.get_xw_sheet(t_path, self.current_sheet)
        if not ws: return

        try:
            rows_to_add = s_df.iloc[indices_to_process]
            data = rows_to_add.values.tolist()

            if ws.used_range.last_cell.row == 1 and ws.range('A1').value is None:
                last_row = self.spin_header.value() 
            else:
                last_row = ws.used_range.last_cell.row + 1
            
            ws.range(f'A{last_row}').value = data
            wb.save()
            QMessageBox.information(self, "완료", f"{len(data)}행 추가 완료.")
            wb.close()
        except Exception as e:
            QMessageBox.critical(self, "오류", str(e))
        finally:
            app.quit()
            self.run_analysis()

    def save_file(self, target):
        path = self.path_a if target == 'A' else self.path_b
        fname, _ = QFileDialog.getSaveFileName(self, f'{target} 파일 저장', path, 'Excel Files (*.xlsx)')
        if fname:
            try:
                shutil.copy2(path, fname)
                QMessageBox.information(self, "저장", f"백업 완료:\n{fname}")
            except Exception as e:
                QMessageBox.critical(self, "실패", str(e))

    def navigate_error(self, direction):
        if not self.error_targets:
            return
            
        self.current_error_idx = (self.current_error_idx + direction) % len(self.error_targets)
            
        r_a, r_b = self.error_targets[self.current_error_idx]
        
        if r_a < self.table_a.rowCount():
            self.table_a.clearSelection()
            self.table_a.selectRow(r_a)
            item_a = self.table_a.item(r_a, 0)
            if item_a:
                self.table_a.scrollToItem(item_a, QAbstractItemView.PositionAtCenter)
                
        if r_b < self.table_b.rowCount():
            self.table_b.clearSelection()
            self.table_b.selectRow(r_b)
            item_b = self.table_b.item(r_b, 0)
            if item_b:
                self.table_b.scrollToItem(item_b, QAbstractItemView.PositionAtCenter)

if __name__ == '__main__':
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    app = QApplication(sys.argv)
    
    arg_file_a = sys.argv[1] if len(sys.argv) > 1 else None
    arg_file_b = sys.argv[2] if len(sys.argv) > 2 else None

    ex = ExcelSyncPro(arg_file_a, arg_file_b)
    ex.show()
    sys.exit(app.exec_())