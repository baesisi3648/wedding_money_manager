import sys
import sqlite3
import datetime
import openpyxl
from openpyxl.styles import Font, Alignment
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                             QTableWidget, QTableWidgetItem, QHeaderView, 
                             QComboBox, QSpinBox, QGroupBox, QMessageBox, 
                             QAbstractItemView, QDialog, QDialogButtonBox, QFormLayout)
from PyQt6.QtCore import Qt

# --- 데이터베이스 관리 클래스 ---
class Database:
    def __init__(self, db_name="wedding_list.db"):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self.create_table()

    def create_table(self):
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                amount INTEGER,
                meal_ticket INTEGER,
                category TEXT,
                note TEXT,
                created_at TEXT
            )
        """)
        self.conn.commit()

    def insert_record(self, name, amount, meal, category, note):
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.cursor.execute("""
            INSERT INTO records (name, amount, meal_ticket, category, note, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (name, amount, meal, category, note, now))
        self.conn.commit()

    def update_record(self, record_id, name, amount, meal, category, note):
        self.cursor.execute("""
            UPDATE records 
            SET name=?, amount=?, meal_ticket=?, category=?, note=?
            WHERE id=?
        """, (name, amount, meal, category, note, record_id))
        self.conn.commit()

    def delete_record(self, record_id):
        self.cursor.execute("DELETE FROM records WHERE id=?", (record_id,))
        self.conn.commit()

    def delete_all_records(self):
        self.cursor.execute("DELETE FROM records")
        self.cursor.execute("DELETE FROM sqlite_sequence WHERE name='records'")
        self.conn.commit()

    def fetch_all(self):
        self.cursor.execute("SELECT * FROM records ORDER BY id DESC")
        return self.cursor.fetchall()
    
    def check_duplicate_name(self, name):
        self.cursor.execute("SELECT count(*) FROM records WHERE name=?", (name,))
        return self.cursor.fetchone()[0] > 0

    def get_summary(self):
        self.cursor.execute("""
            SELECT count(*), sum(amount), sum(meal_ticket) FROM records
        """)
        count, total_amt, total_meal = self.cursor.fetchone()
        return count or 0, total_amt or 0, total_meal or 0

# --- 수정 팝업 다이얼로그 ---
class EditDialog(QDialog):
    def __init__(self, data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("내역 수정")
        self.resize(300, 250)
        self.data = data 
        
        layout = QFormLayout(self)
        
        self.name_edit = QLineEdit(str(data[1]))
        self.amount_edit = QLineEdit(str(data[2]))
        self.meal_edit = QSpinBox()
        self.meal_edit.setRange(0, 100)
        self.meal_edit.setValue(int(data[3]))
        self.cat_edit = QComboBox()
        self.cat_edit.addItems(["친척", "직장", "지인", "기타"])
        self.cat_edit.setCurrentText(data[4])
        self.note_edit = QLineEdit(str(data[5]))
        
        layout.addRow("이름:", self.name_edit)
        layout.addRow("금액(원):", self.amount_edit)
        layout.addRow("식권:", self.meal_edit)
        layout.addRow("분류:", self.cat_edit)
        layout.addRow("비고:", self.note_edit)
        
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def get_data(self):
        return (
            self.name_edit.text(),
            int(self.amount_edit.text().replace(',', '')),
            self.meal_edit.value(),
            self.cat_edit.currentText(),
            self.note_edit.text()
        )

# --- 메인 프로그램 ---
class WeddingManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = Database()
        self.setWindowTitle("축의금 관리 도우미 (Made by BAESISI)")
        self.resize(1100, 800)
        self.apply_theme()
        self.initUI()
        self.load_data()

    def apply_theme(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #FDFCF0; }
            
            /* 모든 위젯의 기본 글자색을 진한 갈색으로 고정 */
            QWidget { color: #5D4037; font-family: 'Malgun Gothic', sans-serif; }
            
            QLabel { color: #5D4037; font-size: 14px; }
            
            QLineEdit, QComboBox, QSpinBox {
                padding: 8px; border: 1px solid #BCAAA4; border-radius: 5px;
                background-color: #FFFFFF; font-size: 13px;
                color: #5D4037; /* 입력창 글자색 고정 */
            }
            
            QPushButton {
                background-color: #8D6E63; color: white; border-radius: 5px;
                padding: 10px; font-weight: bold; font-size: 13px;
            }
            QPushButton:hover { background-color: #6D4C41; }
            
            QTableWidget {
                background-color: #FFFFFF; 
                border: 1px solid #D7CCC8;
                gridline-color: #EFEBE9;
                color: #5D4037; /* [중요] 표 안의 글자색을 갈색으로 강제 고정 */
            }
            QTableWidget::item {
                color: #5D4037; /* 아이템 글자색도 확실하게 고정 */
            }
            QHeaderView::section {
                background-color: #EFEBE9; padding: 5px; border: 1px solid #D7CCC8;
                font-weight: bold; color: #5D4037;
            }
            
            QGroupBox {
                font-weight: bold; border: 2px solid #D7CCC8; border-radius: 10px; margin-top: 10px;
                color: #5D4037; /* 그룹박스 제목 색상 고정 */
            }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 5px; }
        """)

    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        top_layout = QHBoxLayout()
        title_box = QVBoxLayout()
        title_lbl = QLabel("❤ 축의금 관리 도우미")
        title_lbl.setStyleSheet("font-size: 26px; font-weight: bold; color: #4E342E;")
        sub_lbl = QLabel("결혼식 축의금과 식권을 완벽하게 관리하세요")
        title_box.addWidget(title_lbl)
        title_box.addWidget(sub_lbl)
        
        info_box = QVBoxLayout()
        info_box.addWidget(QLabel("Made by BAESISI", alignment=Qt.AlignmentFlag.AlignRight))
        info_box.addWidget(QLabel("문의: baesisi3648@gmail.com", alignment=Qt.AlignmentFlag.AlignRight))
        
        top_layout.addLayout(title_box)
        top_layout.addStretch()
        top_layout.addLayout(info_box)
        main_layout.addLayout(top_layout)

        self.dash_count = self.create_dash_card("총 인원", "0", "명")
        self.dash_amount = self.create_dash_card("총 축의금", "0", "원")
        self.dash_avg = self.create_dash_card("평균 금액", "0", "원")
        self.dash_meal = self.create_dash_card("총 식권", "0", "장")
        
        dash_layout = QHBoxLayout()
        dash_layout.addWidget(self.dash_count)
        dash_layout.addWidget(self.dash_amount)
        dash_layout.addWidget(self.dash_avg)
        dash_layout.addWidget(self.dash_meal)
        main_layout.addLayout(dash_layout)

        search_layout = QHBoxLayout()
        search_layout.addStretch()
        search_lbl = QLabel("🔍 이름 검색:")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("이름 입력...")
        self.search_input.setFixedWidth(200)
        self.search_input.textChanged.connect(self.filter_table)
        search_layout.addWidget(search_lbl)
        search_layout.addWidget(self.search_input)
        main_layout.addLayout(search_layout)

        input_group = QGroupBox("신규 등록 (Tab키로 이동, Enter로 저장)")
        input_layout = QHBoxLayout()
        
        self.in_name = QLineEdit()
        self.in_name.setPlaceholderText("이름")
        self.in_name.setFixedWidth(120)
        
        self.in_amount = QLineEdit()
        self.in_amount.setPlaceholderText("금액")
        self.in_amount.setFixedWidth(150)
        
        self.in_meal = QLineEdit()
        self.in_meal.setPlaceholderText("식권")
        self.in_meal.setFixedWidth(120)
        
        self.in_cat = QComboBox()
        self.in_cat.addItems(["친구", "친척", "직장", "가족", "지인", "기타"])
        
        self.in_note = QLineEdit()
        self.in_note.setPlaceholderText("비고 (선택)")

        btn_add = QPushButton("등록 (+)")
        btn_add.clicked.connect(self.add_entry)
        
        self.in_name.returnPressed.connect(self.in_amount.setFocus)
        self.in_amount.returnPressed.connect(self.in_meal.setFocus)
        self.in_meal.returnPressed.connect(self.in_cat.setFocus)
        self.in_note.returnPressed.connect(self.add_entry)

        input_layout.addWidget(QLabel("이름:"))
        input_layout.addWidget(self.in_name)
        input_layout.addWidget(QLabel("금액:"))
        input_layout.addWidget(self.in_amount)
        input_layout.addWidget(QLabel("식권:"))
        input_layout.addWidget(self.in_meal)
        input_layout.addWidget(QLabel("분류:"))
        input_layout.addWidget(self.in_cat)
        input_layout.addWidget(QLabel("비고:"))
        input_layout.addWidget(self.in_note)
        input_layout.addWidget(btn_add)
        
        input_group.setLayout(input_layout)
        main_layout.addWidget(input_group)

        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels(["No", "봉투번호", "이름", "금액", "식권", "분류", "비고", "등록시간"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.doubleClicked.connect(self.edit_entry_popup)
        main_layout.addWidget(self.table)

        btn_layout = QHBoxLayout()
        btn_del = QPushButton("선택 삭제")
        btn_del.setStyleSheet("background-color: #E57373; color: white;")
        btn_del.clicked.connect(self.delete_entry)
        
        btn_reset = QPushButton("전체 초기화")
        btn_reset.setStyleSheet("background-color: #C62828; color: white;")
        btn_reset.clicked.connect(self.reset_all)
        
        btn_excel = QPushButton("엑셀 내보내기")
        btn_excel.setStyleSheet("background-color: #2E7D32; color: white;")
        btn_excel.clicked.connect(self.export_excel)
        
        btn_layout.addWidget(btn_del)
        btn_layout.addWidget(btn_reset)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_excel)
        main_layout.addLayout(btn_layout)

        self.in_name.setFocus()

    def create_dash_card(self, title, val, unit):
        box = QGroupBox()
        box.setStyleSheet("background-color: #EFEBE9; border: none;")
        l = QVBoxLayout()
        t = QLabel(title)
        t.setAlignment(Qt.AlignmentFlag.AlignCenter)
        v = QLabel(f"{val} {unit}")
        v.setAlignment(Qt.AlignmentFlag.AlignCenter)
        v.setStyleSheet("font-size: 20px; font-weight: bold; color: #3E2723;")
        l.addWidget(t)
        l.addWidget(v)
        box.setLayout(l)
        box.obj_val = v
        return box

    def add_entry(self):
        name = self.in_name.text().strip()
        raw_amount = self.in_amount.text().strip().replace(',', '')
        raw_meal = self.in_meal.text().strip()
        category = self.in_cat.currentText()
        note = self.in_note.text().strip()

        if not name or not raw_amount:
            QMessageBox.warning(self, "입력 오류", "이름과 금액을 입력해주세요.")
            return

        try:
            amount = int(raw_amount)
            if 1 <= amount <= 1000:
                amount = amount * 10000
        except ValueError:
            QMessageBox.warning(self, "입력 오류", "금액은 숫자만 입력 가능합니다.")
            return

        try:
            meal = int(raw_meal) if raw_meal else 0
        except ValueError:
            QMessageBox.warning(self, "입력 오류", "식권은 숫자만 입력하세요.")
            return

        if self.db.check_duplicate_name(name):
            reply = QMessageBox.question(self, "중복 확인", 
                                         f"'{name}'님은 이미 등록된 이름입니다.\n그래도 등록하시겠습니까?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No:
                return

        self.db.insert_record(name, amount, meal, category, note)
        self.in_name.clear()
        self.in_amount.clear()
        self.in_meal.clear()
        self.in_note.clear()
        self.in_name.setFocus()
        self.load_data()

    def load_data(self):
        records = self.db.fetch_all()
        self.table.setRowCount(0)
        for row_idx, data in enumerate(records):
            self.table.insertRow(row_idx)
            self.table.setItem(row_idx, 0, QTableWidgetItem(str(len(records) - row_idx)))
            self.table.setItem(row_idx, 1, QTableWidgetItem(str(data[0])))
            self.table.setItem(row_idx, 2, QTableWidgetItem(str(data[1])))
            self.table.setItem(row_idx, 3, QTableWidgetItem(f"{data[2]:,}"))
            self.table.setItem(row_idx, 4, QTableWidgetItem(str(data[3])))
            self.table.setItem(row_idx, 5, QTableWidgetItem(str(data[4])))
            self.table.setItem(row_idx, 6, QTableWidgetItem(str(data[5])))
            self.table.setItem(row_idx, 7, QTableWidgetItem(str(data[6])))
            self.table.item(row_idx, 0).setData(Qt.ItemDataRole.UserRole, data[0])
        self.update_dashboard()

    def update_dashboard(self):
        count, total_amt, total_meal = self.db.get_summary()
        avg = total_amt // count if count > 0 else 0
        self.dash_count.obj_val.setText(f"{count} 명")
        self.dash_amount.obj_val.setText(f"{total_amt:,} 원")
        self.dash_avg.obj_val.setText(f"{avg:,} 원")
        self.dash_meal.obj_val.setText(f"{total_meal} 장")

    def edit_entry_popup(self):
        row = self.table.currentRow()
        if row < 0: return
        record_id = self.table.item(row, 0).data(Qt.ItemDataRole.UserRole)
        curr_data = (
            record_id,
            self.table.item(row, 2).text(),
            self.table.item(row, 3).text(),
            self.table.item(row, 4).text(),
            self.table.item(row, 5).text(),
            self.table.item(row, 6).text(),
            ""
        )
        dlg = EditDialog(curr_data, self)
        if dlg.exec():
            new_name, new_amt, new_meal, new_cat, new_note = dlg.get_data()
            self.db.update_record(record_id, new_name, new_amt, new_meal, new_cat, new_note)
            self.load_data()
            QMessageBox.information(self, "수정 완료", "내역이 수정되었습니다.")

    def delete_entry(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "선택", "삭제할 항목을 선택해주세요.")
            return
        name = self.table.item(row, 2).text()
        reply = QMessageBox.question(self, "삭제 확인", f"정말 '{name}'님의 내역을 삭제하시겠습니까?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            record_id = self.table.item(row, 0).data(Qt.ItemDataRole.UserRole)
            self.db.delete_record(record_id)
            self.load_data()

    def reset_all(self):
        reply = QMessageBox.critical(self, "전체 초기화", "모든 데이터가 영구적으로 삭제됩니다.\n정말 초기화하시겠습니까?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.db.delete_all_records()
            self.load_data()
            QMessageBox.information(self, "완료", "모든 데이터가 초기화되었습니다.")

    def filter_table(self, text):
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 2)
            self.table.setRowHidden(i, text not in item.text())

    def export_excel(self):
        records = self.db.fetch_all()
        if not records:
            QMessageBox.warning(self, "알림", "저장할 데이터가 없습니다.")
            return

        filename = f"축의금정산_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

        try:
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "전체내역"
            headers = ["No", "봉투번호", "이름", "금액", "식권", "분류", "비고", "등록시간"]
            ws1.append(headers)
            
            header_font = Font(bold=True)
            for cell in ws1[1]:
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center')

            total_count = 0
            total_amount = 0
            total_meal = 0
            category_stats = {}

            for idx, row in enumerate(records):
                excel_row = [len(records) - idx, row[0], row[1], row[2], row[3], row[4], row[5], row[6]]
                ws1.append(excel_row)
                total_count += 1
                total_amount += row[2]
                total_meal += row[3]
                cat = row[4]
                category_stats[cat] = category_stats.get(cat, {'count':0, 'amount':0})
                category_stats[cat]['count'] += 1
                category_stats[cat]['amount'] += row[2]

            for col in ws1.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                ws1.column_dimensions[column].width = (max_length + 2) * 1.2

            ws2 = wb.create_sheet("요약보고서")
            ws2.append(["구분", "인원수", "총 금액"])
            ws2.append(["전체 합계", total_count, total_amount])
            ws2.append(["총 식권", total_meal, "-"])
            ws2.append(["", "", ""])
            ws2.append(["[분류별 상세]", "", ""])
            for cat, stat in category_stats.items():
                ws2.append([cat, stat['count'], stat['amount']])
            for cell in ws2[1]:
                cell.font = header_font

            wb.save(filename)
            QMessageBox.information(self, "성공", f"파일이 저장되었습니다.\n{filename}")
            
        except Exception as e:
            QMessageBox.critical(self, "오류", f"저장 중 오류가 발생했습니다.\n{str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = WeddingManager()
    ex.show()
    sys.exit(app.exec())