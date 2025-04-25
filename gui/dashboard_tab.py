from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QFileDialog, QMessageBox, QTableView, QComboBox, QHBoxLayout, QLabel
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QHeaderView
from PyQt5.QtGui import QStandardItemModel, QStandardItem
import pandas as pd

class DashboardTab(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.model = QStandardItemModel()
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # 데이터 로드 버튼
        load_sales_btn = QPushButton("판매 데이터 불러오기")
        load_sales_btn.clicked.connect(self.load_sales_data)
        layout.addWidget(load_sales_btn)

        load_inventory_btn = QPushButton("총재고 데이터 불러오기")
        load_inventory_btn.clicked.connect(self.load_inventory_data)
        layout.addWidget(load_inventory_btn)

        # 필터 레이아웃
        filter_layout = QHBoxLayout()
        filter_label = QLabel("상품구분 필터:")
        self.filter_combo = QComboBox()
        self.filter_combo.currentTextChanged.connect(self.update_dashboard)
        filter_layout.addWidget(filter_label)
        filter_layout.addWidget(self.filter_combo)
        layout.addLayout(filter_layout)

        # 테이블 뷰 설정
        self.dashboard_table = QTableView()
        self.model = QStandardItemModel()
        # 정렬 기준을 UserRole로 설정
        self.model.setSortRole(Qt.UserRole)
        self.dashboard_table.setModel(self.model)
        self.dashboard_table.setSortingEnabled(True)
        self.dashboard_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.dashboard_table)

        self.setLayout(layout)

    def load_sales_data(self):
        path, _ = QFileDialog.getOpenFileName(self, "판매 데이터 선택", "", "Excel (*.xlsx);;CSV (*.csv)")
        if path:
            self.parent.sales_data = pd.read_excel(path) if path.endswith('.xlsx') else pd.read_csv(path)
            # NaN 데이터 제거
            self.parent.sales_data = self.parent.sales_data.dropna(subset=['상품구분', '브랜드'])
            QMessageBox.information(self, "성공", "판매 데이터 불러오기 완료")
            self.update_filter_combo()
            self.update_dashboard()

    def load_inventory_data(self):
        path, _ = QFileDialog.getOpenFileName(self, "재고 데이터 선택", "", "Excel (*.xlsx);;CSV (*.csv)")
        if path:
            self.parent.inventory_data = pd.read_excel(path) if path.endswith('.xlsx') else pd.read_csv(path)
            QMessageBox.information(self, "성공", "재고 데이터 불러오기 완료")

    def update_filter_combo(self):
        if self.parent.sales_data is not None:
            self.filter_combo.clear()
            self.filter_combo.addItem("전체")
            # 상품구분을 문자열로 변환하여 정렬
            categories = sorted(self.parent.sales_data['상품구분'].astype(str).unique())
            self.filter_combo.addItems(categories)

    def update_dashboard(self):
        if self.parent.sales_data is None:
            return

        selected_category = self.filter_combo.currentText()
        df = self.parent.sales_data.copy()

        # 선택된 상품구분으로 필터링 (문자열로 비교)
        if selected_category != "전체":
            df = df[df['상품구분'].astype(str) == selected_category]

        df = df.groupby(['상품구분', '브랜드'])[['판매수량', '판매금액']].sum().reset_index()
        # 판매금액 높은 순으로 정렬
        df = df.sort_values('판매금액', ascending=False)
        # 표시용 데이터 포맷팅
        df['판매수량'] = df['판매수량'].apply(lambda x: f"{x:,.0f}")
        df['판매금액'] = df['판매금액'].apply(lambda x: f"{x:,.0f}")

        self.show_table(df)

    def show_table(self, df):
        self.model.clear()
        self.model.setHorizontalHeaderLabels(['상품구분', '브랜드', '판매수량', '판매금액'])

        for i in range(len(df)):
            row = []
            for j, col in enumerate(['상품구분', '브랜드', '판매수량', '판매금액']):
                item = QStandardItem(str(df.iat[i, df.columns.get_loc(col)]))
                item.setTextAlignment(Qt.AlignCenter)
                # 정렬을 위한 원본 데이터 설정
                if col in ['판매수량', '판매금액']:
                    # 천단위 구분자 제거 후 float로 변환
                    value = float(str(df.iat[i, df.columns.get_loc(col)]).replace(',', ''))
                    item.setData(value, Qt.UserRole)
                row.append(item)
            self.model.appendRow(row) 