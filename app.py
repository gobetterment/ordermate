# 📦 PyQt5 기반 OrderMate GUI (기능 동일)
# 기능: 판매데이터, 재고데이터, 상품코드 업로드 → 발주수량 입력 → 거래처별 엑셀 추출 (ZIP)

import sys
import os
import pandas as pd
import zipfile
import datetime
from io import BytesIO
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QLabel,
    QHBoxLayout, QHeaderView, QTabWidget
)
from PyQt5.QtCore import Qt

class OrderMateApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("OrderMate - 발주도우미 (PyQt5)")
        self.setGeometry(100, 100, 1400, 800)

        # 데이터 저장
        self.sales_data = None
        self.inventory_data = None
        self.product_data = None
        self.final_data = None

        # UI 구성
        self.tabs = QTabWidget()
        self.dashboard_tab = QWidget()
        self.order_tab = QWidget()

        self.tabs.addTab(self.dashboard_tab, "📊 대시보드")
        self.tabs.addTab(self.order_tab, "📝 발주서 작성")

        self.setCentralWidget(self.tabs)

        self.init_dashboard_ui()
        self.init_order_ui()

    # ---------------------- 대시보드 UI ----------------------
    def init_dashboard_ui(self):
        layout = QVBoxLayout()

        load_sales_btn = QPushButton("판매 데이터 불러오기")
        load_sales_btn.clicked.connect(self.load_sales_data)
        layout.addWidget(load_sales_btn)

        load_inventory_btn = QPushButton("총재고 데이터 불러오기")
        load_inventory_btn.clicked.connect(self.load_inventory_data)
        layout.addWidget(load_inventory_btn)

        self.dashboard_table = QTableWidget()
        layout.addWidget(self.dashboard_table)

        self.dashboard_tab.setLayout(layout)

    def load_sales_data(self):
        path, _ = QFileDialog.getOpenFileName(self, "판매 데이터 선택", "", "Excel (*.xlsx);;CSV (*.csv)")
        if path:
            self.sales_data = pd.read_excel(path) if path.endswith('.xlsx') else pd.read_csv(path)
            QMessageBox.information(self, "성공", "판매 데이터 불러오기 완료")
            self.update_dashboard()

    def load_inventory_data(self):
        path, _ = QFileDialog.getOpenFileName(self, "재고 데이터 선택", "", "Excel (*.xlsx);;CSV (*.csv)")
        if path:
            self.inventory_data = pd.read_excel(path) if path.endswith('.xlsx') else pd.read_csv(path)
            QMessageBox.information(self, "성공", "재고 데이터 불러오기 완료")

    def update_dashboard(self):
        if self.sales_data is None:
            return
        df = self.sales_data.groupby(['상품구분', '브랜드'])[['판매수량', '판매금액']].sum().reset_index()
        self.show_table(df, self.dashboard_table)

    # ---------------------- 발주서 작성 UI ----------------------
    def init_order_ui(self):
        layout = QVBoxLayout()

        load_codes_btn = QPushButton("상품코드 리스트 불러오기")
        load_codes_btn.clicked.connect(self.load_product_codes)
        layout.addWidget(load_codes_btn)

        export_btn = QPushButton("거래처별 발주서 ZIP 저장")
        export_btn.clicked.connect(self.export_zip)
        layout.addWidget(export_btn)

        self.order_table = QTableWidget()
        layout.addWidget(self.order_table)

        self.order_tab.setLayout(layout)

    def load_product_codes(self):
        path, _ = QFileDialog.getOpenFileName(self, "상품코드 리스트 선택", "", "Excel (*.xlsx);;CSV (*.csv)")
        if path:
            self.product_data = pd.read_excel(path) if path.endswith('.xlsx') else pd.read_csv(path)
            self.process_order_data()
            QMessageBox.information(self, "성공", "상품코드 리스트 불러오기 완료")
            self.show_table(self.final_data, self.order_table)

    def process_order_data(self):
        df = self.product_data.copy()
        df['발주수량'] = 0
        df['판매수량'] = 0
        for 점포 in ['본사', '홍대점', '평대점', '협재점', '대청호점']:
            df[점포] = 0

        if self.inventory_data is not None:
            for i, row in df.iterrows():
                자사코드 = row['자사코드']
                for 지점 in ['본사', '홍대점', '평대점', '협재점', '대청호점']:
                    재고 = self.inventory_data.loc[
                        (self.inventory_data['자사코드'] == 자사코드) & (self.inventory_data['창고/매장명'] == 지점), '재고']
                    df.at[i, 지점] = 재고.values[0] if not 재고.empty else 0

        if self.sales_data is not None:
            for i, row in df.iterrows():
                판매 = self.sales_data[self.sales_data['자사바코드'] == row['자사코드']]['판매수량'].sum()
                df.at[i, '판매수량'] = 판매

        df['재고합계'] = df[['본사', '홍대점', '평대점', '협재점', '대청호점']].sum(axis=1)
        self.final_data = df

    def export_zip(self):
        if self.final_data is None:
            QMessageBox.warning(self, "오류", "상품코드를 먼저 불러와 주세요")
            return

        df = self.final_data.copy()
        df = df[df['발주수량'] > 0]
        df['공급가합계'] = df['발주수량'] * df['공급가']

        path, _ = QFileDialog.getSaveFileName(self, "ZIP 저장 위치", "", "ZIP (*.zip)")
        if not path:
            return

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            now_str = datetime.datetime.now().strftime('%y%m%d')
            for 거래처 in df['거래처명'].unique():
                거래처별 = df[df['거래처명'] == 거래처][['거래처명', '자사코드', '상품코드', '상품명', '칼라명', '사이즈명', '발주수량', '공급가', '공급가합계']]
                거래처별.reset_index(drop=True, inplace=True)
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                    거래처별.to_excel(writer, index=False, sheet_name='발주서')
                zipf.writestr(f"(홀라인){거래처}_발주서_{now_str}.xlsx", excel_buffer.getvalue())

        with open(path, "wb") as f:
            f.write(zip_buffer.getvalue())
        QMessageBox.information(self, "완료", "발주서 ZIP 저장이 완료되었습니다")

    def show_table(self, df, table_widget):
        table_widget.setRowCount(0)
        table_widget.setColumnCount(0)
        table_widget.setColumnCount(len(df.columns))
        table_widget.setHorizontalHeaderLabels(df.columns.tolist())
        table_widget.setRowCount(len(df))

        for i in range(len(df)):
            for j in range(len(df.columns)):
                item = QTableWidgetItem(str(df.iat[i, j]))
                item.setTextAlignment(Qt.AlignCenter)
                table_widget.setItem(i, j, item)

        table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = OrderMateApp()
    window.show()
    sys.exit(app.exec_())
