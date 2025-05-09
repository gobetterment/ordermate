from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QFileDialog, QMessageBox, QTableView
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QHeaderView
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QBrush, QColor
import pandas as pd
from data.data_processor import DataProcessor

class OrderTab(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.data_processor = DataProcessor()
        self.model = QStandardItemModel()
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        load_codes_btn = QPushButton("상품코드리스트 불러오기")
        load_codes_btn.clicked.connect(self.load_product_codes)
        layout.addWidget(load_codes_btn)

        export_btn = QPushButton("거래처별 발주서 파일 생성")
        export_btn.clicked.connect(self.export_zip)
        layout.addWidget(export_btn)

        self.order_table = QTableView()
        self.order_table.setModel(self.model)
        self.order_table.setSortingEnabled(True)
        
        # 컬럼 너비 설정
        header = self.order_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)  # 수동 조정 가능하도록 설정
        
        # 기본 너비 설정
        self.order_table.setColumnWidth(0, 100)  # 거래처명
        self.order_table.setColumnWidth(1, 100)  # 자사코드
        self.order_table.setColumnWidth(2, 100)  # 상품코드
        self.order_table.setColumnWidth(3, 300)  # 상품명 (더 넓게)
        self.order_table.setColumnWidth(4, 100)  # 칼라명
        self.order_table.setColumnWidth(5, 80)   # 사이즈명
        self.order_table.setColumnWidth(6, 80)   # 발주수량
        self.order_table.setColumnWidth(7, 80)   # 판매수량
        self.order_table.setColumnWidth(8, 80)   # 본사
        self.order_table.setColumnWidth(9, 80)   # 홍대점
        self.order_table.setColumnWidth(10, 80)  # 평대점
        self.order_table.setColumnWidth(11, 80)  # 협재점
        self.order_table.setColumnWidth(12, 80)  # 대청호점
        self.order_table.setColumnWidth(13, 80)  # TAG가
        self.order_table.setColumnWidth(14, 80)  # 공급가
        self.order_table.setColumnWidth(15, 100) # 공급가합계
        
        layout.addWidget(self.order_table)

        # 발주수량 변경 시 공급가합계 업데이트
        self.model.itemChanged.connect(self.on_item_changed)

        self.setLayout(layout)

    def on_item_changed(self, item):
        # 발주수량 컬럼의 인덱스 찾기
        row = item.row()
        col = item.column()
        columns = [
            '거래처명', '자사코드', '상품코드', '상품명', '칼라명', 
            '사이즈명', '발주수량', '판매수량', '본사재고', '홍대점재고', 
            '평대점재고', '협재점재고', '대청호점재고', 'TAG가', '공급가', '공급가합계'
        ]
        
        if columns[col] == '발주수량':
            try:
                # 발주수량과 공급가 가져오기
                order_qty = float(str(item.text()).replace(',', ''))
                supply_price = float(str(self.model.item(row, columns.index('공급가')).text()).replace(',', ''))
                
                # 공급가합계 계산
                total = order_qty * supply_price
                
                # 공급가합계 업데이트
                total_item = self.model.item(row, columns.index('공급가합계'))
                total_item.setText(f"{total:,.0f}")
                total_item.setData(total, Qt.UserRole)

                # === 데이터프레임에도 값 반영 ===
                code = self.model.item(row, columns.index('자사코드')).text()
                size = self.model.item(row, columns.index('사이즈명')).text()
                # 자사코드와 사이즈명으로 행 찾기
                mask = (self.parent.final_data['자사코드'] == code) & (self.parent.final_data['사이즈명'] == size)
                self.parent.final_data.loc[mask, '발주수량'] = order_qty
                self.parent.final_data.loc[mask, '공급가합계'] = total
            except (ValueError, AttributeError, KeyError, IndexError) as e:
                QMessageBox.warning(self, "오류", f"발주수량 동기화 중 오류가 발생했습니다: {e}")

    def load_product_codes(self):
        path, _ = QFileDialog.getOpenFileName(self, "상품코드 리스트 선택", "", "Excel (*.xlsx);;CSV (*.csv)")
        if path:
            try:
                if path.endswith('.xlsx'):
                    self.parent.product_data = pd.read_excel(path)
                else:
                    self.parent.product_data = pd.read_csv(path)
                self.process_order_data()
                QMessageBox.information(self, "성공", "상품코드 리스트 불러오기 완료")
                self.show_table(self.parent.final_data)
            except Exception as e:
                QMessageBox.critical(self, "파일 오류", f"파일을 불러오는 중 오류가 발생했습니다.\n\n{str(e)}\n\n다시 시도해 주세요.")
                self.parent.product_data = None

    def process_order_data(self):
        self.parent.final_data = self.data_processor.process_order_data(
            self.parent.product_data,
            self.parent.inventory_data,
            self.parent.sales_data
        )

    def export_zip(self):
        if self.parent.final_data is None:
            QMessageBox.warning(self, "오류", "상품코드를 먼저 불러와 주세요")
            return

        # 발주수량이 입력된 데이터만 필터링
        filtered_data = self.parent.final_data[self.parent.final_data['발주수량'] > 0]
        # print(f"\n발주서 저장 시작")
        # print(f"전체 데이터 수: {len(self.parent.final_data)}")
        # print(f"발주수량 > 0 데이터 수: {len(filtered_data)}")
        
        if len(filtered_data) == 0:
            QMessageBox.warning(self, "오류", "발주수량을 입력한 상품이 없습니다")
            return

        # 폴더 선택 다이얼로그
        save_path = QFileDialog.getExistingDirectory(self, "발주서 저장 위치")
        if not save_path:
            return

        # print(f"저장 경로: {save_path}")

        # 엑셀 파일로 저장
        if self.data_processor.export_excel_files(filtered_data, save_path):
            QMessageBox.information(self, "완료", "발주서 엑셀 파일 저장이 완료되었습니다")
        else:
            QMessageBox.warning(self, "오류", "발주서 저장 중 오류가 발생했습니다")

    def show_table(self, df):
        self.model.clear()
        
        # 지정된 컬럼 순서대로 설정
        columns = [
            '거래처명', '자사코드', '상품코드', '상품명', '칼라명', 
            '사이즈명', '발주수량', '판매수량', '본사재고', '홍대점재고', 
            '평대점재고', '협재점재고', '대청호점재고', 'TAG가', '공급가', '공급가합계'
        ]
        self.model.setHorizontalHeaderLabels(columns)
        
        for i in range(len(df)):
            row = []
            for col in columns:
                try:
                    value = df.iat[i, df.columns.get_loc(col)]
                except (KeyError, IndexError):
                    value = 0
                
                item = QStandardItem(str(value))
                
                # 컬럼별 정렬 설정
                if col in ['거래처명', '자사코드', '상품코드', '상품명', '칼라명', '사이즈명']:
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                elif col in ['TAG가', '공급가', '공급가합계']:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    item.setTextAlignment(Qt.AlignCenter)
                
                # 발주수량 컬럼은 편집 가능하도록 설정
                if col == '발주수량':
                    item.setFlags(item.flags() | Qt.ItemIsEditable)
                
                # 숫자 컬럼의 경우 정렬을 위한 UserRole 설정
                if col in ['발주수량', '판매수량', '본사재고', '홍대점재고', '평대점재고', 
                         '협재점재고', '대청호점재고', 'TAG가', '공급가', '공급가합계']:
                    try:
                        numeric_value = float(str(value).replace(',', ''))
                        item.setData(numeric_value, Qt.UserRole)
                        
                        # 정수로 변환하여 표시 (소수점 제거)
                        if numeric_value.is_integer():
                            item.setText(str(int(numeric_value)))
                        
                        # 판매수량이 1 이상일 때 빨간색 글자 적용
                        if col == '판매수량' and numeric_value >= 1:
                            item.setForeground(QBrush(QColor(255, 0, 0)))  # 빨간색
                            
                        # 본사재고 컬럼이 0일 때 회색 배경 적용
                        if col == '본사재고' and numeric_value == 0:
                            item.setBackground(QBrush(QColor(200, 200, 200)))  # 회색
                            
                        # 본사재고 컬럼의 글자를 진하게 표시
                        if col == '본사재고':
                            font = item.font()
                            font.setBold(True)
                            item.setFont(font)
                    except (ValueError, AttributeError):
                        item.setData(0, Qt.UserRole)
                row.append(item)
            self.model.appendRow(row)
            
        # 컬럼 너비 설정
        header = self.order_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)  # 수동 조정 가능하도록 설정
        
        # 기본 너비 설정
        self.order_table.setColumnWidth(0, 100)  # 거래처명
        self.order_table.setColumnWidth(1, 100)  # 자사코드
        self.order_table.setColumnWidth(2, 100)  # 상품코드
        self.order_table.setColumnWidth(3, 300)  # 상품명 (더 넓게)
        self.order_table.setColumnWidth(4, 100)  # 칼라명
        self.order_table.setColumnWidth(5, 80)   # 사이즈명
        self.order_table.setColumnWidth(6, 80)   # 발주수량
        self.order_table.setColumnWidth(7, 80)   # 판매수량
        self.order_table.setColumnWidth(8, 80)   # 본사
        self.order_table.setColumnWidth(9, 80)   # 홍대점
        self.order_table.setColumnWidth(10, 80)  # 평대점
        self.order_table.setColumnWidth(11, 80)  # 협재점
        self.order_table.setColumnWidth(12, 80)  # 대청호점
        self.order_table.setColumnWidth(13, 80)  # TAG가
        self.order_table.setColumnWidth(14, 80)  # 공급가
        self.order_table.setColumnWidth(15, 100) # 공급가합계

        # # 디버깅을 위한 로그 출력
        # print("\n=== 오더탭 디버깅 로그 ===")
        # print("1. 표시된 데이터 샘플:")
        # for i in range(min(5, len(df))):
        #     print(f"자사코드: {df.iat[i, df.columns.get_loc('자사코드')]}, 판매수량: {df.iat[i, df.columns.get_loc('판매수량')]}")
        # print("===========================\n") 