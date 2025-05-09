import sys
from PyQt5.QtWidgets import QMainWindow, QTabWidget
from .dashboard_tab import DashboardTab
from .order_tab import OrderTab
from .manual_tab import ManualTab

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
        self.dashboard_tab = DashboardTab(self)
        self.order_tab = OrderTab(self)
        self.manual_tab = ManualTab(self)

        self.tabs.addTab(self.dashboard_tab, "📊 대시보드")
        self.tabs.addTab(self.order_tab, "📝 발주서 작성")
        self.tabs.addTab(self.manual_tab, "🧑‍🎓 매뉴얼")

        self.setCentralWidget(self.tabs) 