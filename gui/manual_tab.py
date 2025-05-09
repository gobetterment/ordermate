from PyQt5.QtWidgets import QWidget, QVBoxLayout, QTextBrowser
import os

class ManualTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout()
        manual_browser = QTextBrowser()
        manual_browser.setOpenExternalLinks(True)

        # img 폴더의 경로를 절대경로로 변환
        img_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "../img"))

        html = f"""
        <h1>📝 필수 조건 설정 가이드</h1>
        <h2>데이터 엑셀 다운로드시 조건</h2>
        <h3>1. 판매 데이터 (메뉴명: 매장판매일보)</h3><br>
        <img src=\"file:///{img_dir}/1.png\"><br><br>
        <p>① 판매기간 설정</p>
        <p>② 매장검색조건에서 홍대점,평대점,협재점,대청호점,자사몰,타플랫폼,직원구매로 설정</p>
        <p>③ 기타검색조건에서 매장소계보기 체크해제</p>
        <p>④ 브랜드,상품구분 체크 / 자사바코드 체크해제 (자사바코드는 xmd의 버그로 체크를 해제해야 보임)</p>
        <p>⑤ 조회 후 조회 버튼 옆 엑셀 버튼 눌러 엑셀 다운로드</p><br><br>

        <h3>2. 총재고현황 데이터 (메뉴명: 총재고현황)</h3><br>
        <img src=\"file:///{img_dir}/2.png\"><br><br>
        <p>① 기준일자는 현재일로 둘 다 맞춰줘야함</p>
        <p>② 기타검색조건에서 모두 체크 해제 필수</p>
        <p>③ 상품검색조건/매장검색조건/창고검색조건은 캡쳐 이미지와 같이 설정</p>
        <p>④ 조회 후 조회 버튼 옆 엑셀 버튼 눌러 엑셀 다운로드</p><br><br>

        <h3>3. 총재고현황 엑셀 데이터 수정 후 저장 방법</h3><br>
        <img src=\"file:///{img_dir}/3.png\">
        <img src=\"file:///{img_dir}/4.png\"><br><br>
        <p>①-② 엑셀파일을 열고 두번째 줄 오른쪽 클릭하여 삭제</p>
        <p>③-④ O열부터 W열 까지 컬럼 선택 후 오른쪽 클릭하여 삭제 후 파일저장</p>
        <p>(O열 클릭 후 시프트 키 누른 상태로 W열 클릭)</p><br><br>

        <h3>4. 상품코드리스트 데이터 다운로드시 조건 (메뉴명:상품코드리스트)</h3><br>
        <img src=\"file:///{img_dir}/5.png\"><br><br>
        ① 검색조건에서 발주 할 브랜드만 체크
        <p>② 조회필터에 브랜드 체크 필수</p>
        <p>③ 조회 후 조회 버튼 옆 엑셀 버튼 눌러 엑셀 다운로드</p><br><br>
        """

        manual_browser.setHtml(html)
        layout.addWidget(manual_browser)
        self.setLayout(layout) 