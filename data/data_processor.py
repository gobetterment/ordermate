import pandas as pd
import zipfile
import datetime
from io import BytesIO
import os
from PyQt5.QtWidgets import QMessageBox

class DataProcessor:
    def process_order_data(self, product_data, inventory_data, sales_data):
        # 상품코드리스트와 재고데이터 병합
        df = product_data.copy()
        
        # 재고데이터 처리
        if inventory_data is not None:
            for 지점 in ['본사', '홍대점', '평대점', '협재점', '대청호점']:
                df[지점] = 0
                for i, row in df.iterrows():
                    자사코드 = row['자사코드']
                    재고 = inventory_data.loc[
                        (inventory_data['자사코드'] == 자사코드) & 
                        (inventory_data['창고/매장명'] == 지점), '재고']
                    df.at[i, 지점] = 재고.values[0] if not 재고.empty else 0
        
        # 판매데이터 처리
        if sales_data is not None:
            sales_data['자사바코드'] = sales_data['자사바코드'].astype(str)
            df['자사코드'] = df['자사코드'].astype(str)
            
            # 판매수량 합계 계산
            sales_sum = sales_data.groupby('자사바코드')['판매수량'].sum().reset_index()
            sales_sum = sales_sum.rename(columns={'자사바코드': '자사코드'})
            
            # 판매수량 병합
            df = pd.merge(df, sales_sum, on='자사코드', how='left')
            df['판매수량'] = df['판매수량'].fillna(0)
        
        # 발주수량과 공급가 컬럼 추가
        df['발주수량'] = 0
        df['공급가'] = df['사전원가']  # 사전원가를 공급가로 사용
        df['공급가합계'] = df['발주수량'] * df['공급가']
        
        # 컬럼명 변경
        df = df.rename(columns={
            '본사': '본사재고',
            '홍대점': '홍대점재고',
            '평대점': '평대점재고',
            '협재점': '협재점재고',
            '대청호점': '대청호점재고'
        })
        
        return df

    def export_zip(self, final_data, path):
        df = final_data.copy()
        df = df[df['발주수량'] > 0]
        df['공급가합계'] = df['발주수량'] * df['공급가']

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

    def export_excel_files(self, data, save_path):
        """발주서를 개별 엑셀 파일로 저장"""
        try:
            # print(f"저장 시작: 총 {len(data)}개 데이터")
            
            # 발주수량이 0보다 큰 데이터만 필터링
            data = data[data['발주수량'] > 0]
            # print(f"발주수량 > 0 필터링 후: {len(data)}개 데이터")
            
            if len(data) == 0:
                QMessageBox.warning(self, "오류", f"저장할 데이터가 없습니다.")
                # print("저장할 데이터가 없습니다.")
                return False
            
            # 거래처별로 데이터 그룹화
            grouped_data = data.groupby('거래처명')
            # print(f"거래처 수: {len(grouped_data)}")
            
            # 현재 날짜 가져오기
            current_date = datetime.datetime.now().strftime('%y%m%d')
            
            # 각 거래처별로 엑셀 파일 생성
            for 거래처명, group_data in grouped_data:
                # print(f"\n거래처 '{거래처명}' 처리 중...")
                # print(f"데이터 수: {len(group_data)}")
                
                # 필요한 컬럼만 선택
                export_data = group_data[['거래처명', '자사코드', '상품코드', '상품명', '칼라명', '사이즈명', '발주수량', '공급가', '공급가합계', 'TAG가']]
                
                # 파일명 생성
                filename = f"(홀라인){거래처명}_발주서_{current_date}.xlsx"
                file_path = os.path.join(save_path, filename)
                # print(f"저장 경로: {file_path}")
                
                # 엑셀 파일로 저장 (xlsxwriter 엔진 사용)
                with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                    export_data.to_excel(writer, index=False, sheet_name='발주서')
                    worksheet = writer.sheets['발주서']
                    
                    # 컬럼 너비 설정
                    worksheet.set_column('A:A', 13)  # 거래처명
                    worksheet.set_column('B:B', 15)  # 자사코드
                    worksheet.set_column('C:C', 10)  # 상품코드
                    worksheet.set_column('D:D', 30)  # 상품명
                    worksheet.set_column('E:E', 10)  # 칼라명
                    worksheet.set_column('F:F', 7)  # 사이즈명
                    worksheet.set_column('G:G', 7)  # 발주수량
                    worksheet.set_column('H:H', 10)  # 공급가
                    worksheet.set_column('I:I', 12)  # 공급가합계
                    worksheet.set_column('J:J', 10)  # TAG가

                    # 헤더 스타일 설정
                    header_format = writer.book.add_format({
                        'bold': True,
                        'bg_color': '#D3D3D3',
                        'align': 'center',
                        'valign': 'vcenter',
                        'border': 1,
                        'font_size': 9
                    })
                    # 숫자 형식 및 테두리
                    number_format = writer.book.add_format({
                        'num_format': '#,##0',
                        'align': 'right',
                        'border': 1,
                        'font_size': 9
                    })
                    # 일반 셀 테두리
                    border_format = writer.book.add_format({'border': 1, 'font_size': 9})
                    
                    # 헤더 행 스타일 적용
                    for col_num, value in enumerate(export_data.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                    
                    
                    
                    # 데이터 행에 스타일 적용 및 셀 테두리 적용
                    for row_num in range(len(export_data)):
                        for col_num in range(len(export_data.columns)):
                            value = export_data.iloc[row_num, col_num]
                            # 공급가, 공급가합계는 숫자 포맷
                            if col_num in [7, 8]:
                                worksheet.write(row_num + 1, col_num, value, number_format)
                            else:
                                worksheet.write(row_num + 1, col_num, value, border_format)
                        # 공급가합계 수식 설정
                        worksheet.write_formula(row_num + 1, 8, f'=G{row_num + 2}*H{row_num + 2}', number_format)
                    
                    # === 하단 합계 수식 ===
                    total_row = len(export_data) + 1
                    # 테두리 없는 숫자 포맷
                    noborder_number_format = writer.book.add_format({'num_format': '#,##0', 'align': 'right', 'font_size': 9})
                    worksheet.write_formula(total_row, 6, f'SUM(G2:G{total_row})', noborder_number_format)  # 발주수량 합계
                    worksheet.write_formula(total_row, 8, f'SUM(I2:I{total_row})', noborder_number_format)  # 공급가합계 합계
                # print(f"파일 저장 완료: {filename}")
            QMessageBox.information(None, "완료", "모든 파일 저장이 완료되었습니다.")
            return True
        except Exception as e:
            QMessageBox.warning(None, "오류", f"엑셀 파일 저장 중 오류 발생: {str(e)}")
            return False 