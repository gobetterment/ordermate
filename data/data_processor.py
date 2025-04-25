import pandas as pd
import zipfile
import datetime
from io import BytesIO

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