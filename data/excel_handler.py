import pandas as pd

class ExcelHandler:
    @staticmethod
    def read_excel(path):
        return pd.read_excel(path)

    @staticmethod
    def read_csv(path):
        return pd.read_csv(path)

    @staticmethod
    def write_excel(df, path, sheet_name='Sheet1'):
        with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name) 