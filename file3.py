import pandas as pd
import streamlit as st
from io import BytesIO

def process(uploaded_file1, uploaded_file3, new_data):
    # file1（GSHISAN_DATA.CSV）を読み込む
    file1_io = BytesIO(uploaded_file1.getvalue())
    df_file1 = pd.read_csv(file1_io)

    # file3（GBUMON.CSV）を読み込む（ヘッダーなし）
    file3_io = BytesIO(uploaded_file3.getvalue())
    df_file3 = pd.read_csv(file3_io, header=None)

    # file1の「部門コード」とfile3の0列目を照合
    for index, row in df_file1.iterrows():
        department_code = row['部門コード']
        matched_row = df_file3[df_file3[0] == department_code]

        if not matched_row.empty:
            # 一致する場合の処理：file3の1列目の文字列をnew_dataの「部門」に設定
            new_data.at[index, '部門'] = matched_row.iloc[0, 1]

    return new_data