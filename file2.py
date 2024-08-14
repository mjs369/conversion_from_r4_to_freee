import pandas as pd
import streamlit as st
from io import BytesIO

def process(uploaded_file1, uploaded_file2, new_data):
    # file1（GSHISAN_DATA.CSV）を読み込む
    file1_io = BytesIO(uploaded_file1.getvalue())
    df_file1 = pd.read_csv(file1_io)

    # file2（GKANJO.CSV）を読み込む（ヘッダーなし）
    file2_io = BytesIO(uploaded_file2.getvalue())
    df_file2 = pd.read_csv(file2_io, header=None)

    # 勘定科目のマッピング
    account_mapping = {
        11: "建物", 21: "付属設備", 31: "構築物", 41: "機械装置", 42: "機械装置",
        51: "車両運搬具", 61: "工具器具備品", 62: "工具器具備品", 65: "工具器具備品",
        66: "工具器具備品", 67: "一括償却資産", 71: "船舶", 72: "航空機",
        73: "土地", 74: "建設仮勘定", 81: "ソフトウェア", 82: "工業所有権", 89: "営業権"
    }

    # file1の「勘定コード」とfile2のA列（0行目）を照合
    for index, row in df_file1.iterrows():
        account_code = row['勘定コード']
        matched_row = df_file2[df_file2[0] == account_code]

        if not matched_row.empty:
            account_value = matched_row.iloc[0, 0]
            if account_value in account_mapping:
                # 一致する場合の処理
                new_data.at[index, '勘定科目'] = account_mapping[account_value]
            else:
                # 一致しない場合のデフォルト値
                new_data.at[index, '勘定科目'] = "開業費"
        else:
            # 見つからない場合のデフォルト値
            new_data.at[index, '勘定科目'] = "開業費"

    return new_data
