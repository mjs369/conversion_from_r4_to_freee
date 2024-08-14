import streamlit as st
import pandas as pd
import os
from excel_processing_1 import process_excel_data_1  
from excel_processing_2 import process_excel_data_2
from excel_processing_3 import process_excel_data_3
from excel_processing_4 import process_excel_data_4
from excel_processing_5 import process_excel_data_5
from excel_processing_6 import process_excel_data_6
from excel_processing_7 import process_excel_data_7
from excel_processing_8 import process_excel_data_8
from total_display import total_display

def app4():
    st.title("freee部門分割")

    st.markdown("""
    #### freeeからの出力手順

    1. freeeの「レポート」から「月次推移」を選択
    2. 「検索条件」の「表示するタグ」で「部門」を選択し絞り込む
    3. 「エクスポート」→「CSV・PDFエクスポート」を選択
    4. 「CSV」→「項目を整理する/勘定科目コードを含める」どちらもチェックを入れる
    5. 「出力する帳表の選択」で「損益計算書」のみにチェックを入れて出力
    6. 「エクスポート」→「出力した帳表一覧」を選択し該当ファイルをダウンロード
    """)

    uploaded_file = st.file_uploader("ファイルをアップロードしてください", type=['csv'])

    if uploaded_file is not None:
        if st.button('Start'):
            uploaded_filename = uploaded_file.name  # アップロードされたファイルの名前を取得
            df_uploaded = pd.read_csv(uploaded_file, encoding='cp932', header=1)
            dynamic_columns = [col for col in df_uploaded.columns if col not in ('勘定科目コード', 'Unnamed: 1', '部門', '期間累計')]

            excel_headers = ['勘定科目', '決算書表示名'] + dynamic_columns + ['期間累計']
            initial_values = ["売上高", "売上原価", "売上総損益", "販売費及び一般管理費",
                              "営業損益", "営業外利益", "営業外損失", "経常損益", "特別利益",
                              "特別損失", "税引前当期純利益"]
            spacing_pairs = [("売上高", "売上高"), ("売上高", "売上原価"), ("売上総損益", "販売費及び一般管理費"),
                             ("営業損益", "営業外利益"), ("営業外利益", "営業外損失"),
                             ("経常損益", "特別利益"), ("特別利益", "特別損失")]

            data = []
            data.extend([[''] * len(excel_headers)] * 50)

            for i, value in enumerate(initial_values):
                data.append([value] * 2 + [''] * len(dynamic_columns))
                if i < len(initial_values) - 1 and (value, initial_values[i + 1]) in spacing_pairs:
                    data.extend([[''] * len(excel_headers)] * 50)

            df_excel = pd.DataFrame(data, columns=excel_headers)
            excel_filename = '月次推移_損益計算書.xlsx'
            df_excel.to_excel(excel_filename, index=False)

            unique_departments = df_uploaded['部門'].dropna().unique()
            with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
                df_excel.to_excel(writer, index=False, sheet_name='全体')
                for dept in unique_departments:
                    safe_sheet_name = dept[:30]
                    df_excel.to_excel(writer, index=False, sheet_name=safe_sheet_name)

            # excel_processingの関数を呼び出して、Excelファイルの更新を行う
            process_excel_data_1(df_uploaded, excel_filename, dynamic_columns)

            # excel_processingの関数を呼び出して、Excelファイルの更新を行う
            process_excel_data_2(df_uploaded, excel_filename, dynamic_columns)

            # excel_processingの関数を呼び出して、Excelファイルの更新を行う
            process_excel_data_3(df_uploaded, excel_filename, dynamic_columns)

            # excel_processingの関数を呼び出して、Excelファイルの更新を行う
            process_excel_data_4(df_uploaded, excel_filename, dynamic_columns)

            # excel_processingの関数を呼び出して、Excelファイルの更新を行う
            process_excel_data_5(df_uploaded, excel_filename, dynamic_columns)

            # excel_processingの関数を呼び出して、Excelファイルの更新を行う
            process_excel_data_6(df_uploaded, excel_filename, dynamic_columns)

            # excel_processingの関数を呼び出して、Excelファイルの更新を行う
            process_excel_data_7(df_uploaded, excel_filename, dynamic_columns)

            # excel_processingの関数を呼び出して、Excelファイルの更新を行う
            process_excel_data_8(df_uploaded, excel_filename, dynamic_columns, uploaded_filename)




            if os.path.isfile(excel_filename):
                with open(excel_filename, "rb") as file:
                    btn = st.download_button(
                            label="Excelをダウンロード",
                            data=file,
                            file_name=excel_filename,
                            mime="application/vnd.ms-excel"
                        )
                st.balloons()        
            else:
                st.write("ファイルが見つかりません。")
