import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

def app5():

    st.title('推移表変換')

    def preprocess_dataframe(df):
        df = df.rename(columns={
            'Unnamed: 0': '階層1',
            'Unnamed: 1': '階層2',
            'Unnamed: 2': '階層3',
            'Unnamed: 3': '階層4',
            'Unnamed: 4': '階層5',
            'Unnamed: 5': '階層6'
        })

        for i, row in df.iterrows():
            values = row[['階層2', '階層3', '階層4', '階層5']].dropna().tolist()
            if values:
                df.at[i, '階層1'] = ' '.join(values)

        df[['階層2', '階層3', '階層4', '階層5']] = None
        df = df.drop(columns=['階層2', '階層3', '階層4', '階層5', '階層6'])
        df = df.dropna(subset=['期間累計'])

        new_columns = ['前期累計', '当期累計', '前年差額']
        for col in new_columns:
            df[col] = np.nan

        cols = list(df.columns)
        insert_position = cols.index('階層1') + 1
        cols = cols[:insert_position] + new_columns + cols[insert_position:-len(new_columns)]
        df = df[cols]

        return df

    def keisan(df, df2, x):
        df['当期累計'] = df.iloc[:, 4:x].sum(axis=1)
        df2['指定累計'] = df2.iloc[:, 4:x].sum(axis=1)

        df['前期累計'] = 0
        df['前年差額'] = 0

        for idx, row in df.iterrows():
            hierarchy1_value = row['階層1']
            if hierarchy1_value in df2['階層1'].values:
                idx2 = df2[df2['階層1'] == hierarchy1_value].index[0]
                df.at[idx, '前期累計'] = df2.at[idx2, '指定累計']
                df.at[idx, '前年差額'] = df.at[idx, '当期累計'] - df2.at[idx2, '指定累計']

        for idx2, row in df2.iterrows():
            hierarchy1_value = row['階層1']
            if hierarchy1_value not in df['階層1'].values:
                new_row = pd.Series(index=df.columns, dtype='object')
                new_row['階層1'] = hierarchy1_value
                new_row['当期累計'] = 0.0
                new_row['前期累計'] = row['指定累計']
                new_row['前年差額'] = -row['指定累計']

                previous_value = None
                next_value = None
                if idx2 > 0:
                    previous_value = df2.at[idx2 - 1, '階層1']
                if idx2 + 1 < len(df2):
                    next_value = df2.at[idx2 + 1, '階層1']

                if previous_value in df['階層1'].values and next_value in df['階層1'].values:
                    previous_index = df.index[df['階層1'] == previous_value].tolist()[0]
                    next_index = df.index[df['階層1'] == next_value].tolist()[0]
                    insert_position = previous_index + 1
                elif previous_value in df['階層1'].values:
                    previous_index = df.index[df['階層1'] == previous_value].tolist()[0]
                    insert_position = previous_index + 1
                elif next_value in df['階層1'].values:
                    next_index = df.index[df['階層1'] == next_value].tolist()[0]
                    insert_position = next_index
                else:
                    insert_position = len(df)

                df = pd.concat([df.iloc[:insert_position], new_row.to_frame().T, df.iloc[insert_position:]]).reset_index(drop=True)

        df = df.rename(columns={'階層1': '科目名'})
        month_columns = [col for col in df.columns if col not in ['科目名', '前期累計', '当期累計', '前年差額']]
        new_columns = {'前期累計': '前期累計', '当期累計': '当期累計', '前年差額': '前年差額'}
        for col in month_columns:
            if '-' in col:
                month = int(col.split('-')[1])
                new_columns[col] = f"{month}月"
            else:
                new_columns[col] = col

        df = df.rename(columns=new_columns)
        df = df.fillna(0)

        return df

    uploaded_file = st.file_uploader("当期分をアップロードしてください", type="csv", key="file_uploader1")
    uploaded_file2 = st.file_uploader("前期分をアップロードしてください", type="csv", key="file_uploader2")

    if uploaded_file and uploaded_file2:
        df = pd.read_csv(uploaded_file, encoding='cp932', header=1)
        df = preprocess_dataframe(df)
        
        df2 = pd.read_csv(uploaded_file2, encoding='cp932', header=1)
        df2 = preprocess_dataframe(df2)

        df.to_csv('hen.csv', encoding='cp932', index=None)
        df2.to_csv('hen2.csv', encoding='cp932', index=None)
        df = pd.read_csv('hen.csv', encoding='cp932')
        df2 = pd.read_csv('hen2.csv', encoding='cp932')

        month_options = {
            '12ヵ月': 16,
            '11ヵ月': 15,
            '10ヵ月': 14,
            '9ヵ月': 13,
            '8ヵ月': 12,
            '7ヵ月': 11,
            '6ヵ月': 10,
            '5ヵ月': 9,
            '4ヵ月': 8,
            '3ヵ月': 7,
            '2ヵ月': 6,
            '1ヵ月': 5
        }

        selected_month = st.selectbox("累計月を選択してください", list(month_options.keys()))
        x = month_options[selected_month]

        df_result = keisan(df, df2, x)
        
        def convert_df_to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            processed_data = output.getvalue()
            return processed_data

        excel = convert_df_to_excel(df_result)

        st.download_button(
            label="変換されたExcelをダウンロード",
            data=BytesIO(excel),
            file_name='freee推移表.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
        st.snow()

app5()
