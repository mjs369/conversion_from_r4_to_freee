import pandas as pd
from io import BytesIO

def process(uploaded_file1, uploaded_file4, new_data):
    # file1（GSHISAN_DATA.CSV）を読み込む
    file1_io = BytesIO(uploaded_file1.getvalue())
    df_file1 = pd.read_csv(file1_io,)

    # file4（GBASHO.CSV）を読み込む（ヘッダーなし）
    file4_io = BytesIO(uploaded_file4.getvalue())
    df_file4 = pd.read_csv(file4_io, header=None)

    # 市町村データを読み込む
    df_city = pd.read_csv('市町村.csv', encoding='cp932')

    # file1の「設置場所コード」とfile4の0列目を照合
    for index, row in df_file1.iterrows():
        location_code = row['設置場所コード']
        matched_row = df_file4[df_file4[0] == location_code]

        if not matched_row.empty:
            city_name = matched_row.iloc[0, 3]  # file4の3列目のデータ
            new_data.at[index, '申告先市区町村'] = city_name

            # 市区町村データを用いて都道府県名を最適に取得
            matched_cities = df_city[df_city['都道府県&市区町村'].str.contains(city_name)]
            if not matched_cities.empty:
                # 文字数が最も多い市区町村を選択
                best_match = matched_cities.loc[matched_cities['都道府県&市区町村'].str.len().idxmax()]
                # 同じ文字数の場合、東京都を優先
                if matched_cities['都道府県&市区町村'].str.len().max() == len(best_match['都道府県&市区町村']):
                    tokyo_match = matched_cities[matched_cities['都道府県'] == '東京都']
                    if not tokyo_match.empty:
                        best_match = tokyo_match.iloc[0]

                prefecture_name = best_match['都道府県']
                new_data.at[index, '申告先都道府県'] = prefecture_name
            else:
                new_data.at[index, '申告先都道府県'] = "不明"

    return new_data
