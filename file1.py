import pandas as pd
import streamlit as st

def process(uploaded_file, new_data):
    # file1を"cp932"で読み込む
    df_file1 = pd.read_csv(uploaded_file)

    # file1の特定の列をnew_dataにコピー
    new_data['資産の名前'] = df_file1['資産名']
    new_data['取得日'] = df_file1['取得年月日']
    new_data['事業供用開始日'] = df_file1['事業供用年月日']
    new_data['取得価額'] = df_file1['取得価額']
    new_data['期首残高'] = df_file1['期首帳簿価額']
    new_data['数量'] = df_file1['数量']
    new_data['耐用年数'] = df_file1['耐用年数']
    new_data['改定取得価額'] = df_file1['改定取得価額']

    # 償却方法の変換
    def convert_depreciation_method(method):
        mapping = {
            0: "償却なし", 2: "定額法", 3: "定率法", 4: "定率法",
            5: "任意償却", 6: "任意償却", 7: "均等償却", 9: "一括償却",
            10: "定額法", 902: "旧定額法", 904: "旧定率法"
        }
        return mapping.get(method, "")

    new_data['償却方法'] = df_file1['償却方法'].map(convert_depreciation_method)

    # 特定の償却方法の場合、ユーザーにプルダウンで選択させる
    special_methods = [8, 12, 13, 14, 910, 911, 912, 914]
    dropdown_options = [
        '少額償却', '一括償却', '定額法', '定率法', '償却なし', 
        '任意償却', '即時償却', '均等償却', '旧定額法', '旧定率法'
    ]
    for index, row in df_file1.iterrows():
        if row['償却方法'] in special_methods:
            choice = st.selectbox(
                f"行 {index}: 償却方法を選択してください",
                options=dropdown_options,
                key=f"dep_method_{index}"
            )
            new_data.at[index, '償却方法'] = choice

    # file1の特定の列をnew_dataにコピー
    new_data['資産の名前'] = df_file1['資産名']
    new_data['取得日'] = df_file1['取得年月日']
    new_data['事業供用開始日'] = df_file1['事業供用年月日']
    new_data['取得価額'] = df_file1['取得価額']
    new_data['期首残高'] = df_file1['期首帳簿価額']
    new_data['数量'] = df_file1['数量']
    new_data['耐用年数'] = df_file1['耐用年数']
    new_data['減少年月日'] = df_file1['除売却年月日']

    # 新しい条件の追加
    # 減少事由に基づく値の設定
    def get_decrease_reason(value):
        if value == 1:
            return '売却'
        elif value == 2:
            return '除却'
        elif value == 3:
            return 'その他減少'
        return ''  # 既定値

    new_data['減少事由'] = df_file1['減少事由'].apply(get_decrease_reason)

    # 製造原価割合が1.0以上の場合、その値を製造業利用比率に転記
    new_data['製造業利用比率'] = df_file1.apply(lambda row: row['製造原価割合'] if row['製造原価割合'] >= 1.0 else '', axis=1)

