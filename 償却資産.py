import streamlit as st
import pandas as pd
from io import BytesIO

def app3():
    st.title("固定資産台帳インポート")

    st.markdown("""
    	#### 減価償却R4からの出力手順
        1.画面左上の「ファイル」→「外部ファイル作成/取込」→資産データにチェックを入れる。
        「参照」から保存場所をデスクトップに変更→「OK」ボタンで出力
        2.画面上部「導入」→「設置場所情報一覧表」→ファイル出力
        3.画面上部「導入」→「部門一覧表」→ファイル出力
        画面上部「導入」→「勘定一覧表」→ファイル出力
        4.それぞれを下部の指定場所にアップロード
    	""")

    # CSVファイルのアップロード
    uploaded_file1 = st.file_uploader("GSHISAN_DATAのcsvをアップロード", type=['csv'], key="file1")
    uploaded_file2 = st.file_uploader("GKANJOのcsvをアップロード", type=['csv'], key="file2")
    uploaded_file3 = st.file_uploader("GBUMONのcsvをアップロード", type=['csv'], key="file3")
    uploaded_file4 = st.file_uploader("GBASHOのcsvをアップロード", type=['csv'], key="file4")

    # OKボタンの配置
    if st.button('Start Conversion'):
        # 新しいCSVファイル（new）のヘッダーを作成
        new_data = pd.DataFrame(columns=[
            '資産の名前', '取得価額', '数量', '数量単位', '勘定科目', '取得日', '耐用年数', 
            '償却方法', '事業供用開始日', '期首残高', '改定取得価額', '特別償却費', 
            '管理番号等', '摘要', '部門', '申告先都道府県', '申告先市区町村', 
            '減少事由', '減少年月日', '製造業利用比率', '減価償却に使う勘定科目'
        ])

        # file1〜file4の処理モジュールをインポート
        import file1
        import file2
        import file3
        import file4

        # file1がアップロードされていれば、file1.pyを実行
        if uploaded_file1 is not None:
            file1.process(uploaded_file1, new_data)

        # file2がアップロードされていれば、file2.pyを実行
        if uploaded_file2 is not None:
            file2.process(uploaded_file1, uploaded_file2, new_data)

        # file3がアップロードされていれば、file3.pyを実行
        if uploaded_file3 is not None:
            file3.process(uploaded_file1, uploaded_file3, new_data)

        # file4がアップロードされていれば、file4.pyを実行
        if uploaded_file4 is not None:
            file4.process(uploaded_file1, uploaded_file4, new_data)

        # "減価償却に使う勘定科目"列に"減価償却費"という文字列を設定
        new_data["減価償却に使う勘定科目"] = "減価償却費"

        # '資産の名前'が入力されており、'期首残高'が空の場合、'0'を設定
        new_data.loc[(new_data['資産の名前'].notna()) & (new_data['期首残高'].isna()), '期首残高'] = 0

        # '期首残高'を整数に変換
        new_data['期首残高'] = new_data['期首残高'].fillna(0).astype(float).astype(int)

        # NaNでない値を整数に変換
        new_data['改定取得価額'] = new_data['改定取得価額'].apply(lambda x: int(x) if pd.notna(x) else x)

        # NaN値を0に置き換える
        new_data['改定取得価額'] = new_data['改定取得価額'].fillna(0)

        # 0を再度NaNに設定
        new_data.loc[new_data['改定取得価額'] == 0, '改定取得価額'] = pd.NA


        # 出力するCSVファイルを"cp932"エンコードで保存
        new_data.to_csv('new_data.csv', index=False, encoding='cp932')

        # ダウンロードリンクを生成
        with open('new_data.csv', 'rb') as f:
            st.download_button('Download CSV', f, file_name='償却資産インポート.csv')
