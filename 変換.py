import streamlit as st
import pandas as pd
from io import BytesIO

def app2():
    # Streamlitアプリのタイトル
    st.title('仕訳インポート')

    # 説明文の追加（Markdown形式）
    st.markdown("""
        ### R4での出力手順
        1. 連動「仕訳データ作成」を選択
        2. 「ヘッダー有」「通常」であることを確認
        3. 「実行」ではなく「Excel(CF10)」から出力
        4. 出力された「Book1」で「Ctrl+S」を押し名前を付けてデスクトップに保存。
        5. 保存する際の形式は".csv"を選択 ※ファイルの種類から必ず変更
        """)

    # ファイルアップロードのセクション
    uploaded_file = st.file_uploader("ファイルをアップロードしてください", type=['csv'])

    # アップロードされたファイルがあれば処理を実行
    if uploaded_file is not None:
        # データフレームを読み込み
        df = pd.read_csv(uploaded_file, encoding='cp932')

        # 指定されたヘッダー
        headers = ["[表題行]", "日付", "伝票番号", "決算整理仕訳", "借方勘定科目", "借方科目コード", "借方補助科目", "借方取引先", "借方取引先コード",
                   "借方部門", "借方品目", "借方メモタグ", "借方セグメント1", "借方セグメント2", "借方セグメント3", "借方金額", "借方税区分",
                   "借方税額", "貸方勘定科目", "貸方科目コード", "貸方補助科目", "貸方取引先", "貸方取引先コード", "貸方部門", "貸方品目",
                   "貸方メモタグ", "貸方セグメント1", "貸方セグメント2", "貸方セグメント3", "貸方金額", "貸方税区分", "貸方税額", "摘要"]

        # ヘッダーを持つ空のDataFrameを作成
        df_with_headers = pd.DataFrame(columns=headers)

        df_with_headers["日付"] = df["伝票日付"]
        df_with_headers["借方勘定科目"] = df["借方科目名"]
        df_with_headers["貸方勘定科目"] = df["貸方科目名"]
        df_with_headers["借方補助科目"] = df["借方補助科目名"]
        df_with_headers["貸方補助科目"] = df["貸方補助科目名"]
        df_with_headers["借方金額"] = df["借方金額"]
        df_with_headers["貸方金額"] = df["貸方金額"]
        df_with_headers["摘要"] = df["摘要"]
        df_with_headers["貸方部門"] = df["貸方部門名"]
        df_with_headers["借方部門"] = df["借方部門名"]


        def encodeable_in_shift_jis(value):
            # 文字列でない場合はそのまま返す
            if not isinstance(value, str):
                return value

            # エンコードできる文字だけを取り出して新しい文字列を生成
            return ''.join([char for char in value if can_encode_shift_jis(char)])

        # エンコードできるかを判断するヘルパー関数
        def can_encode_shift_jis(char):
            try:
                char.encode('shift_jis')
                return True
            except UnicodeEncodeError:
                return False

        # 取り込んだCSVのヘッダー"借方消費税コード"の値に基づき、新しい値をセット
        df_with_headers["借方税区分"] = df["借方消費税コード"].map({
            2: "課税売上(込)",
            1: "課税売上(抜)",
            4: "課税売上(22条)",
            5: "輸出売上",
            6: "輸出売上(非課税)",
            7: "売上海外支店",
            8: "課税売上みなし(抜)",
            9: "課税売上みなし(込)",
            10: "有価証券売却",
            11: "課税売上返還(抜)",
            12: "課税売上返還(込)",
            14: "課税売上返還(22条)",
            15: "輸出売上返還",
            16: "輸出売上返還(非課税)",
            17: "売上海外支店返還",
            20: "非課税売上",
            21: "非課税売上返還",
            23: "課税貸倒損失(抜)",
            24: "課税貸倒損失(込)",
            25: "課税貸倒損失(22条)",
            26: "課税貸倒回収(抜)",
            27: "課税貸倒回収(込)",
            28: "課税貸倒回収(22条)",
            30: "非課税仕入",
            31: "仕入対課税(抜)",
            32: "仕入対課税(込)",
            33: "仕入対課税(22条)",
            34: "輸入対課税",
            35: "仕入対課税返還(抜)",
            36: "仕入対課税返還(込)",
            37: "仕入対課税返還(22条)",
            38: "輸入対課税返還",
            41: "仕入共通(抜)",
            42: "仕入共通(込)",
            43: "仕入共通(22条)",
            44: "輸入共通",
            45: "仕入共通返還(抜)",
            46: "仕入共通返還(込)",
            47: "仕入共通返還(22条)",
            48: "輸入共通返還",
            51: "仕入対非課税(抜)",
            52: "仕入対非課税(込)",
            53: "仕入対非課税22条)",
            54: "輸入対非課税",
            55: "仕入対非課税返還(抜)",
            56: "仕入対非課税返還(込)",
            57: "仕入対非課税返還(22条)",
            58: "輸入対非課税返還",
            61: "輸入地方税(対課税)",
            62: "輸入地方税(共通)",
            63: "輸入地方税(対非課税)",
            64: "輸入国税込(対課税)",
            65: "輸入国税込(共通)",
            66: "輸入国税込(対非課税)",
            67: "輸入地税込(対課税)",
            68: "輸入地税込(共通)",
            69: "輸入地税込(対非課税)",
            71: "輸入国地税(対課税)",
            72: "輸入国地税(共通)",
            73: "輸入国地税(対非課税)",
            74: "輸入国地込(対課税)",
            75: "輸入国地込(共通)",
            76: "輸入国地込(対非課税)",
            91: "特定課仕(対課税)",
            92: "特定課仕(共通)",
            93: "特定課仕(対非課税)",
            94: "特定課仕返(共通)",
            95: "特定課仕(対非課税)",
            96: "特定課仕返(対非課税)",
            80: "不課税売上",
            81: "不課税仕入",
            99: "不明",
        }).fillna(df_with_headers["借方税区分"])

        # 取り込んだCSVのヘッダー"借方消費税コード"の値に基づき、新しい値をセット
        df_with_headers["貸方税区分"] = df["貸方消費税コード"].map({
            2: "課税売上(込)",
            1: "課税売上(抜)",
            4: "課税売上(22条)",
            5: "輸出売上",
            6: "輸出売上(非課税)",
            7: "売上海外支店",
            8: "課税売上みなし(抜)",
            9: "課税売上みなし(込)",
            10: "有価証券売却",
            11: "課税売上返還(抜)",
            12: "課税売上返還(込)",
            14: "課税売上返還(22条)",
            15: "輸出売上返還",
            16: "輸出売上返還(非課税)",
            17: "売上海外支店返還",
            20: "非課税売上",
            21: "非課税売上返還",
            23: "課税貸倒損失(抜)",
            24: "課税貸倒損失(込)",
            25: "課税貸倒損失(22条)",
            26: "課税貸倒回収(抜)",
            27: "課税貸倒回収(込)",
            28: "課税貸倒回収(22条)",
            30: "非課税仕入",
            31: "仕入対課税(抜)",
            32: "仕入対課税(込)",
            33: "仕入対課税(22条)",
            34: "輸入対課税",
            35: "仕入対課税返還(抜)",
            36: "仕入対課税返還(込)",
            37: "仕入対課税返還(22条)",
            38: "輸入対課税返還",
            41: "仕入共通(抜)",
            42: "仕入共通(込)",
            43: "仕入共通(22条)",
            44: "輸入共通",
            45: "仕入共通返還(抜)",
            46: "仕入共通返還(込)",
            47: "仕入共通返還(22条)",
            48: "輸入共通返還",
            51: "仕入対非課税(抜)",
            52: "仕入対非課税(込)",
            53: "仕入対非課税22条)",
            54: "輸入対非課税",
            55: "仕入対非課税返還(抜)",
            56: "仕入対非課税返還(込)",
            57: "仕入対非課税返還(22条)",
            58: "輸入対非課税返還",
            61: "輸入地方税(対課税)",
            62: "輸入地方税(共通)",
            63: "輸入地方税(対非課税)",
            64: "輸入国税込(対課税)",
            65: "輸入国税込(共通)",
            66: "輸入国税込(対非課税)",
            67: "輸入地税込(対課税)",
            68: "輸入地税込(共通)",
            69: "輸入地税込(対非課税)",
            71: "輸入国地税(対課税)",
            72: "輸入国地税(共通)",
            73: "輸入国地税(対非課税)",
            74: "輸入国地込(対課税)",
            75: "輸入国地込(共通)",
            76: "輸入国地込(対非課税)",
            91: "特定課仕(対課税)",
            92: "特定課仕(共通)",
            93: "特定課仕(対非課税)",
            94: "特定課仕返(共通)",
            95: "特定課仕(対非課税)",
            96: "特定課仕返(対非課税)",
            80: "不課税売上",
            81: "不課税仕入",
            99: "不明",
        }).fillna(df_with_headers["貸方税区分"])

        for index, row in df.iterrows():
            tax_rate = row['借方消費税税率']
            current_tax_division = df_with_headers.at[index, '借方税区分']

            # 特殊な文字列ケース 'K8.0' をチェック
            if tax_rate == "K8.0":
                if pd.notna(current_tax_division):
                    df_with_headers.at[index, '借方税区分'] += " 8%(軽)"
            elif tax_rate in ["10", "3", "5", "8"]:  # 文字列としての通常の数値ケース
                if pd.notna(current_tax_division):
                    df_with_headers.at[index, '借方税区分'] += f" {tax_rate}%"

            # 浮動小数点数であり、NaNでない場合に整数に変換し、その後文字列に変換
            elif isinstance(tax_rate, float) and not pd.isna(tax_rate):
                tax_rate = str(int(tax_rate))
                if pd.notna(current_tax_division):
                    if tax_rate == "10":
                        df_with_headers.at[index, '借方税区分'] += " 10%"
                    elif tax_rate == "3":
                        df_with_headers.at[index, '借方税区分'] += " 3%"
                    elif tax_rate == "5":
                        df_with_headers.at[index, '借方税区分'] += " 5%"
                    elif tax_rate == "8":
                        df_with_headers.at[index, '借方税区分'] += " 8%"

        for index, row in df.iterrows():
            tax_rate = row['貸方消費税税率']
            current_tax_division = df_with_headers.at[index, '貸方税区分']

            # 特殊な文字列ケース 'K8.0' をチェック
            if tax_rate == "K8.0":
                if pd.notna(current_tax_division):
                    df_with_headers.at[index, '貸方税区分'] += " 8%(軽)"
            elif tax_rate in ["10", "3", "5", "8"]:  # 文字列としての通常の数値ケース
                if pd.notna(current_tax_division):
                    df_with_headers.at[index, '貸方税区分'] += f" {tax_rate}%"

            # 浮動小数点数であり、NaNでない場合に整数に変換し、その後文字列に変換
            elif isinstance(tax_rate, float) and not pd.isna(tax_rate):
                tax_rate = str(int(tax_rate))
                if pd.notna(current_tax_division):
                    if tax_rate == "10":
                        df_with_headers.at[index, '貸方税区分'] += " 10%"
                    elif tax_rate == "3":
                        df_with_headers.at[index, '貸方税区分'] += " 3%"
                    elif tax_rate == "5":
                        df_with_headers.at[index, '貸方税区分'] += " 5%"
                    elif tax_rate == "8":
                        df_with_headers.at[index, '貸方税区分'] += " 8%"


        for index, row in df.iterrows():
            tax_rate = row['貸方消費税業種']
            # 税率が文字列でない場合は、文字列に変換
            if not isinstance(tax_rate, str):
                tax_rate = str(tax_rate)
            if tax_rate == "1":
                df_with_headers.at[index, '貸方税区分'] += " 第1種事業"
            elif tax_rate == "2":
                df_with_headers.at[index, '貸方税区分'] += " 第2種事業"
            elif tax_rate == "3":
                df_with_headers.at[index, '貸方税区分'] += " 第3種事業"
            elif tax_rate == "4":
                df_with_headers.at[index, '貸方税区分'] += " 第4種事業"
            elif tax_rate == "5":
                df_with_headers.at[index, '貸方税区分'] += " 第5種事業"
            elif tax_rate == "6":
                df_with_headers.at[index, '貸方税区分'] += " 第6種事業"

        for index, row in df.iterrows():
            tax_rate = row['借方消費税業種']
            # 税率が文字列でない場合は、文字列に変換
            if not isinstance(tax_rate, str):
                tax_rate = str(tax_rate)
            if tax_rate == "1":
                df_with_headers.at[index, '借方税区分'] += " 第1種事業"
            elif tax_rate == "2":
                df_with_headers.at[index, '借方税区分'] += " 第2種事業"
            elif tax_rate == "3":
                df_with_headers.at[index, '借方税区分'] += " 第3種事業"
            elif tax_rate == "4":
                df_with_headers.at[index, '借方税区分'] += " 第4種事業"
            elif tax_rate == "5":
                df_with_headers.at[index, '借方税区分'] += " 第5種事業"
            elif tax_rate == "6":
                df_with_headers.at[index, '借方税区分'] += " 第6種事業"

        # 税額を計算して更新するための関数
        def calculate_tax_amount(row):
            if row['借方消費税税率'] == "10":
                # 借方金額を110で割って10を掛け、小数点以下切り捨て
                return int((row['借方金額'] / 110) * 10)
            elif row['借方消費税税率'] == "8":
              return int((row['借方金額'] / 108) * 8)
            elif row['借方消費税税率'] == "5":
              return int((row['借方金額'] / 105) * 5)
            elif row['借方消費税税率'] == "3":
              return int((row['借方金額'] / 103) * 3)
            elif row['借方消費税税率'] == "K8.0":
              return int((row['借方金額'] / 108) * 8)

        # dfを使用して税額を計算し、df_with_headersに'借方税額'として追加
        df_with_headers['借方税額'] = df.apply(calculate_tax_amount, axis=1)

        def calculate_tax_amount(row):
            if row['貸方消費税税率'] == "10":
                # 貸方金額を110で割って10を掛け、小数点以下切り捨て
                return int((row['貸方金額'] / 110) * 10)
            elif row['貸方消費税税率'] == "8":
              return int((row['貸方金額'] / 108) * 8)
            elif row['貸方消費税税率'] == "5":
              return int((row['貸方金額'] / 105) * 5)
            elif row['貸方消費税税率'] == "3":
              return int((row['貸方金額'] / 103) * 3)
            elif row['貸方消費税税率'] == "K8.0":
              return int((row['貸方金額'] / 108) * 8)

        # "借方インボイス情報"と"貸方インボイス情報"に基づき、df_with_headersの税区分を更新する
        for index, row in df.iterrows():
            try:
                # 借方インボイス情報を整数に変換
                debit_invoice_info = int(row["借方インボイス情報"])
                # 値が8の場合、借方税区分に"(控80)"を追加
                if debit_invoice_info == 8:
                    df_with_headers.at[index, "借方税区分"] += " (控80)"
            except (ValueError, TypeError):
                # 変換できない場合は無視
                pass

            try:
                # 貸方インボイス情報を整数に変換
                credit_invoice_info = int(row["貸方インボイス情報"])
                # 値が8の場合、貸方税区分に"(控80)"を追加
                if credit_invoice_info == 8:
                    df_with_headers.at[index, "貸方税区分"] += " (控80)"
            except (ValueError, TypeError):
                # 変換できない場合は無視
                pass

        # CSVファイルとしてユーザーにダウンロードを提供する関数
        def convert_df_to_csv(df):
            # CSVに変換
            return df.to_csv(index=False, encoding='cp932').encode('cp932')

        # データフレームをCSVファイルとしてメモリ上に書き出す
        csv = convert_df_to_csv(df_with_headers)

        # ダウンロードボタンの作成
        st.download_button(
            label="変換されたCSVをダウンロード",
            data=BytesIO(csv),
            file_name='freee仕訳取込仕訳.csv',
            mime='text/csv',
        )
        st.snow()
