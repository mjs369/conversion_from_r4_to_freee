import streamlit as st
import pandas as pd
from io import StringIO
from io import BytesIO

def app1():
    # Streamlitアプリのタイトル
    st.title('開始残高インポート')

    # 説明文の追加（Markdown形式）
    st.markdown("""
        ### R4での出力手順
        1. 帳表「合計残高試算表」を選択
        2. 条件設定「出力形式：通常/出力区分：累計/金額：円」であることを確認
        3. 出力設定「表紙：しない/貸借対照表：する/損益計算書：しない/原価報告書：しない」を選択
        4. 「実行」ではなくヘッダー右側の「テキスト(CF12)」から.txtファイルで出力
        """)


    # ファイルアップロードのセクション
    uploaded_file = st.file_uploader("テキストファイルをアップロードしてください", type=['txt'])

    # アップロードされたファイルがあれば処理を実行
    if uploaded_file is not None:
        try:
            # cp932エンコーディングで読み込み
            string_data = uploaded_file.read().decode('cp932')
            stringio = StringIO(string_data)
            df = pd.read_csv(stringio, header=3)  # delimiterはファイルに応じて変更してください

            # 勘定科目コードが文字列であれば数値に変換
            df['勘定科目コード'] = pd.to_numeric(df['勘定科目コード'], errors='coerce')

            # NaNを含む行を除外
            df = df.dropna(subset=['勘定科目コード'])
            
            # 勘定科目コードを整数に変換
            df['勘定科目コード'] = df['勘定科目コード'].astype(int)

            # 除外したい勘定科目コードを含む行を除外
            exclude_codes = [9025, 9045, 9055, 9075, 9145, 9150, 9200, 9210,9230, 9240, 9300, 9350, 9370, 9400, 9440, 9445, 9450, 9500,9380]
            df = df[~df["勘定科目コード"].isin(exclude_codes)]
            
            # 新しいDataFrameを作成
            new_headers = ["勘定科目", "借方残高", "貸方残高"]
            new_df = pd.DataFrame(columns=new_headers)

            # 条件に基づいて前月残高を貸方残高または借方残高に分割
            new_df["勘定科目"] = df["勘定科目名"]
            new_df["借方残高"] = df.apply(lambda x: x["前月残高"] if 100 <= x["勘定科目コード"] <= 199 else 0, axis=1)
            new_df["貸方残高"] = df.apply(
                lambda x: x["前月残高"] 
                          if (200 <= x["勘定科目コード"] <= 299) or 
                             (300 <= x["勘定科目コード"] <= 399) or 
                             (x["勘定科目コード"] == 9430) 
                          else 0, axis=1)

            # 条件に基づく分割を行った後で、条件に一致しない行のインデックスを取得
            unmatched_idx = df.index[
                ~df["勘定科目コード"].between(100, 199) &
                ~df["勘定科目コード"].between(200, 299) &
                ~df["勘定科目コード"].between(300, 399) &
                ~(df["勘定科目コード"] == 9430)
            ]

            # 条件に一致しない各行についてユーザーが選択できるようにする
            for idx in unmatched_idx:
                # 2つのカラムを作成
                col1, col2 = st.columns(2)
                
                # 最初のカラムには勘定科目名とコードを表示
                with col1:
                    st.write(f"勘定科目名: {df.loc[idx, '勘定科目名']}, 勘定科目コード: {df.loc[idx, '勘定科目コード']}")
                
                # 2つ目のカラムにラジオボタンを配置
                with col2:
                    choice = st.radio(
                        "どちらに割り振りますか？",
                        ('借方残高', '貸方残高', '転記しない'),
                        key=f"choice_{idx}"
                    )

                # ユーザーの選択に基づいてDataFrameを更新
                if choice == '借方残高':
                    new_df.loc[idx, "借方残高"] = df.loc[idx, "前月残高"]
                    new_df.loc[idx, "貸方残高"] = 0
                elif choice == '貸方残高':
                    new_df.loc[idx, "貸方残高"] = df.loc[idx, "前月残高"]
                    new_df.loc[idx, "借方残高"] = 0
                elif choice == '転記しない':
                    # '転記しない'が選択された場合、その行をnew_dfから削除する
                    new_df.drop(idx, inplace=True)

            # 削除された行のインデックスをリセットする
            new_df.reset_index(drop=True, inplace=True)


            # 新しいCSVデータとしてユーザーにダウンロードさせる関数
            def convert_df_to_csv(df):
                # データフレームをCSVに変換する
                csv_bytes = df.to_csv(index=False, encoding='cp932').encode('cp932')
                return BytesIO(csv_bytes)

        except UnicodeDecodeError:
            st.error('ファイルのエンコーディングがcp932ではないかもしれません。別のエンコーディングを試してください。')
            st.stop()
        except Exception as e:
            st.error(f'ファイルの読み込み中にエラーが発生しました: {e}')
            st.stop()

        # ダウンロードボタンを表示する
        st.download_button(
            label="新しいCSVファイルをダウンロード",
            data=convert_df_to_csv(new_df),
            file_name='開始残高インポート.csv',
            mime='text/csv',
        )
