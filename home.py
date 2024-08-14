import streamlit as st
import 残高試算表
import 変換
import 償却資産
import DepartmentalReportCreator
import suii
import tb_bs
import tb_pl

# 初期状態では何も表示しないようにセッション状態を設定
if 'current_app' not in st.session_state:
    st.session_state['current_app'] = None

st.title('EG_freee関連アプリ')


# ボタンが押されたときに実行する関数
def show_app1():
    st.session_state['current_app'] = 'app1'

def show_app2():
    st.session_state['current_app'] = 'app2'

def show_app3():
    st.session_state['current_app'] = 'app3'

def show_app4():
    st.session_state['current_app'] = 'app4'

def show_app5():
    st.session_state['current_app'] = 'app5'

def show_app6():
    st.session_state['current_app'] = 'app6'

def show_app7():
    st.session_state['current_app'] = 'app7'

# サイドバーでアプリ選択用のボタンを表示
with st.sidebar:
    if st.button('残高試算表(txtファイル)', key='app1'):
        show_app1()

    if st.button('仕訳(csvファイル)', key='app2'):
        show_app2()

    if st.button('償却資産(csvファイル)', key='app3'):
        show_app3()

    if st.button('部門分割(csvファイル)', key='app4'):
        show_app4()

    if st.button('推移表変換(csvファイル)', key='app5'):
        show_app5()

    if st.button('試算表BS変換(csvファイル)', key='app6'):
        show_app6()

    if st.button('試算表変換PL(csvファイル)', key='app7'):
        show_app7()

# 選択されたアプリを表示
if st.session_state['current_app'] == 'app1':
    残高試算表.app1()
elif st.session_state['current_app'] == 'app2':
    変換.app2()
elif st.session_state['current_app'] == 'app3':
    償却資産.app3()
elif st.session_state['current_app'] == 'app4':
    DepartmentalReportCreator.app4()
elif st.session_state['current_app'] == 'app5':
    suii.app5()
elif st.session_state['current_app'] == 'app6':
    tb_bs.app6()
elif st.session_state['current_app'] == 'app7':
    tb_pl.app7()