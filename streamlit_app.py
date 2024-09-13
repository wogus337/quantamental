

import streamlit as st
from st_aggrid import AgGrid, GridUpdateMode, JsCode
from st_aggrid.grid_options_builder import GridOptionsBuilder
import pandas as pd
from functools import reduce
from pandas.tseries.offsets import BDay, DateOffset
import numpy as np
import openpyxl
import plotly.graph_objs as go
from plotly.subplots import make_subplots
import datetime
from datetime import datetime, timedelta
import streamlit_authenticator as stauth
import yaml
import statsmodels.api as sm
from statsmodels.stats.stattools import durbin_watson
from sklearn.linear_model import LinearRegression
import requests
import zipfile
import xml.etree.ElementTree as ET
import io
from io import BytesIO
from scipy.odr import ODR, Model, RealData
from scipy.stats import linregress
import itertools
import os
from PIL import Image


series_path = "data/streamlit_24.xlsx"
cylfile_path = "data/streamlit_24_cycle.xlsx"
simfile_path = "data/streamlit_24_sim.xlsx"
fx_path = "data/streamlit_24_fx.xlsx"
model_path = "data/streamlit_24_signal.xlsx"
macro_path = "data/streamlit_24_macro.xlsx"
usig_path = "data/streamlit_24_usigpick.xlsx"
market_path = "data/streamlit_24_marketVV.xlsx"
allo_path = "data/streamlit_24_allocation.xlsx"
fds_path = "data/streamlit_24_fds.xlsx"
pairres_path = "data/relativ_analysis_out_240830.csv"
slidepath = "images/QIS_Sep24"
image_path = "images/miraeasset.png"
igimage_path = "images/usig.png"
#
# series_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24.xlsx"
# cylfile_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24_cycle.xlsx"
# simfile_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24_sim.xlsx"
# fx_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24_fx.xlsx"
# model_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24_signal.xlsx"
# macro_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24_macro.xlsx"
# usig_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24_usigpick.xlsx"
# market_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24_marketVV.xlsx"
# allo_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24_allocation.xlsx"
# fds_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24_fds.xlsx"
# pairres_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\relativ_analysis_out_240830.csv"
# slidepath = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\QIS_Sep24"
# image_path = r"D:\Anaconda_envs\streamlit\pycharmprj\miraeasset.png"
# igimage_path = r"D:\Anaconda_envs\streamlit\pycharmprj\usig.png"


def get_closest_business_day(date, df):
    if date in df['DATE'].values:
        return date
    else:
        # Find the next closest business day
        previous_business_day = date - BDay(0)
        return previous_business_day


def calculate_change_with_offset(df, chgopt, months_offset=0, years_offset=0, days_offset=0):
    changes = {}
    for col in df.columns:
        if col != 'DATE':
            current_value = df[col].iloc[-1]
            # Calculate target date
            if months_offset:
                target_date = df['DATE'].iloc[-1] - DateOffset(months=months_offset)
            elif years_offset:
                target_date = df['DATE'].iloc[-1] - DateOffset(years=years_offset)
            elif days_offset:
                target_date = df['DATE'].iloc[-1] - BDay(days_offset)

            # Find the closest business day to the target date
            closest_date = get_closest_business_day(target_date, df)
            if closest_date in df['DATE'].values:
                previous_value = df[df['DATE'] == closest_date][col].values[0]
                if chgopt == 1:
                    changes[col] = ((current_value - previous_value) / previous_value) * 100
                elif chgopt == 2:
                    changes[col] = (current_value - previous_value)
            else:
                changes[col] = np.nan  # 데이터가 없을 경우 NaN 처리
    return pd.Series(changes)


def cal_table(df, chgopt, spechk):
    stable = pd.DataFrame(index=[
        'Latest', '1W', '1M', '3M', '6M', '1Y', '2Y', '3Y', '5Y', '10Y', 'MTD', 'YTD'
    ])
    for column in df.columns:
        if column != 'DATE':  # DATE 열을 제외한 나머지 열에 대해 계산
            latest_value = df[column].iloc[-1]  # 최신 값

            # 각 기간별 변동률 계산
            stable.at['Latest', column] = latest_value
            stable.at['1W', column] = float(
                calculate_change_with_offset(df[['DATE', column]], chgopt, days_offset=5).iloc[0])
            stable.at['1M', column] = float(
                calculate_change_with_offset(df[['DATE', column]], chgopt, months_offset=1).iloc[0])
            stable.at['3M', column] = float(
                calculate_change_with_offset(df[['DATE', column]], chgopt, months_offset=3).iloc[0])
            stable.at['6M', column] = float(
                calculate_change_with_offset(df[['DATE', column]], chgopt, months_offset=6).iloc[0])
            stable.at['1Y', column] = float(
                calculate_change_with_offset(df[['DATE', column]], chgopt, years_offset=1).iloc[0])
            stable.at['2Y', column] = float(
                calculate_change_with_offset(df[['DATE', column]], chgopt, years_offset=2).iloc[0])
            stable.at['3Y', column] = float(
                calculate_change_with_offset(df[['DATE', column]], chgopt, years_offset=3).iloc[0])
            stable.at['5Y', column] = float(
                calculate_change_with_offset(df[['DATE', column]], chgopt, years_offset=5).iloc[0])
            stable.at['10Y', column] = float(
                calculate_change_with_offset(df[['DATE', column]], chgopt, years_offset=10).iloc[0])
            stable.at['MTD', column] = float(calculate_change_with_offset(df[['DATE', column]], chgopt, days_offset=df['DATE'].iloc[-1].day - 1).iloc[0])
            stable.at['YTD', column] = float(calculate_change_with_offset(df[['DATE', column]], chgopt, days_offset=df['DATE'].iloc[-1].timetuple().tm_yday - 1).iloc[0])

    # 모든 숫자를 소수점 둘째 자리까지 반올림
    stable = stable.round(2)
    stable = stable.T
    stable = stable.reset_index()
    numcols = stable.columns[1:]
    stable[numcols] = stable[numcols].apply(pd.to_numeric)
    stable.rename(columns={"index": "X"}, inplace=True)

    if spechk == 1:
        cellstyle_jscode = JsCode("""
        function(params) {
            let style = {
                'textAlign': params.colDef.field === 'X' ? 'left' : 'right'
            };
            if (params.node.rowIndex % 2 === 0) {
                style['backgroundColor'] = 'gainsboro';
            } else {
                style['backgroundColor'] = 'white';
            }
            return style;
        }
        """)
    elif spechk == 2:
        cellstyle_jscode = JsCode("""
        function(params) {
            let style = {
                'textAlign': params.colDef.field === 'X' ? 'left' : 'right'
            };
            
            let rowIndex = params.node.rowIndex;
            if (rowIndex >= 0 && rowIndex <= 4) {
                style['backgroundColor'] = 'gainsboro';
            } else if (rowIndex >= 5 && rowIndex <= 10) {
                style['backgroundColor'] = 'white';
            } else if (rowIndex >= 11 && rowIndex <= 16) {
                style['backgroundColor'] = 'gainsboro';
            } else if (rowIndex >= 17 && rowIndex <= 22) {
                style['backgroundColor'] = 'white';
            } else if (rowIndex >= 23 && rowIndex <= 28) {
                style['backgroundColor'] = 'gainsboro';
            } else if (rowIndex >= 29 && rowIndex <= 34) {
                style['backgroundColor'] = 'white';
            } else if (rowIndex >= 35 && rowIndex <= 40) {
                style['backgroundColor'] = 'gainsboro';
            }
            return style;
        }
        """)

    # 숫자 포맷 설정
    formatter_jscode = JsCode("""
    function(params) {
        return Number(params.value).toFixed(2);
    }
    """)

    # columnDefs 설정
    grid_options = {
        'columnDefs': [
            {'field': 'X', 'width': 300, 'cellStyle': cellstyle_jscode},
            {'field': 'Latest', 'width': 200, 'cellStyle': cellstyle_jscode, 'valueFormatter': formatter_jscode},
            {'field': '1W', 'width': 150, 'cellStyle': cellstyle_jscode, 'valueFormatter': formatter_jscode},
            {'field': '1M', 'width': 150, 'cellStyle': cellstyle_jscode, 'valueFormatter': formatter_jscode},
            {'field': '3M', 'width': 150, 'cellStyle': cellstyle_jscode, 'valueFormatter': formatter_jscode},
            {'field': '6M', 'width': 150, 'cellStyle': cellstyle_jscode, 'valueFormatter': formatter_jscode},
            {'field': '1Y', 'width': 150, 'cellStyle': cellstyle_jscode, 'valueFormatter': formatter_jscode},
            {'field': '2Y', 'width': 150, 'cellStyle': cellstyle_jscode, 'valueFormatter': formatter_jscode},
            {'field': '3Y', 'width': 150, 'cellStyle': cellstyle_jscode, 'valueFormatter': formatter_jscode},
            {'field': '5Y', 'width': 150, 'cellStyle': cellstyle_jscode, 'valueFormatter': formatter_jscode},
            {'field': '10Y', 'width': 150, 'cellStyle': cellstyle_jscode, 'valueFormatter': formatter_jscode},
            {'field': 'MTD', 'width': 150, 'cellStyle': cellstyle_jscode, 'valueFormatter': formatter_jscode},
            {'field': 'YTD', 'width': 150, 'cellStyle': cellstyle_jscode, 'valueFormatter': formatter_jscode},
        ],
        'defaultColDef': {
            'resizable': True,
        },
        'domLayout': 'normal',
        'suppressHorizontalScroll': False
    }

    return stable, grid_options


st.set_page_config(layout="wide")

usernames = ['admin']
names = ['admin']
passwords = ['admin']
hashed_passwords = stauth.Hasher(passwords).generate()

# YAML
config = {
    'credentials': {
        'usernames': {
            usernames[i]: {
                'name': names[i],
                'password': hashed_passwords[i]
            } for i in range(len(usernames))
        }
    },
    'cookie': {
        'expiry_days': 1,
        'key': 'some_signature_key',
        'name': 'some_cookie_name'
    }
}

# authenticator
authenticator = stauth.Authenticate(
    config['credentials'],
    config['cookie']['name'],
    config['cookie']['key'],
    config['cookie']['expiry_days'],
)

col1, col2 = st.columns(2)
with col1:
    name, authentication_status, username = authenticator.login(location='main')

if authentication_status:

    st.sidebar.write("")
    st.sidebar.write("")
    st.sidebar.write("")
    st.sidebar.image(image_path, use_column_width=True, output_format='PNG')
    st.sidebar.write("")
    st.sidebar.markdown("<h1 style='font-size: 35px; font-weight: bold;'>QIS Square</h1>", unsafe_allow_html=True)

    # main_menu_options = ["Market", "국면", "유사국면", "모델전망 & Signal", "Allocation", "시나리오"]
    main_menu_options = ["Main", "Market", "Relative", "국면", "유사국면", "Macro 분석", "모델전망 & Signal", "DART공시정보 검색"]
    selected_main_menu = st.sidebar.selectbox("Select a Main Menu", main_menu_options)

    if selected_main_menu == "Main":
        sub_menu_options = ["Main", "PPT_QIS"]

    if selected_main_menu == "DART공시정보 검색":
        sub_menu_options = ["최근 공시정보 검색"]

    elif selected_main_menu == "Market":
        sub_menu_options = ["MarketBoard", "MarketChart", "주요국 만기별 금리"]

    elif selected_main_menu == "Relative":
        sub_menu_options = ["Relative(Trend)", "Relative(Momentum)", "Relative(Pair)", "현재위치"]

    elif selected_main_menu == "국면":
        sub_menu_options = ["Economic Cycle", "Credit Cycle"]

    elif selected_main_menu == "유사국면":
        sub_menu_options = ["유사국면분석"]

    elif selected_main_menu == "Macro 분석":
        sub_menu_options = ["Macro Driver", "Macro: Actual vs. Survey"]

    elif selected_main_menu == "모델전망 & Signal":
        sub_menu_options = ["금리", "USIG스프레드", "USIG 추천종목", "RankingModel", "FX", "FDS"]

    selected_sub_menu = st.sidebar.selectbox("Select a Sub Menu", sub_menu_options)

    if selected_main_menu == "Main":
        if selected_sub_menu == "Main":

            st.title("Quantamental Investment Strategy - 메뉴설명")
            st.write("")
            st.write("")
            st.subheader("Market: 금리, 환율, 주가지수 등의 현황 및 추이 조회")
            st.subheader("Relative: 둘 이상 지표간의 상대적 흐름 파악")
            st.subheader("국면: 퀀타멘탈운용본부에서 산출하는 Economic Cycle, Credit Cycle 조회")
            st.subheader("유사국면: 월간 단위로 산출하는 유사국면 정보 조회")
            st.subheader("Macro 분석: 매크로 지표와 가격 지표의 상관성 파악")
            st.subheader("모델전망 & Signal: 퀀타멘탈운용본부의 AI/Quant 모델 기반 예측정보 및 모델산출물 조회")
            st.subheader("DART공시정보 검색: 금감원 DART 공시자료 조회")

        elif selected_sub_menu == "PPT_QIS":

            def load_images(slidepath):
                image_files = sorted([f for f in os.listdir(slidepath) if f.endswith(".PNG")],
                                     key=lambda x: int(os.path.splitext(x)[0]))
                return image_files


            def change_image(direction):
                if direction == "next" and st.session_state.current_image_index < len(st.session_state.image_files) - 1:
                    st.session_state.current_image_index += 1
                elif direction == "previous" and st.session_state.current_image_index > 0:
                    st.session_state.current_image_index -= 1


            def go_to_page():
                page_number = st.session_state.page_number
                if 1 <= page_number <= len(st.session_state.image_files):
                    st.session_state.current_image_index = page_number - 1


            def main(slidepath = slidepath):

                if "image_files" not in st.session_state:
                    st.session_state.image_files = load_images(slidepath)

                if "current_image_index" not in st.session_state:
                    st.session_state.current_image_index = 0

                total_images = len(st.session_state.image_files)

                col1, col2, col3 = st.columns([2, 2, 6])
                with col1:
                    st.write(f"{st.session_state.current_image_index + 1} of {total_images}")
                    st.number_input("페이지 이동: ", min_value=1, max_value=total_images, value=1, step=1, key="page_number",
                                    on_change=go_to_page)

                current_image_file = os.path.join(slidepath,
                                                  st.session_state.image_files[st.session_state.current_image_index])

                image = Image.open(current_image_file)
                width, height = image.size
                new_width = int(width * 1.3)
                new_height = int(height * 1.3)
                image_resized = image.resize((new_width, new_height))

                col1, col2 = st.columns([8, 2])
                with col1:
                    #st.image(current_image_file, use_column_width=True, output_format='PNG')
                    st.image(image_resized, use_column_width=True)

                col1, col2, col3, col4 = st.columns([2, 4, 2, 2])
                with col1:
                    st.button("⬅️ Previous", on_click=change_image, args=("previous",))
                with col3:
                    st.button("Next ➡️", on_click=change_image, args=("next",))

            if __name__ == "__main__":
                main()

    if selected_main_menu == "DART공시정보 검색":
        if selected_sub_menu == "최근 공시정보 검색":

            st.title("금융감독원 DART API - 공시정보 검색")

            # DART API Key
            api_key = "a3b2b551fee2036c0ebeb01e412887bcb30e5962"

            # 회사 코드와 주식 코드 매핑
            def create_stock_to_corp_mapping(api_key):
                api_url = "https://opendart.fss.or.kr/api/corpCode.xml"
                params = {'crtfc_key': api_key}
                response = requests.get(api_url, params=params)

                if response.status_code == 200:
                    zip_data = response.content
                    with zipfile.ZipFile(BytesIO(zip_data)) as z:
                        xml_file = z.open(z.namelist()[0])
                        xml_data = xml_file.read()

                    root = ET.fromstring(xml_data)
                    mapping = {}
                    for corp in root.findall('.//list'):
                        stock_code = corp.find('stock_code').text.strip()
                        corp_code = corp.find('corp_code').text
                        if stock_code:
                            mapping[stock_code] = corp_code
                    return mapping
                else:
                    raise Exception(f"API 요청 실패: {response.status_code}")


            # 매핑
            stock_to_corp_mapping = create_stock_to_corp_mapping(api_key)

            col1, col2 = st.columns([4,1])
            with col1:
                # 검색 기준 선택
                search_type = st.radio("검색 기준을 선택하세요:", ("주식 코드", "기업 코드"), horizontal=True)

            # 코드 입력 받기
            if search_type == "주식 코드":
                codes = st.text_area("주식 코드를 입력하세요 (쉼표로 구분):", "005930, 000660, 005380")
            elif search_type == "기업 코드":
                codes = st.text_area("기업 코드를 입력하세요 (쉼표로 구분):", "00126380, 003550, 005930")

            days = st.number_input("최근 X일:", min_value=1, value=7)

            end_date = datetime.now().strftime('%Y%m%d')
            start_date = (datetime.now() - timedelta(days=days)).strftime('%Y%m%d')

            # 공시정보 검색 함수
            def get_disclosure_info(corp_code, start_date, end_date):
                api_url = "https://opendart.fss.or.kr/api/list.json"
                params = {
                    'crtfc_key': api_key,
                    'corp_code': corp_code,
                    'bgn_de': start_date,
                    'end_de': end_date,
                    'page_no': 1,
                    'page_count': 100
                }
                response = requests.get(api_url, params=params)

                if response.status_code == 200:
                    return response.json()
                else:
                    st.error(f"API 요청에 실패했습니다. 상태 코드: {response.status_code}")
                    return None

            # 검색 버튼
            if st.button("검색"):
                codes = [code.strip() for code in codes.split(',')]
                all_data = []

                for code in codes:
                    if search_type == "주식 코드":
                        corp_code = stock_to_corp_mapping.get(code)
                        if not corp_code:
                            st.warning(f"{code}에 대한 회사 코드가 없습니다.")
                            continue
                    else:
                        corp_code = code

                    json_data = get_disclosure_info(corp_code, start_date, end_date)
                    if json_data and 'list' in json_data:
                        for item in json_data['list']:
                            item['corp_code'] = corp_code
                            item['rcept_url'] = f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={item['rcept_no']}"
                            all_data.append(item)

                if all_data:
                    df = pd.DataFrame(all_data)

                    # 하이퍼링크 추가
                    df['report_nm'] = df.apply(
                        lambda row: f'<a href="{row["rcept_url"]}" target="_blank">{row["report_nm"]}</a>',
                        axis=1)

                    # Streamlit에 HTML 테이블 출력
                    st.markdown(df.to_html(escape=False, index=False), unsafe_allow_html=True)
                else:
                    st.error("공시정보를 가져올 수 없습니다.")

    if selected_main_menu == "Market":
        if selected_sub_menu == "MarketBoard":

            st.title("MarketBoard")

            st.subheader("Global 10Y")
            df1 = pd.read_excel(market_path, sheet_name='G10Y')
            st.write("as of: ", df1['DATE'].max())
            stable, grid_options = cal_table(df=df1, chgopt=2, spechk=1)
            AgGrid(stable, gridOptions=grid_options, fit_columns_on_grid_load=True, allow_unsafe_jscode=True)

            st.subheader("주요국 만기별 금리")
            df2 = pd.read_excel(market_path, sheet_name='Cntry')
            st.write("as of: ", df2['DATE'].max())
            stable, grid_options = cal_table(df=df2, chgopt=2, spechk=2)
            AgGrid(stable, gridOptions=grid_options, fit_columns_on_grid_load=True, allow_unsafe_jscode=True)

            st.subheader("Credit Spread")
            df3 = pd.read_excel(market_path, sheet_name='OAS')
            st.write("as of: ", df3['DATE'].max())
            stable, grid_options = cal_table(df=df3, chgopt=2, spechk=1)
            AgGrid(stable, gridOptions=grid_options, fit_columns_on_grid_load=True, allow_unsafe_jscode=True)

            st.subheader("FX")
            df4 = pd.read_excel(market_path, sheet_name='FX')
            st.write("as of: ", df4['DATE'].max())
            stable, grid_options = cal_table(df=df4, chgopt=1, spechk=1)
            AgGrid(stable, gridOptions=grid_options, fit_columns_on_grid_load=True, allow_unsafe_jscode=True)

            st.subheader("Stock Index")
            df5 = pd.read_excel(market_path, sheet_name='StockIndex')
            st.write("as of: ", df5['DATE'].max())
            stable, grid_options = cal_table(df=df5, chgopt=1, spechk=1)
            AgGrid(stable, gridOptions=grid_options, fit_columns_on_grid_load=True, allow_unsafe_jscode=True)

            st.subheader("SPX by Sector")
            df6 = pd.read_excel(market_path, sheet_name='SPXsector')
            st.write("as of: ", df6['DATE'].max())
            stable, grid_options = cal_table(df=df6, chgopt=1, spechk=1)
            AgGrid(stable, gridOptions=grid_options, fit_columns_on_grid_load=True, allow_unsafe_jscode=True)

            st.subheader("Commodity(S&P GSCI)")
            df7 = pd.read_excel(market_path, sheet_name='SPGSCI')
            st.write("as of: ", df7['DATE'].max())
            stable, grid_options = cal_table(df=df7, chgopt=1, spechk=1)
            AgGrid(stable, gridOptions=grid_options, fit_columns_on_grid_load=True, allow_unsafe_jscode=True)

            st.subheader("Energy(Oil & Gas)")
            df8 = pd.read_excel(market_path, sheet_name='Energy')
            st.write("as of: ", df8['DATE'].max())
            stable, grid_options = cal_table(df=df8, chgopt=1, spechk=1)
            AgGrid(stable, gridOptions=grid_options, fit_columns_on_grid_load=True, allow_unsafe_jscode=True)

        elif selected_sub_menu == "MarketChart":

            st.title("Market Chart")

            sel_cate = st.selectbox("Category",
                                    ['Global 10Y', 'Credit Spread', 'FX', 'StockIndex', 'SPX Sector', 'S&P GSCI', 'Energy', 'All'])

            def plot_ts(df, sel_colx, selecpr):
                edate = df['DATE'].max()

                if selecpr == "1M":
                    sdate = edate - pd.DateOffset(months=1)
                elif selecpr == "3M":
                    sdate = edate - pd.DateOffset(months=3)
                elif selecpr == "6M":
                    sdate = edate - pd.DateOffset(months=6)
                elif selecpr == "1Y":
                    sdate = edate - pd.DateOffset(years=1)
                elif selecpr == "3Y":
                    sdate = edate - pd.DateOffset(years=3)
                elif selecpr == "5Y":
                    sdate = edate - pd.DateOffset(years=5)
                else:
                    sdate = edate - pd.DateOffset(years=10)

                fdf = df[(df['DATE'] >= pd.to_datetime(sdate)) & (df['DATE'] <= pd.to_datetime(edate))]

                columns = st.columns(4)
                for i, col in enumerate(sel_colx):
                    xdf = fdf[['DATE', col]]
                    xdf = xdf.dropna()
                    fig1 = go.Figure()
                    fig1.add_trace(go.Scatter(x=xdf['DATE'], y=xdf[col], name=col, mode='lines', line=dict(color='rgb(245, 130, 32)')))
                    fig1.update_layout(
                        xaxis_title='Date',
                        yaxis_title=col,
                        template='plotly_dark'
                    )
                    with columns[i % 4]:
                        st.subheader(col)
                        st.plotly_chart(fig1)


            if sel_cate in ['Global 10Y', 'Credit Spread', 'FX', 'StockIndex', 'SPX Sector', 'S&P GSCI', 'Energy', 'All']:

                if sel_cate == "All":
                    df1 = pd.read_excel(market_path, sheet_name="G10Y")
                    df2 = pd.read_excel(market_path, sheet_name="OAS")
                    df3 = pd.read_excel(market_path, sheet_name="FX")
                    df4 = pd.read_excel(market_path, sheet_name="StockIndex")
                    df5 = pd.read_excel(market_path, sheet_name="SPXsector")
                    df6 = pd.read_excel(market_path, sheet_name="SPGSCI")
                    df7 = pd.read_excel(market_path, sheet_name="Energy")
                    dfs = [df1, df2, df3, df4, df5, df6, df7]
                    df = reduce(lambda left, right: pd.merge(left, right, on='DATE', how='outer'), dfs)

                else:
                    if sel_cate == "Global 10Y":
                        shtnm = "G10Y"
                    elif sel_cate == "Credit Spread":
                        shtnm = "OAS"
                    elif sel_cate == "FX":
                        shtnm = "FX"
                    elif sel_cate == "StockIndex":
                        shtnm = "StockIndex"
                    elif sel_cate == "SPX Sector":
                        shtnm = "SPXsector"
                    elif sel_cate == "S&P GSCI":
                        shtnm = "SPGSCI"
                    elif sel_cate == "Energy":
                        shtnm = "Energy"
                    df = pd.read_excel(market_path, sheet_name=shtnm)

                sel_cols = [col for col in df.columns if col != 'DATE']
                sel_colx = st.multiselect(
                    "Select:",
                    sel_cols,
                    default=sel_cols
                )

                if sel_cate != "All":
                    dfa = df[['DATE'] + sel_colx]
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        kdt = st.date_input("Date", min_value=dfa['DATE'].min(), max_value=dfa['DATE'].max(), value=dfa['DATE'].max())
                    with col2:
                        selchg = st.radio("등락계산 기준", ["Change", "Percent"], horizontal=True)

                    col1, col2 = st.columns(2)
                    with col1:
                        selpr1 = st.radio("기준구간", ["1W", "1M", "3M", "6M", "1Y", "3Y", "5Y", "10Y", "MTD", "YTD"],
                                          horizontal=True, index=0)
                    with col2:
                        selpr2 = st.radio("비교구간", ["1W", "1M", "3M", "6M", "1Y", "3Y", "5Y", "10Y", "MTD", "YTD"],
                                          horizontal=True, index=1)

                    if selchg == "Change":
                        chgoptx = 2
                    elif selchg == "Percent":
                        chgoptx = 1

                    dfb = dfa[(dfa['DATE'] <= pd.to_datetime(kdt))]
                    dfb['DATE'] = pd.to_datetime(dfb['DATE'])
                    stable, grid_options = cal_table(df=dfb, chgopt=chgoptx, spechk=1)
                    if selpr1 == selpr2:
                        stablex = stable[['X', selpr1]]
                    else:
                        stablex = stable[['X', selpr1, selpr2]]
                    stablex = stablex.sort_values(by=selpr1, ascending=False)

                    figb = go.Figure()
                    figb.add_trace(go.Bar(
                        x=stablex['X'],
                        y=stablex[selpr1],
                        name=selpr1,
                        marker=dict(
                            color='rgb(245,130,32)',
                            line=dict(
                                width=0
                            )
                        )
                    ))
                    if len(stablex.columns) > 2:
                        figb.add_trace(go.Bar(
                            x=stablex['X'],
                            y=stablex[selpr2],
                            name=selpr2,
                            marker=dict(
                                color='rgb(13,45,79)',
                                line=dict(
                                    width=0
                                )
                            )
                        ))
                    figb.update_layout(
                        barmode='group' if len(stablex.columns) > 2 else 'overlay',
                        xaxis_title=sel_cate,
                        yaxis_title='Chg/Pct',
                        title=''
                    )
                    st.plotly_chart(figb)

                selecpr = st.radio("", ["1M", "3M", "6M", "1Y", "3Y", "5Y", "10Y"], horizontal=True)
                plot_ts(df, sel_colx, selecpr)

        elif selected_sub_menu == "주요국 만기별 금리":

            st.title("주요국 만기별 금리")
            df = pd.read_excel(market_path, sheet_name='Cntry')
            st.write("as of: ", df['DATE'].max())

            sel_cntry = st.selectbox("Country",
                                     ['US', 'UK', 'Germany', 'Italy', 'Jppan', 'China', 'Australia'])

            sdt = st.date_input("Start", min_value=df['DATE'].min(), max_value=df['DATE'].max(), value=df['DATE'].min())

            st.write("")


            sels = ['DATE'] + [col for col in df.columns if col.startswith(sel_cntry)]
            dfx = df[sels]
            dfx = dfx[(dfx['DATE'] >= pd.to_datetime(sdt))]

            col1, col2 = st.columns(2)
            with col1:
                st.subheader(f"{sel_cntry}: 만기별 금리")
                sel_cols = [col for col in dfx.columns if col != 'DATE']
                sel_colx = st.multiselect(
                    "Select:",
                    sel_cols,
                    default=sel_cols
                )
                dfxx = dfx[['DATE'] + sel_colx]

                fig1 = go.Figure()
                for col in dfxx.columns:
                    if col != 'DATE':
                        fig1.add_trace(go.Scatter(x=dfxx['DATE'], y=dfxx[col], name=col, mode='lines'))
                st.plotly_chart(fig1)

            with col2:
                st.subheader(f"{sel_cntry}: 장단기 스프레드 추이")
                if sel_cntry != "China":
                    xcol = [col for col in dfx.columns if
                           col.startswith(sel_cntry) and col[len(sel_cntry):] in [' 2Y', ' 5Y', ' 10Y', ' 30Y']]
                    xcol = ['DATE'] + xcol
                    dfy = dfx[xcol]
                    dfy['Spr_2_5'] = dfy.iloc[:, 2] - dfy.iloc[:, 1]
                    dfy['Spr_2_10'] = dfy.iloc[:, 3] - dfy.iloc[:, 1]
                    dfy['Spr_2_30'] = dfy.iloc[:, 4] - dfy.iloc[:, 1]
                    dfy['Spr_5_10'] = dfy.iloc[:, 3] - dfy.iloc[:, 2]
                    dfy['Spr_5_30'] = dfy.iloc[:, 4] - dfy.iloc[:, 2]
                    dfy['Spr_10_30'] = dfy.iloc[:, 4] - dfy.iloc[:, 3]
                    dfy = dfy[['DATE', 'Spr_2_5', 'Spr_2_10', 'Spr_2_30', 'Spr_5_10', 'Spr_5_30', 'Spr_10_30']]
                elif sel_cntry == "China":
                    xcol = [col for col in dfx.columns if
                            col.startswith(sel_cntry) and col[len(sel_cntry):] in [' 2Y', ' 5Y', ' 10Y']]
                    xcol = ['DATE'] + xcol
                    dfy = dfx[xcol]
                    dfy['Spr_2_5'] = dfy.iloc[:, 2] - dfy.iloc[:, 1]
                    dfy['Spr_2_10'] = dfy.iloc[:, 3] - dfy.iloc[:, 1]
                    dfy['Spr_5_10'] = dfy.iloc[:, 3] - dfy.iloc[:, 2]
                    dfy = dfy[['DATE', 'Spr_2_5', 'Spr_2_10', 'Spr_5_10']]

                dfyy = dfy[(dfy['DATE'] >= pd.to_datetime(sdt))]
                sel_cols = [col for col in dfyy.columns if col != 'DATE']
                sel_coly = st.multiselect(
                    "Select:",
                    sel_cols,
                    default=sel_cols
                )
                dfyyy = dfyy[['DATE'] + sel_coly]

                fig1 = go.Figure()
                for col in dfyyy.columns:
                    if col != 'DATE':
                        fig1.add_trace(go.Scatter(x=dfyyy['DATE'], y=dfyyy[col], name=col, mode='lines'))

                st.plotly_chart(fig1)

            dfxx['DATE'] = pd.to_datetime(dfxx['DATE'])
            df_frix = dfxx[dfxx['DATE'].dt.dayofweek == 4]
            df_frifx = df_frix.tail(4)
            colb = [col for col in df_frifx.columns if col != 'DATE']

            figx = make_subplots(rows=1, cols=len(colb), shared_yaxes=True,
                                subplot_titles=colb)
            for i, col in enumerate(colb, start=1):
                ftextx = df_frifx[col].apply(lambda x: f'{x:.2f}')
                figx.add_trace(go.Bar(x=df_frifx['DATE'], y=df_frifx[col], name=col, text=ftextx, textposition='outside'),
                              row=1, col=i)
            figx.update_layout(
                title="만기별 금리: 최근 4주 추이",
                xaxis_title="Date",
                yaxis_title="금리",
                barmode='group',
                height=400
            )
            st.plotly_chart(figx)

            dfyyy['DATE'] = pd.to_datetime(dfyyy['DATE'])
            df_fri = dfyyy[dfyyy['DATE'].dt.dayofweek == 4]
            df_frif = df_fri.tail(4)
            cola = [col for col in df_frif.columns if col != 'DATE']

            fig = make_subplots(rows=1, cols=len(cola), shared_yaxes=True,
                                subplot_titles=cola)
            for i, col in enumerate(cola, start=1):
                ftext = df_frif[col].apply(lambda x: f'{x:.3f}')
                fig.add_trace(go.Bar(x=df_frif['DATE'], y=df_frif[col], name=col, text=ftext, textposition='outside'),
                              row=1, col=i)
            fig.update_layout(
                title="스프레드: 최근 4주 추이",
                xaxis_title="Date",
                yaxis_title="Spr",
                barmode='group',
                height=400
            )
            st.plotly_chart(fig)

    if selected_main_menu == "Relative":

        if selected_sub_menu == "Relative(Trend)":

            st.title("Relative(Trend)")
            col1, col2, col3, col4, col5 = st.columns(5)

            with col1:
                sel_cate1 = st.selectbox("Category1",
                                         ['Global 10Y', 'Credit Spread', 'FX', 'StockIndex', 'SPX Sector', 'S&P GSCI', 'Energy'])

            with col2:
                if sel_cate1 == "Global 10Y":
                    df1 = pd.read_excel(market_path, sheet_name="G10Y")
                elif sel_cate1 == "Credit Spread":
                    df1 = pd.read_excel(market_path, sheet_name="OAS")
                elif sel_cate1 == "FX":
                    df1 = pd.read_excel(market_path, sheet_name="FX")
                elif sel_cate1 == "StockIndex":
                    df1 = pd.read_excel(market_path, sheet_name="StockIndex")
                elif sel_cate1 == "SPX Sector":
                    df1 = pd.read_excel(market_path, sheet_name="SPXsector")
                elif sel_cate1 == "S&P GSCI":
                    df1 = pd.read_excel(market_path, sheet_name="SPGSCI")
                elif sel_cate1 == "Energy":
                    df1 = pd.read_excel(market_path, sheet_name="Energy")

                columns = [col for col in df1.columns if col != 'DATE']
                sel_cate11 = st.selectbox("Sel1(Y)", columns)
                dfx1 = df1[['DATE', sel_cate11]]
                dfx1 = dfx1.dropna()

            with col3:
                sel_cate2 = st.selectbox("Category2",
                                         ['Global 10Y', 'Credit Spread', 'FX', 'StockIndex', 'SPX Sector', 'S&P GSCI', 'Energy'])

            with col4:
                if sel_cate2 == "Global 10Y":
                    df1 = pd.read_excel(market_path, sheet_name="G10Y")
                elif sel_cate2 == "Credit Spread":
                    df1 = pd.read_excel(market_path, sheet_name="OAS")
                elif sel_cate2 == "FX":
                    df1 = pd.read_excel(market_path, sheet_name="FX")
                elif sel_cate2 == "StockIndex":
                    df1 = pd.read_excel(market_path, sheet_name="StockIndex")
                elif sel_cate2 == "SPX Sector":
                    df1 = pd.read_excel(market_path, sheet_name="SPXsector")
                elif sel_cate2 == "S&P GSCI":
                    df1 = pd.read_excel(market_path, sheet_name="SPGSCI")
                elif sel_cate2 == "Energy":
                    df1 = pd.read_excel(market_path, sheet_name="Energy")

                columns = [col for col in df1.columns if col != 'DATE']
                sel_cate21 = st.selectbox("Sel2(X)", columns)
                dfx2 = df1[['DATE', sel_cate21]]
                dfx2 = dfx2.dropna()

            col1, col2, col3, col4 = st.columns(4)

            if sel_cate11 != '선택 없음':
                with col1:
                    sdate = dfx1['DATE'].min().strftime('%Y/%m/%d')
                    edate = dfx1['DATE'].max().strftime('%Y/%m/%d')
                    st.subheader(f"Date: {sdate} ~ {edate}")
                    start_date = st.date_input("Start", min_value=dfx1['DATE'].min(), max_value=dfx1['DATE'].max(),
                                               value=dfx1['DATE'].min())
                    st.write("")
                with col2:
                    st.subheader("")
                    end_date = st.date_input("End", min_value=dfx1['DATE'].min(), max_value=dfx1['DATE'].max(),
                                             value=dfx1['DATE'].max())
                    st.write("")

                fdf = dfx1[(dfx1['DATE'] >= pd.to_datetime(start_date)) & (dfx1['DATE'] <= pd.to_datetime(end_date))]
                st.subheader(f"{sel_cate11}")
                recent_data1 = fdf[['DATE', sel_cate11]].tail(5)
                recent_data1.set_index('DATE', inplace=True)

                col1, col2 = st.columns([4, 1])
                with col1:
                    st.dataframe(recent_data1, use_container_width=True)
                    fig1 = go.Figure()
                    fig1.add_trace(
                        go.Scatter(x=fdf['DATE'], y=fdf[sel_cate11], name=sel_cate11, mode='lines', line=dict(color='rgb(245, 130, 32)')))
                    fig1.update_layout(
                        xaxis_title='Date',
                        yaxis_title=sel_cate11,
                        template='plotly_dark'
                    )
                    st.plotly_chart(fig1, use_container_width=True)

                if sel_cate21 != '선택 없음' and sel_cate11 != sel_cate21:
                    col1, col2, col3 = st.columns([2, 2, 1])
                    with col1:
                        fdf = pd.merge(dfx1, dfx2, on='DATE', how='inner')
                        corr = fdf[sel_cate11].corr(fdf[sel_cate21])

                        X = fdf[[sel_cate21]]
                        y = fdf[sel_cate11]
                        model = LinearRegression()
                        model.fit(X, y)
                        reg_coeff = model.coef_[0]

                        st.write("")
                        st.subheader(f"{sel_cate11} & {sel_cate21}")
                        st.write("Correlation:", round(corr, 4))
                        st.write("Regression Coefficient:", round(reg_coeff, 2))

                        recent_data2 = fdf[['DATE', sel_cate11, sel_cate21]].tail(5)
                        recent_data2.set_index('DATE', inplace=True)
                        st.dataframe(recent_data2, use_container_width=True)

                        fig2 = go.Figure()
                        fig2.add_trace(
                            go.Scatter(x=fdf['DATE'], y=fdf[sel_cate11], name=sel_cate11, mode='lines', line=dict(color='rgb(245, 130, 32)')))
                        fig2.add_trace(
                            go.Scatter(x=fdf['DATE'], y=fdf[sel_cate21], name=sel_cate21, mode='lines', line=dict(color='rgb(13, 45, 79)'),
                                       yaxis='y2'))

                        fig2.update_layout(
                            xaxis_title='Date',
                            yaxis_title=sel_cate11,
                            yaxis2=dict(
                                title=sel_cate21,
                                overlaying='y',
                                side='right'
                            ),
                            template='plotly_dark',
                            legend=dict(
                                orientation='h',
                                yanchor='top',
                                y=1.1,
                                xanchor='center',
                                x=0.5
                            )
                        )
                        st.plotly_chart(fig2, use_container_width=True)

                    with col2:

                        selectr = st.radio("Relative:",
                                           ["Spread", "Ratio"]
                                           )

                        if selectr == "Spread":
                            fdf['rel'] = fdf[sel_cate11] - fdf[sel_cate21]
                            st.subheader(f"Spr({sel_cate11}-{sel_cate21})")
                        elif selectr == "Ratio":
                            fdf['rel'] = fdf[sel_cate11] / fdf[sel_cate21]
                            st.subheader(f"Ratio({sel_cate11}/{sel_cate21})")

                        recent_data3 = fdf[['DATE', 'rel']].tail(5)
                        recent_data3.set_index('DATE', inplace=True)
                        st.dataframe(recent_data3, use_container_width=True)

                        fig3 = go.Figure()
                        fig3.add_trace(go.Scatter(x=fdf['DATE'], y=fdf['rel'], name='relative', mode='lines',
                                                  line=dict(color='orange')))
                        fig3.update_layout(
                            xaxis_title='Date',
                            yaxis_title='relative',
                            template='plotly_dark'
                        )
                        st.plotly_chart(fig3, use_container_width=True)


            def convert_df_to_csv(df):
                return df.to_csv(index=False).encode('utf-8')

            # 선택된 열의 데이터만 추출
            if sel_cate11 != '선택 없음' and sel_cate21 != '선택 없음' and sel_cate11 != sel_cate21:
                data_to_download = fdf[['DATE', sel_cate11, sel_cate21, 'rel']]
                csv_data = convert_df_to_csv(data_to_download)
                st.download_button(
                    label="Data Download(CSV)",
                    data=csv_data,
                    file_name='timeseries_data.csv',
                    mime='text/csv'
                )
            elif ((sel_cate11 != '선택 없음' and sel_cate21 == '선택 없음') or
                  (sel_cate11 != '선택 없음' and sel_cate11 == sel_cate21)):
                data_to_download = fdf[['DATE', sel_cate11]]
                csv_data = convert_df_to_csv(data_to_download)
                st.download_button(
                    label="Data Download(CSV)",
                    data=csv_data,
                    file_name='timeseries_data.csv',
                    mime='text/csv'
                )
            else:
                pass

        elif selected_sub_menu == "Relative(Momentum)":

            st.title("Relative(Momentum)")
            col1, col2 = st.columns(2)

            with col1:
                sel_cate = st.selectbox("Category",
                                        ['Global 10Y', 'Credit Spread', 'FX', 'StockIndex', 'SPX Sector', 'S&P GSCI', 'Energy'])

                if sel_cate == "Global 10Y":
                    df = pd.read_excel(market_path, sheet_name="G10Y")
                    chgopt = 2
                elif sel_cate == "Credit Spread":
                    df = pd.read_excel(market_path, sheet_name="OAS")
                    chgopt = 2
                elif sel_cate == "FX":
                    df = pd.read_excel(market_path, sheet_name="FX")
                    chgopt = 1
                elif sel_cate == "StockIndex":
                    df = pd.read_excel(market_path, sheet_name="StockIndex")
                    chgopt = 1
                elif sel_cate == "SPX Sector":
                    df = pd.read_excel(market_path, sheet_name="SPXsector")
                    chgopt = 1
                elif sel_cate == "S&P GSCI":
                    df = pd.read_excel(market_path, sheet_name="SPGSCI")
                    chgopt = 1
                elif sel_cate == "Energy":
                    df = pd.read_excel(market_path, sheet_name="Energy")
                    chgopt = 1

            col1, col2 = st.columns(2)

            with col1:
                selecpr1 = st.radio("X-axis", ["1W", "1M", "3M", "6M", "1Y", "3Y", "5Y", "10Y", "MTD", "YTD"], horizontal=True)
            with col2:
                selecpr2 = st.radio("Y-axis", ["1W", "1M", "3M", "6M", "1Y", "3Y", "5Y", "10Y", "MTD", "YTD"], horizontal=True)

            if selecpr1 == "1W":
                chg1 = calculate_change_with_offset(df, chgopt, days_offset=5)
            elif selecpr1 == "1M":
                chg1 = calculate_change_with_offset(df, chgopt, months_offset=1)
            elif selecpr1 == "3M":
                chg1 = calculate_change_with_offset(df, chgopt, months_offset=3)
            elif selecpr1 == "6M":
                chg1 = calculate_change_with_offset(df, chgopt, months_offset=6)
            elif selecpr1 == "1Y":
                chg1 = calculate_change_with_offset(df, chgopt, years_offset=1)
            elif selecpr1 == "2Y":
                chg1 = calculate_change_with_offset(df, chgopt, years_offset=2)
            elif selecpr1 == "3Y":
                chg1 = calculate_change_with_offset(df, chgopt, years_offset=3)
            elif selecpr1 == "5Y":
                chg1 = calculate_change_with_offset(df, chgopt, years_offset=5)
            elif selecpr1 == "10Y":
                chg1 = calculate_change_with_offset(df, chgopt, years_offset=10)
            elif selecpr1 == "MTD":
                chg1 = calculate_change_with_offset(df, chgopt, days_offset=df['DATE'].iloc[-1].day - 1)
            elif selecpr1 == "YTD":
                chg1 = calculate_change_with_offset(df, chgopt,days_offset=df['DATE'].iloc[-1].timetuple().tm_yday - 1)

            chg1 = pd.DataFrame(chg1).reset_index()
            chg1.rename(columns={chg1.columns[0]: 'X', chg1.columns[1]: f"X_{selecpr1}"}, inplace=True)

            if selecpr2 == "1W":
                chg2 = calculate_change_with_offset(df, chgopt, days_offset=5)
            elif selecpr2 == "1M":
                chg2 = calculate_change_with_offset(df, chgopt, months_offset=1)
            elif selecpr2 == "3M":
                chg2 = calculate_change_with_offset(df, chgopt, months_offset=3)
            elif selecpr2 == "6M":
                chg2 = calculate_change_with_offset(df, chgopt, months_offset=6)
            elif selecpr2 == "1Y":
                chg2 = calculate_change_with_offset(df, chgopt, years_offset=1)
            elif selecpr2 == "2Y":
                chg2 = calculate_change_with_offset(df, chgopt, years_offset=2)
            elif selecpr2 == "3Y":
                chg2 = calculate_change_with_offset(df, chgopt, years_offset=3)
            elif selecpr2 == "5Y":
                chg2 = calculate_change_with_offset(df, chgopt, years_offset=5)
            elif selecpr2 == "10Y":
                chg2 = calculate_change_with_offset(df, chgopt, years_offset=10)
            elif selecpr2 == "MTD":
                chg2 = calculate_change_with_offset(df, chgopt, days_offset=df['DATE'].iloc[-1].day - 1)
            elif selecpr2 == "YTD":
                chg2 = calculate_change_with_offset(df, chgopt, days_offset=df['DATE'].iloc[-1].timetuple().tm_yday - 1)

            chg2 = pd.DataFrame(chg2).reset_index()
            chg2.rename(columns={chg2.columns[0]: 'X', chg2.columns[1]: f"Y_{selecpr2}"}, inplace=True)

            chgx = pd.merge(chg1, chg2, on='X', how='inner')
            names = chgx['X'].unique()
            sel_nms = st.multiselect("Select", options=names, default=names)
            chgxf = chgx[chgx['X'].isin(sel_nms)]

            X = chgxf[[f"X_{selecpr1}"]].values
            y = chgxf[f"Y_{selecpr2}"].values
            model = LinearRegression()
            model.fit(X, y)
            slope = model.coef_[0]
            intercept = model.intercept_
            x_values = np.linspace(chgxf[f"X_{selecpr1}"].min(), chgxf[f"X_{selecpr1}"].max(), 100)
            y_values = slope * x_values + intercept

            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=chgxf[f"X_{selecpr1}"],
                y=chgxf[f"Y_{selecpr2}"],
                mode='markers+text',
                text=chgxf['X'],
                textposition='top right',
                marker=dict(size=10)
            ))
            fig.add_trace(go.Scatter(
                x=x_values,
                y=y_values,
                mode='lines',
                line=dict(color='red', dash='dash'),
                name='Regression Line'
            ))

            fig.update_layout(
                title='Scatter Plot',
                xaxis_title=f"X_{selecpr1}",
                yaxis_title=f"Y_{selecpr2}",
                template='plotly_dark',
                height=800
            )
            st.plotly_chart(fig)

        elif selected_sub_menu == "Relative(Pair)":

            st.title("Relative(Pair)")

            df1 = pd.read_excel(market_path, sheet_name='G10Y')
            df2 = pd.read_excel(market_path, sheet_name='FX')
            df3 = pd.read_excel(market_path, sheet_name='Cntry')
            df3 = df3.loc[:, ~df3.columns.str.endswith('10Y')]
            dfs = [df1, df2, df3]
            df = reduce(lambda left, right: pd.merge(left, right, on='DATE', how='outer'), dfs)
            dfx = df[df["DATE"].dt.weekday == 4]

            dffil = pd.read_csv(pairres_path)
            dffil_lst = dffil.copy()
            dffil_lst.insert(0, 'Pair', dffil_lst['Col_X'] + ' - ' + dffil_lst['Col_Y'])

            st.dataframe(dffil)
            html = """
                                <style>
                                    .custom-text {
                                        line-height: 1.2;
                                    }
                                </style>
                                <div class="custom-text">
                                    <p>Col_X, Col_Y가 FX이면 rvalue는 ratio, 금리이면 diff(Col_X - Col_Y)</p>
                                    <p>rvalue가 +(-) 이면서 upper(lower)보다 높다면(낮다면), Col_X(Col_Y)의 레벨이 abnormally high</p>
                                    <p>ortho_resid도 같은 해석</p>                        
                                </div>
                                """
            st.markdown(html, unsafe_allow_html=True)
            st.write("")

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                na = st.radio(label="분석기간", options=["26", "52", "104", "156"], horizontal=True)

            ana_n = int(na)

            col1, col2 = st.columns(2)
            pair0 = dffil_lst[dffil_lst['N_analysis'] == ana_n]
            pairlst = pair0['Pair'].tolist()
            with col1:
                sel_pairlst = st.selectbox("Pair 선택", pairlst)

            dffil_x = pair0[pair0['Pair'] == sel_pairlst]
            st.dataframe(dffil_x)
            st.write("")

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                last_n = st.number_input(label="관찰기간", value=52, step=1)

            sel_case = pair0[pair0['Pair'] == sel_pairlst]
            x_col = sel_case['Col_X'].iloc[0]
            y_col = sel_case['Col_Y'].iloc[0]

            chgopt = 1 if x_col in df2.columns else 2

            dfa = dfx[['DATE', x_col, y_col]]

            df_ratio = dfa.copy()
            if chgopt == 1:
                df_ratio['ratio'] = df_ratio.iloc[:, 1] / df_ratio.iloc[:, 2]
                df_ratio['ratio_upper'] = (df_ratio.iloc[:, 3].rolling(window=ana_n).mean()
                                           + 2 * df_ratio.iloc[:, 3].rolling(window=ana_n).std())
                df_ratio['ratio_lower'] = (df_ratio.iloc[:, 3].rolling(window=ana_n).mean()
                                           - 2 * df_ratio.iloc[:, 3].rolling(window=ana_n).std())
            elif chgopt == 2:
                df_ratio['diff'] = df_ratio.iloc[:, 1] - df_ratio.iloc[:, 2]
                df_ratio['diff_upper'] = (df_ratio.iloc[:, 3].rolling(window=ana_n).mean()
                                          + 2 * df_ratio.iloc[:, 3].rolling(window=ana_n).std())
                df_ratio['diff_lower'] = (df_ratio.iloc[:, 3].rolling(window=ana_n).mean()
                                          - 2 * df_ratio.iloc[:, 3].rolling(window=ana_n).std())

            f_ratio = df_ratio.tail(last_n)

            df_ortho = dfa.copy()
            df_ortho['DATE'] = pd.to_datetime(df_ortho['DATE'])
            df_ortho = df_ortho.sort_values(by='DATE', ascending=False)

            f_ortho = []
            for i in range(last_n - 1, -1, -1):
                ana_s = i + ana_n
                dfr = df_ortho.iloc[i:ana_s]
                dfr = dfr.sort_values(by='DATE', ascending=True)
                x = dfr.iloc[:, 1].values
                y = dfr.iloc[:, 2].values

                corr_m = np.corrcoef(x, y)
                corr = corr_m[0, 1]


                def linear_function(beta, x):
                    return beta[0] * x + beta[1]


                real_data = RealData(x, y)
                linear_model = Model(linear_function)
                odr = ODR(real_data, linear_model, beta0=[1., 0.])  # 초기값 beta0: [기울기, 절편]
                output = odr.run()
                intercept = output.beta[1]
                slope = output.beta[0]
                dfr['pred'] = dfr.iloc[:, 1] * slope + intercept
                dfr['ortho_resid'] = dfr['pred'] - dfr.iloc[:, 2]
                dfr['ortho_upper'] = dfr['ortho_resid'].mean() + 2 * dfr['ortho_resid'].std()
                dfr['ortho_lower'] = dfr['ortho_resid'].mean() - 2 * dfr['ortho_resid'].std()
                dfr['corr'] = corr
                dfr['ortho_slope'] = slope
                dfr['ortho_inter'] = intercept
                dfr = dfr.tail(1)
                dfr = dfr[['DATE', 'ortho_resid', 'ortho_upper', 'ortho_lower', 'corr', 'ortho_slope', 'ortho_inter']]
                f_ortho.append(dfr)

            f_ortho = pd.concat(f_ortho, ignore_index=True)

            f_final = pd.merge(f_ratio, f_ortho, on='DATE', how='inner')

            if chgopt == 1:
                f_final = f_final[['DATE', x_col, y_col, 'ratio_lower', 'ratio', 'ratio_upper',
                                   'ortho_lower', 'ortho_resid', 'ortho_upper']]
            if chgopt == 2:
                f_final = f_final[['DATE', x_col, y_col, 'diff_lower', 'diff', 'diff_upper',
                                   'ortho_lower', 'ortho_resid', 'ortho_upper']]

            df = f_final.copy()

            fig1 = go.Figure()
            fig1.add_trace(go.Scatter(x=df['DATE'], y=df[x_col], mode='lines', name=x_col, yaxis='y1',
                                      line=dict(color='rgb(245, 130, 32)')))
            fig1.add_trace(go.Scatter(x=df['DATE'], y=df[y_col], mode='lines', name=y_col, yaxis='y2',
                                      line=dict(color='rgb(13, 45, 79)')))
            fig1.update_layout(
                title=f"{x_col} & {y_col}",
                xaxis=dict(title='DATE'),
                yaxis=dict(title=x_col, side='left'),
                yaxis2=dict(title=y_col, side='right', overlaying='y', anchor='x')
            )

            slope, intercept, r_value, p_value, std_err = linregress(df[x_col], df[y_col])
            regression_line_x = df[x_col]
            regression_line_y = slope * regression_line_x + intercept

            fig2 = go.Figure()
            fig2.add_trace(
                go.Scatter(x=df[x_col], y=df[y_col], mode='markers', marker=dict(color='rgb(245, 130, 32)', size=5)))
            fig2.add_trace(go.Scatter(x=[df[x_col].iloc[-1]], y=[df[y_col].iloc[-1]],
                                      mode='markers', marker=dict(color='rgb(13, 45, 79)', size=15)))
            fig2.add_trace(go.Scatter(x=regression_line_x, y=regression_line_y, mode='lines',
                                      line=dict(color='rgb(0, 0, 0)', width=2)))
            fig2.update_layout(title="Scatter Plot", xaxis_title=x_col, yaxis_title=y_col)

            if chgopt == 1:
                fig3 = go.Figure()
                fig3.add_trace(go.Scatter(x=df['DATE'], y=df['ratio_lower'], mode='lines', name='ratio_lower',
                                          line=dict(color='rgb(13, 45, 79)')))
                fig3.add_trace(go.Scatter(x=df['DATE'], y=df['ratio'], mode='lines', name='ratio',
                                          line=dict(width=5, color='rgb(245, 130, 32)')))
                fig3.add_trace(go.Scatter(x=df['DATE'], y=df['ratio_upper'], mode='lines', name='ratio_upper',
                                          line=dict(color='rgb(13, 45, 79)')))
                fig3.update_layout(title="Ratio", xaxis_title='DATE')
            elif chgopt == 2:
                fig3 = go.Figure()
                fig3.add_trace(go.Scatter(x=df['DATE'], y=df['diff_lower'], mode='lines', name='ratio_lower',
                                          line=dict(color='rgb(13, 45, 79)')))
                fig3.add_trace(go.Scatter(x=df['DATE'], y=df['diff'], mode='lines', name='ratio',
                                          line=dict(width=5, color='rgb(245, 130, 32)')))
                fig3.add_trace(go.Scatter(x=df['DATE'], y=df['diff_upper'], mode='lines', name='ratio_upper',
                                          line=dict(color='rgb(13, 45, 79)')))
                fig3.update_layout(title="Diff", xaxis_title='DATE')

            fig4 = go.Figure()
            fig4.add_trace(go.Scatter(x=df['DATE'], y=df['ortho_lower'], mode='lines', name='ortho_lower',
                                      line=dict(color='rgb(13, 45, 79)')))
            fig4.add_trace(go.Scatter(x=df['DATE'], y=df['ortho_resid'], mode='lines', name='ortho_resid',
                                      line=dict(width=5, color='rgb(245, 130, 32)')))
            fig4.add_trace(go.Scatter(x=df['DATE'], y=df['ortho_upper'], mode='lines', name='ortho_upper',
                                      line=dict(color='rgb(13, 45, 79)')))
            fig4.update_layout(title="Orthogonal Regression Residuals", xaxis_title='DATE')

            col1, col2, col3 = st.columns([2, 2, 1])
            with col1:
                st.plotly_chart(fig1)
                st.plotly_chart(fig3)
            with col2:
                st.plotly_chart(fig2)
                st.plotly_chart(fig4)

        elif selected_sub_menu == "현재위치":

            st.title("현재위치")
            col1, col2 = st.columns(2)

            with col1:
                sel_cate = st.selectbox("Category",
                                        ['Global 10Y', 'Credit Spread', 'FX', 'StockIndex', 'SPX Sector', 'S&P GSCI', 'Energy'])

                if sel_cate == "Global 10Y":
                    df = pd.read_excel(market_path, sheet_name="G10Y")
                elif sel_cate == "Credit Spread":
                    df = pd.read_excel(market_path, sheet_name="OAS")
                elif sel_cate == "FX":
                    df = pd.read_excel(market_path, sheet_name="FX")
                elif sel_cate == "StockIndex":
                    df = pd.read_excel(market_path, sheet_name="StockIndex")
                elif sel_cate == "SPX Sector":
                    df = pd.read_excel(market_path, sheet_name="SPXsector")
                elif sel_cate == "S&P GSCI":
                    df = pd.read_excel(market_path, sheet_name="SPGSCI")
                elif sel_cate == "Energy":
                    df = pd.read_excel(market_path, sheet_name="Energy")

            col1, col2 = st.columns(2)
            with col1:
                selecpr = st.radio("분석기간", ["1Y", "3Y", "5Y", "10Y", "Max"], horizontal=True)

                edate = df['DATE'].max()
                if selecpr == "1Y":
                    sdate = edate - pd.DateOffset(years=1)
                elif selecpr == "3Y":
                    sdate = edate - pd.DateOffset(years=3)
                elif selecpr == "5Y":
                    sdate = edate - pd.DateOffset(years=5)
                elif selecpr == "10Y":
                    sdate = edate - pd.DateOffset(years=10)
                elif selecpr == "Max":
                    sdate = df['DATE'].min()

                fdf = df[(df['DATE'] >= pd.to_datetime(sdate)) & (df['DATE'] <= pd.to_datetime(edate))]

            with col2:
                st.write("분석날짜")
                st.write(sdate, "~", edate)


            def genbar(df, min_value, max_value, current_value, curv_trans, title):
                percentile = (current_value - min_value) / (max_value - min_value) * 100

                fig = go.Figure()
                fig.add_trace(go.Scatter(
                    x=[0, 1], y=[0, 0],
                    mode='lines', line=dict(color='gray', width=2), showlegend=False
                ))
                fig.add_trace(go.Scatter(
                    x=[curv_trans], y=[0],
                    mode='markers', marker=dict(color='red', size=10), showlegend=False
                ))
                fig.add_annotation(
                    x=0, y=0,
                    text=title,
                    showarrow=False, font=dict(size=16), bgcolor="white", xanchor='right', yanchor='bottom',
                    yshift=20
                )
                fig.add_annotation(
                    x=0, y=0,
                    text=f'{min_value:.2f}',
                    showarrow=True, arrowhead=2, ax=-30, ay=0, font=dict(size=16), bgcolor="white"
                )
                fig.add_annotation(
                    x=1, y=0,
                    text=f'{max_value:.2f}',
                    showarrow=True, arrowhead=2, ax=30, ay=0, font=dict(size=16), bgcolor="white"
                )
                fig.add_annotation(
                    x=curv_trans, y=0,
                    text=f'{current_value:.2f} ({percentile:.2f}%)',
                    showarrow=False, font=dict(size=16), bgcolor="white", xanchor='left', yshift=20
                )
                fig.update_layout(
                    showlegend=False,
                    xaxis=dict(
                        range=[-0.1, 1.1],
                        title='Value', title_text='', showticklabels=True, visible=False, zeroline=False
                    ),
                    yaxis=dict(
                        visible=False, zeroline=False
                    ),
                    height=150,
                    width=1200
                )

                return fig

            sel_cols = [col for col in fdf.columns if col != 'DATE']
            sel_colx = st.multiselect(
                "Select:",
                sel_cols,
                default=sel_cols
            )

            fdfx = fdf[sel_colx]
            currv = fdfx.iloc[-1]

            if not fdfx.empty:
                for col in sel_colx:
                    min_value = fdfx[col].min()
                    max_value = fdfx[col].max()
                    current_value = currv[col]
                    fdfn = (fdfx[col] - min_value) / (max_value - min_value)
                    fdfn = pd.DataFrame(fdfn)
                    fdfn.columns = [col]
                    curv_trans = fdfn.iloc[-1].values.item()
                    fig = genbar(fdfn, min_value, max_value, current_value, curv_trans, title=col)
                    #st.write(min_value, max_value, current_value, curv_trans)
                    st.plotly_chart(fig)
            else:
                st.write('No variables selected or no data available.')

    elif selected_main_menu == "국면":
        if selected_sub_menu == "Economic Cycle":

            st.title("Economic Cycle")

            df1 = pd.read_excel(market_path, sheet_name="G10Y")
            df2 = pd.read_excel(market_path, sheet_name="OAS")
            df3 = pd.read_excel(market_path, sheet_name="FX")
            df4 = pd.read_excel(market_path, sheet_name="StockIndex")
            df5 = pd.read_excel(market_path, sheet_name="SPXsector")
            df6 = pd.read_excel(market_path, sheet_name="SPGSCI")
            df7 = pd.read_excel(market_path, sheet_name="Energy")
            dfs = [df1, df2, df3, df4, df5, df6, df7]
            tseries = reduce(lambda left, right: pd.merge(left, right, on='DATE', how='outer'), dfs)

            if 'DATE' in tseries.columns:
                tseries['DATE'] = pd.to_datetime(tseries['DATE'])
            else:
                st.error("시트에 'DATE' 열이 없습니다.")
                st.stop()
            columns = [col for col in tseries.columns if col != 'DATE']

            df_cyclee = pd.read_excel(cylfile_path, sheet_name='CycleE')
            df_cycleemeta = pd.read_excel(cylfile_path, sheet_name='CycleE_Meta')
            phase_nm = ['Expansion', 'Downturn', 'Repair', 'Recovery']

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                start_date = st.date_input('Start', df_cyclee['DATE'].min())
            with col2:
                end_date = st.date_input('End', df_cyclee['DATE'].max())
            with col3:
                selected_column1 = st.selectbox("비교", columns)

            # 입력받은 날짜 구간으로 데이터 필터링
            fdf = df_cyclee[(df_cyclee['DATE'] >= pd.to_datetime(start_date)) &
                            (df_cyclee['DATE'] <= pd.to_datetime(end_date))]

            s_df = tseries[['DATE', selected_column1]]
            s_df['Month'] = s_df['DATE'].dt.to_period('M')
            df_last = s_df.groupby('Month').apply(lambda x: x.loc[x['DATE'].idxmax()]).reset_index(drop=True)
            df_last['DATE'] = df_last['DATE'].apply(lambda x: x.replace(day=1))
            df_last = df_last[['DATE', selected_column1]]
            fdf = pd.merge(fdf, df_last, on='DATE', how='left')

            col1, col2 = st.columns([3, 1])
            with col1:
                fig = go.Figure()
                for phase in phase_nm:
                    fig.add_trace(go.Bar(
                        x=fdf['DATE'], y=fdf[phase],
                        name=f'{phase}',
                        yaxis='y3',
                        opacity=0.4,
                        marker=dict(line=dict(width=0))
                    ))
                fig.add_trace(go.Scatter(
                    x=fdf['DATE'], y=fdf['Economic_Cycle_Indicator'],
                    mode='lines', name='Economic_Cycle_Indicator',
                    yaxis='y1',
                    line=dict(width=4, color='rgb(245, 130, 32)')
                ))
                fig.add_trace(go.Scatter(
                    x=fdf['DATE'], y=fdf[selected_column1],
                    mode='lines', name=selected_column1,
                    yaxis='y2',
                    line=dict(width=4, color='rgb(13, 45, 79)')
                ))

                fig.update_layout(
                    title='Economic Cycle Indicator and Phases',
                    xaxis=dict(title='DATE'),
                    yaxis=dict(
                        title='Economic Cycle Indicator',
                        side='left',
                        title_standoff=20
                    ),
                    yaxis2=dict(
                        title=selected_column1,
                        overlaying='y',
                        side='left',
                        anchor='x',
                        position=0.15,
                        title_standoff=60
                    ),
                    yaxis3=dict(
                        title='Phases',
                        overlaying='y',  # 좌측 y축과 겹치도록 설정
                        side='right',
                        anchor='x',
                        position=0.85
                    ),
                    barmode='overlay',
                    bargap=0,
                    template='plotly_dark',
                    legend=dict(
                        orientation='h',
                        yanchor='top',
                        y=1.1,
                        xanchor='center',  # 범례의 x축 앵커를 가운데로
                        x=0.5
                    ),
                    height=600
                )

                st.plotly_chart(fig)

        elif selected_sub_menu == "Credit Cycle":

            st.title("Credit Cycle")

            df1 = pd.read_excel(market_path, sheet_name="G10Y")
            df2 = pd.read_excel(market_path, sheet_name="OAS")
            df3 = pd.read_excel(market_path, sheet_name="FX")
            df4 = pd.read_excel(market_path, sheet_name="StockIndex")
            df5 = pd.read_excel(market_path, sheet_name="SPXsector")
            df6 = pd.read_excel(market_path, sheet_name="SPGSCI")
            df7 = pd.read_excel(market_path, sheet_name="Energy")
            dfs = [df1, df2, df3, df4, df5, df6, df7]
            tseries = reduce(lambda left, right: pd.merge(left, right, on='DATE', how='outer'), dfs)

            if 'DATE' in tseries.columns:
                tseries['DATE'] = pd.to_datetime(tseries['DATE'])
            else:
                st.error("시트에 'DATE' 열이 없습니다.")
                st.stop()
            columns = [col for col in tseries.columns if col != 'DATE']

            df_cyclee = pd.read_excel(cylfile_path, sheet_name='CycleC')
            df_cycleemeta = pd.read_excel(cylfile_path, sheet_name='CycleC_Meta')
            phase_nm = ['Expansion', 'Downturn', 'Repair', 'Recovery']

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                start_date = st.date_input('Start', df_cyclee['DATE'].min())
            with col2:
                end_date = st.date_input('End', df_cyclee['DATE'].max())
            with col3:
                selected_column1 = st.selectbox("비교", columns)

            # 입력받은 날짜 구간으로 데이터 필터링
            fdf = df_cyclee[(df_cyclee['DATE'] >= pd.to_datetime(start_date)) &
                            (df_cyclee['DATE'] <= pd.to_datetime(end_date))]

            s_df = tseries[['DATE', selected_column1]]
            s_df['Month'] = s_df['DATE'].dt.to_period('M')
            df_last = s_df.groupby('Month').apply(lambda x: x.loc[x['DATE'].idxmax()]).reset_index(drop=True)
            df_last['DATE'] = df_last['DATE'].apply(lambda x: x.replace(day=1))
            df_last = df_last[['DATE', selected_column1]]
            fdf = pd.merge(fdf, df_last, on='DATE', how='left')

            col1, col2 = st.columns([3, 1])
            with col1:
                fig = go.Figure()
                for phase in phase_nm:
                    fig.add_trace(go.Bar(
                        x=fdf['DATE'], y=fdf[phase],
                        name=f'{phase}',
                        yaxis='y3',
                        opacity=0.4,
                        marker=dict(line=dict(width=0))
                    ))
                fig.add_trace(go.Scatter(
                    x=fdf['DATE'], y=fdf['Credit_Cycle_Indicator'],
                    mode='lines', name='Credit_Cycle_Indicator',
                    yaxis='y1',
                    line=dict(width=4, color='rgb(245, 130, 32)')
                ))
                fig.add_trace(go.Scatter(
                    x=fdf['DATE'], y=fdf[selected_column1],
                    mode='lines', name=selected_column1,
                    yaxis='y2',
                    line=dict(width=4, color='rgb(13, 45, 79)')
                ))

                fig.update_layout(
                    title='Credit Cycle Indicator and Phases',
                    xaxis=dict(title='DATE'),
                    yaxis=dict(
                        title='Credit Cycle Indicator',
                        side='left',
                        title_standoff=20
                    ),
                    yaxis2=dict(
                        title=selected_column1,
                        overlaying='y',
                        side='left',
                        anchor='x',
                        position=0.15,
                        title_standoff=60
                    ),
                    yaxis3=dict(
                        title='Phases',
                        overlaying='y',  # 좌측 y축과 겹치도록 설정
                        side='right',
                        anchor='x',
                        position=0.85
                    ),
                    barmode='overlay',
                    bargap=0,
                    template='plotly_dark',
                    legend=dict(
                        orientation='h',
                        yanchor='top',
                        y=1.1,
                        xanchor='center',  # 범례의 x축 앵커를 가운데로
                        x=0.5
                    ),
                    height=600
                )

                st.plotly_chart(fig)

    elif selected_main_menu == "유사국면":
        if selected_sub_menu == "유사국면분석":
            st.title("금리(US10, US2), 스프레드(USIG, EMIG), 주가지수(SPX, MSCIEM), DXY, Gold, Oil 기준 유사국면")

            df_raw = pd.read_excel(simfile_path, sheet_name='RawdataSim')
            df_simdt = pd.read_excel(simfile_path, sheet_name='siminfo')

            df_raw['DATE'] = pd.to_datetime(df_raw['DATE'])
            df_simdt['SDATE'] = pd.to_datetime(df_simdt['SDATE'])
            df_simdt['EDATE'] = pd.to_datetime(df_simdt['EDATE'])
            df_simdt['SDATE_SIM'] = pd.to_datetime(df_simdt['SDATE_SIM'])
            df_simdt['EDATE_SIM'] = pd.to_datetime(df_simdt['EDATE_SIM'])

            dates_rev = df_simdt['EDATE'].unique()[::-1]
            sel_edt = st.selectbox(
                "분석기준일을 선택하면 해당기준일에 산출한 유사국면 리스트가 생성됩니다.",
                dates_rev,
                index=0
            )

            sel_df = df_simdt[df_simdt['EDATE'] == sel_edt]
            st.table(sel_df)

            fil_dt = df_simdt[df_simdt['EDATE'] == sel_edt]['EDATE_SIM']
            sel_simdt = st.selectbox("산출된 유사국면을 선택(EDATE_SIM 기준)하면 정보가 표시됩니다.", fil_dt)

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                selectafter = st.radio("After:", ["20D", "40D", "60D"], horizontal=True)

            if selectafter == "20D":
                numafter = 20
            elif selectafter == "40D":
                numafter = 40
            else:
                numafter = 60

            sdate = pd.to_datetime(
                df_simdt[(df_simdt['EDATE'] == sel_edt) & (df_simdt['EDATE_SIM'] == sel_simdt)]['SDATE'].values[0])
            edate = pd.to_datetime(
                df_simdt[(df_simdt['EDATE'] == sel_edt) & (df_simdt['EDATE_SIM'] == sel_simdt)]['EDATE'].values[0])
            sdate_sim = pd.to_datetime(
                df_simdt[(df_simdt['EDATE'] == sel_edt) & (df_simdt['EDATE_SIM'] == sel_simdt)]['SDATE_SIM'].values[0])
            edate_sim = pd.to_datetime(
                df_simdt[(df_simdt['EDATE'] == sel_edt) & (df_simdt['EDATE_SIM'] == sel_simdt)]['EDATE_SIM'].values[0])

            lenk = df_simdt[(df_simdt['EDATE'] == sel_edt) & (df_simdt['EDATE_SIM'] == sel_simdt)]['LEN_K'].values[0]
            lens = df_simdt[(df_simdt['EDATE'] == sel_edt) & (df_simdt['EDATE_SIM'] == sel_simdt)]['LEN_S'].values[0]
            numaftsim = np.round((lens / lenk) * numafter).astype(int)
            # lenf = max(lenk, lens)

            edtsim_index = df_raw[df_raw['DATE'] == pd.to_datetime(edate_sim)].index[0]
            looka = df_raw.loc[edtsim_index + numaftsim, 'DATE']
            df2 = df_raw[(df_raw['DATE'] >= sdate_sim) & (df_raw['DATE'] <= looka)]

            edate_index = df_raw[df_raw['DATE'] == pd.to_datetime(edate)].index[0]
            target_index = edate_index + numafter
            if target_index >= len(df_raw):
                extra_rows = target_index - len(df_raw) + 1
                last_date = df_raw['DATE'].iloc[-1]
                new_rows = pd.DataFrame({
                    'DATE': [last_date + pd.Timedelta(days=i + 1) for i in range(extra_rows)]
                })
                for col in df_raw.columns:
                    if col != 'DATE':
                        new_rows[col] = np.nan
                df_rawx = pd.concat([df_raw, new_rows], ignore_index=True)
                lookb = df_rawx.loc[target_index, 'DATE']
                df1 = df_rawx[(df_rawx['DATE'] >= sdate) & (df_rawx['DATE'] <= lookb)]
            else:
                lookb = df_raw.loc[target_index, 'DATE']
                df1 = df_raw[(df_raw['DATE'] >= sdate) & (df_raw['DATE'] <= lookb)]


            def genfig(colnm, dfa=df1, dfb=df2):

                dfa = dfa[[colnm]]
                dfb = dfb[[colnm]]
                if len(dfa) == len(dfb):
                    dfa['dindex'] = range(1, len(dfa) + 1)
                    dfb['dindex'] = range(1, len(dfb) + 1)
                    dfb = dfb.rename(columns={colnm: f"{colnm}_sim"})
                    dfc = pd.merge(dfa, dfb, on='dindex', how='left')
                elif len(dfa) > len(dfb):
                    dfa['dindex'] = range(1, len(dfa) + 1)
                    dfb['dindex'] = np.round(np.linspace(1, len(dfa) + 1, num=len(dfb))).astype(int)
                    dfb = dfb.rename(columns={colnm: f"{colnm}_sim"})
                    dfc = pd.merge(dfa, dfb, on='dindex', how='left')
                else:
                    dfa['dindex'] = range(1, len(dfa) + 1)
                    indices = np.round(np.linspace(1, len(dfb) + 1, num=len(dfa))).astype(int)
                    indices = np.clip(indices, 0, len(dfb) - 1)
                    dfb = dfb.iloc[indices].reset_index(drop=True)  # 인덱스 리셋
                    dfb['dindex'] = range(1, len(dfb) + 1)
                    dfb = dfb.rename(columns={colnm: f"{colnm}_sim"})
                    dfc = pd.merge(dfa, dfb, on='dindex', how='left')

                dfc[f"{colnm}_sim"] = dfc[f"{colnm}_sim"].interpolate()
                dfc_a = dfc.iloc[0:120]
                dfc_b = dfc.iloc[120:]
                dfc_a[colnm] = dfc_a[colnm].interpolate()
                dfc = pd.concat([dfc_a, dfc_b])

                fig = go.Figure()
                fig.add_trace(go.Scatter(x=dfc['dindex'], y=dfc[colnm], mode='lines', line=dict(color='rgb(245, 130, 32)'),
                                         name=f'{colnm} ({sdate.date()} ~ {edate.date()})'))
                fig.add_trace(go.Scatter(x=dfc['dindex'], y=dfc[f'{colnm}_sim'], mode='lines', line=dict(color='rgb(13, 45, 79)'),
                                         name=f'{colnm}_sim ({sdate_sim.date()} ~ {edate_sim.date()})', yaxis='y2'))

                specificx = 120
                fig.add_vline(x=specificx, line_width=1, line_dash="dash", line_color="black")

                fig.update_layout(
                    title=f'Similarity: {colnm}',
                    xaxis=dict(title='Date'),
                    yaxis=dict(title='기준시점'),
                    yaxis2=dict(title='유사국면', overlaying='y', side='right'),
                    template='plotly_dark',
                    legend=dict(
                        orientation='h',
                        yanchor='top',
                        y=1.1,
                        xanchor='center',  # 범례의 x축 앵커를 가운데로
                        x=0.5
                    )
                )

                return fig


            col1, col2, col3, col4 = st.columns(4)
            with col1:
                fig_US10Y = genfig('US10Y')
                fig_DXY = genfig('DXY')
                fig_GOLD = genfig('GOLD')
                st.plotly_chart(fig_US10Y)
                st.plotly_chart(fig_DXY)
                st.plotly_chart(fig_GOLD)
            with col2:
                fig_US2Y = genfig('US2Y')
                fig_SPX = genfig('SPX')
                st.plotly_chart(fig_US2Y)
                st.plotly_chart(fig_SPX)
            with col3:
                fig_USIG = genfig('USIG')
                fig_MSCIEM = genfig('MSCIEM')
                st.plotly_chart(fig_USIG)
                st.plotly_chart(fig_MSCIEM)
            with col4:
                fig_EMSOV = genfig('EMSOV')
                fig_OIL = genfig('OIL')
                st.plotly_chart(fig_EMSOV)
                st.plotly_chart(fig_OIL)

    elif selected_main_menu == "Macro 분석":
        if selected_sub_menu == "Macro Driver":

            st.title("Macro Driver")
            st.write("")

            html = """
                    <style>
                        .custom-text {
                            line-height: 1.2; /* 행간을 줄이는 CSS 속성 */
                        }
                    </style>
                    <div class="custom-text">
                        <p>1. Target에 대한 매크로변수의 연관성 분석(regression)</p>
                        <p>2. Daily/Weekly/Monthly의 78개 매크로 지표로 분석</p>
                        <p>3. StartFrom: Regression을 적합할 분석기간 설정(연도)</p>
                        <p>4. input에 Lag를 줄 것인지(Lag는 Weekly/Monthly 지표에만 적용됨)</p>
                        <p>5. Target 고르기</p>
                        <p>6. 78개의 지표에 대한 회귀분석 후 R-Square의 역순으로 정렬</p>                                    
                    </div>
                    """
            st.markdown(html, unsafe_allow_html=True)

            df1 = pd.read_excel(market_path, sheet_name="G10Y")
            df2 = pd.read_excel(market_path, sheet_name="OAS")
            df3 = pd.read_excel(market_path, sheet_name="FX")
            df4 = pd.read_excel(market_path, sheet_name="StockIndex")
            df5 = pd.read_excel(market_path, sheet_name="SPXsector")
            df6 = pd.read_excel(market_path, sheet_name="SPGSCI")
            df7 = pd.read_excel(market_path, sheet_name="Energy")
            dfs = [df1, df2, df3, df4, df5, df6, df7]
            dfseries = reduce(lambda left, right: pd.merge(left, right, on='DATE', how='outer'), dfs)

            dfmeta = pd.read_excel(macro_path, sheet_name='META')
            dfd = pd.read_excel(macro_path, sheet_name='DAILYV', skiprows=1)
            dfw = pd.read_excel(macro_path, sheet_name='WEEKLYV', skiprows=1)
            dfm = pd.read_excel(macro_path, sheet_name='MONTHLYV', skiprows=1)

            if 'DATE' in dfseries.columns:
                dfseries['DATE'] = pd.to_datetime(dfseries['DATE'])
            else:
                st.error("시트에 'DATE' 열이 없습니다.")
                st.stop()
            columns = [col for col in dfseries.columns if col != 'DATE']

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                selyr = st.number_input('StartFrom', value=2015)
            with col2:
                inputlag = st.number_input('Lag', value=1)
            with col3:
                seltgt = st.selectbox("Target", columns)

            dfseriesx = dfseries[['DATE', seltgt]]

            meta_d = dfmeta[dfmeta['Period'] == 'Daily']
            meta_w = dfmeta[dfmeta['Period'] == 'Weekly']
            meta_m = dfmeta[dfmeta['Period'] == 'Monthly']

            def reggen(df, dfmeta, prd, inputlag=inputlag, dfseriesx=dfseriesx, selyr=selyr, seltgt=seltgt):

                resulta = []
                for i in range(0, len(dfmeta)):
                    num1 = int(dfmeta.iloc[i]['COL1']) - 1
                    num2 = int(dfmeta.iloc[i]['COL2']) - 1
                    df_x = df.iloc[:, num1:num2 + 1].dropna()
                    df_x.rename(columns={df_x.columns[0]: 'DATE'}, inplace=True)
                    if prd == "Monthly":
                        df_x = df_x.groupby(df_x['DATE'].dt.to_period('M')).last().reset_index(drop=True)
                        df_x['DATE'] = df_x['DATE'].dt.to_period('M').dt.to_timestamp(how='end').dt.normalize()
                        df_s = dfseriesx
                        df_s = df_s.groupby(df_s['DATE'].dt.to_period('M')).last().reset_index(drop=True)
                        df_s['DATE'] = df_s['DATE'].dt.to_period('M').dt.to_timestamp(how='end').dt.normalize()
                        df_x = pd.merge(df_x, df_s, on='DATE', how='left')
                    else:
                        df_x = pd.merge(df_x, dfseriesx, on='DATE', how='left')

                    if prd == 'Daily':
                        df_x['X'] = df_x.iloc[:, 1].shift(0)
                    else:
                        df_x['X'] = df_x.iloc[:, 1].shift(inputlag)
                    df_x = df_x[['DATE', 'X', seltgt]]
                    df_x = df_x.dropna()
                    df_x['DATE'] = pd.to_datetime(df_x['DATE'])
                    df_x = df_x[df_x['DATE'].dt.year >= selyr]
                    reg_y = df_x.iloc[:, 2]
                    reg_x = df_x.iloc[:, 1]
                    if not reg_y.empty and not reg_x.empty:
                        reg_x = sm.add_constant(reg_x)
                        model = sm.OLS(reg_y, reg_x).fit()
                        coefficients = model.params
                        r_squared = model.rsquared
                        f_value = model.fvalue
                        f_pvalue = model.f_pvalue
                        t_values = model.tvalues
                        p_values = model.pvalues
                        dw_statistic = durbin_watson(model.resid)

                        resultx = pd.DataFrame({
                            'InputName': dfmeta.iloc[i]['Name'],
                            'Ticker': dfmeta.iloc[i]['BB Ticker'],
                            'Sdate': df_x['DATE'].min(),
                            'Edate': df_x['DATE'].max(),
                            'InputUnit': dfmeta.iloc[i]['Unit'],
                            'InputChg1': dfmeta.iloc[i]['MoM / YoY'],
                            'InputChg2': dfmeta.iloc[i]['Change / %'],
                            'Intercept': coefficients.iloc[0],
                            'Slope': coefficients.iloc[1],
                            'R_squared': r_squared,
                            'probf_reg': f_pvalue,
                            'probt_Intercept': p_values.iloc[0],
                            'probt_Slope': p_values.iloc[1],
                            'DW': dw_statistic
                        }, index=[0])
                        resulta.append(resultx)

                fresult = pd.concat(resulta, ignore_index=True)
                fresult = fresult.sort_values(by='R_squared', ascending=False)

                return fresult

            def reggendm(df, dfmeta, prd, inputlag=inputlag, dfseriesx=dfseriesx, selyr=selyr, seltgt=seltgt):

                resulta = []
                for i in range(0, len(dfmeta)):
                    num1 = int(dfmeta.iloc[i]['COL1']) - 1
                    num2 = int(dfmeta.iloc[i]['COL2']) - 1
                    df_x = df.iloc[:, num1:num2 + 1].dropna()
                    df_x.rename(columns={df_x.columns[0]: 'DATE'}, inplace=True)
                    df_x = pd.merge(df_x, dfseriesx, on='DATE', how='left')
                    df_x['X'] = df_x.iloc[:, 1].shift(0)
                    df_x = df_x[['DATE', 'X', seltgt]]
                    df_x['DATE'] = pd.to_datetime(df_x['DATE'])
                    df_x = df_x.groupby(df_x['DATE'].dt.to_period('M')).last().reset_index(drop=True)
                    df_x = df_x.dropna()
                    df_x['DATE'] = pd.to_datetime(df_x['DATE'])
                    df_x = df_x[df_x['DATE'].dt.year >= selyr]
                    reg_y = df_x.iloc[:, 2]
                    reg_x = df_x.iloc[:, 1]
                    if not reg_y.empty and not reg_x.empty:
                        reg_x = sm.add_constant(reg_x)
                        model = sm.OLS(reg_y, reg_x).fit()
                        coefficients = model.params
                        r_squared = model.rsquared
                        f_value = model.fvalue
                        f_pvalue = model.f_pvalue
                        t_values = model.tvalues
                        p_values = model.pvalues
                        dw_statistic = durbin_watson(model.resid)

                        resultx = pd.DataFrame({
                            'InputName': dfmeta.iloc[i]['Name'],
                            'Ticker': dfmeta.iloc[i]['BB Ticker'],
                            'Sdate': df_x['DATE'].min(),
                            'Edate': df_x['DATE'].max(),
                            'InputUnit': dfmeta.iloc[i]['Unit'],
                            'InputChg1': dfmeta.iloc[i]['MoM / YoY'],
                            'InputChg2': dfmeta.iloc[i]['Change / %'],
                            'Intercept': coefficients.iloc[0],
                            'Slope': coefficients.iloc[1],
                            'R_squared': r_squared,
                            'probf_reg': f_pvalue,
                            'probt_Intercept': p_values.iloc[0],
                            'probt_Slope': p_values.iloc[1],
                            'DW': dw_statistic
                        }, index=[0])
                        resulta.append(resultx)

                fresult = pd.concat(resulta, ignore_index=True)
                fresult = fresult.sort_values(by='R_squared', ascending=False)

                return fresult

            resreg_d = reggen(df=dfd, dfmeta=meta_d, prd='Daily')
            resreg_dm = reggendm(df=dfd, dfmeta=meta_d, prd='Daily')
            resreg_w = reggen(df=dfw, dfmeta=meta_w, prd='Weekly')
            resreg_m = reggen(df=dfm, dfmeta=meta_m, prd='Monthly')

            st.subheader("Daily")
            st.table(resreg_d)
            st.subheader("Daily(Monthly변환 후 Reg)")
            st.table(resreg_dm)
            st.subheader("Weekly")
            st.table(resreg_w)
            st.subheader("Monthly")
            st.table(resreg_m)

        elif selected_sub_menu == "Macro: Actual vs. Survey":

            st.title("Actual vs. Survey")
            st.write("")

            df1 = pd.read_excel(market_path, sheet_name="G10Y")
            df2 = pd.read_excel(market_path, sheet_name="OAS")
            df3 = pd.read_excel(market_path, sheet_name="FX")
            df4 = pd.read_excel(market_path, sheet_name="StockIndex")
            df5 = pd.read_excel(market_path, sheet_name="SPXsector")
            # df6 = pd.read_excel(market_path, sheet_name="SPGSCI")
            # df7 = pd.read_excel(market_path, sheet_name="Energy")
            dfs = [df1, df2, df3, df4, df5]
            dfseries = reduce(lambda left, right: pd.merge(left, right, on='DATE', how='outer'), dfs)
            columns = [col for col in dfseries.columns if col != 'DATE']

            dfmeta = pd.read_excel(macro_path, sheet_name='META')
            dfm = pd.read_excel(macro_path, sheet_name='MONTHLYV', skiprows=1)
            dfms = pd.read_excel(macro_path, sheet_name='MSURVEYV', skiprows=1)

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                selyr = st.number_input('StartFrom', value=2015)
            with col2:
                inputlag = st.number_input('Lag', value=1)
            with col3:
                seltgt = st.selectbox("Target", columns)

            dfseriesx = dfseries[['DATE', seltgt]]

            meta_m = dfmeta[dfmeta['Period'] == 'Monthly']
            meta_m = meta_m[meta_m['Surv'] == 'O']

            selm = meta_m['Name'].tolist()
            col1, col2 = st.columns([3, 1])
            with col1:
                selmacro = st.selectbox('지표를 선택해주세요.:', selm)

            num1 = int(meta_m[meta_m['Name'] == selmacro]['COL1']) - 1
            num2 = int(meta_m[meta_m['Name'] == selmacro]['COL2']) - 1

            dfd_x = dfm.iloc[:, num1:num2 + 1].dropna()
            dfs_x = dfms.iloc[:, num1:num2 + 1].dropna()

            dfd_x.rename(columns={dfd_x.columns[0]: 'DATE', dfd_x.columns[1]: 'Actual'}, inplace=True)
            dfs_x.rename(columns={dfs_x.columns[0]: 'DATE', dfs_x.columns[1]: 'Survey'}, inplace=True)
            dfd_x = dfd_x.groupby(dfd_x['DATE'].dt.to_period('M')).last().reset_index(drop=True)
            dfd_x['DATE'] = dfd_x['DATE'].dt.to_period('M').dt.to_timestamp(how='end').dt.normalize()
            dfs_x = dfs_x.groupby(dfs_x['DATE'].dt.to_period('M')).last().reset_index(drop=True)
            dfs_x['DATE'] = dfs_x['DATE'].dt.to_period('M').dt.to_timestamp(how='end').dt.normalize()
            dfds = pd.merge(dfd_x, dfs_x, on='DATE', how='right')
            cond = [
                (dfds['Actual'] > dfds['Survey']),
                (dfds['Actual'] == dfds['Survey']),
                (dfds['Actual'] < dfds['Survey'])
            ]
            choices = [1, 0, -1]
            dfds['Diff'] = np.select(cond, choices)

            dfds['Prev_Actual'] = dfds['Actual'].shift(1)
            cond2 = [
                (dfds['Actual'] > dfds['Prev_Actual']),
                (dfds['Actual'] == dfds['Prev_Actual']),
                (dfds['Actual'] < dfds['Prev_Actual'])
            ]
            choices2 = [1, 0, -1]
            dfds['Actual_Direc'] = np.select(cond2, choices2, default=np.nan)
            dfds.drop(columns=['Prev_Actual'], inplace=True)

            df_s = dfseriesx
            df_s = df_s.groupby(df_s['DATE'].dt.to_period('M')).last().reset_index(drop=True)
            df_s['DATE'] = df_s['DATE'].dt.to_period('M').dt.to_timestamp(how='end').dt.normalize()
            df_s[f"{seltgt}_chg"] = df_s[seltgt].diff()
            dfds = pd.merge(dfds, df_s, on='DATE', how='outer')
            xlag = inputlag * -1
            dfds['tgt'] = dfds.iloc[:, -2].shift(xlag)
            dfds['tgtchg'] = dfds.iloc[:, -2].shift(xlag) # tgt가 생기니까 -2
            dfds = dfds[dfds['DATE'].dt.year >= selyr]

            last_row = dfds.iloc[[-1]]
            df_dropped = dfds.iloc[:-1].dropna()
            dfdsline = pd.concat([df_dropped, last_row], ignore_index=True)
            dfdsbar = dfds.dropna()

            fig1 = make_subplots(specs=[[{"secondary_y": True}]])
            fig1.add_trace(
                go.Scatter(x=dfdsline['DATE'], y=dfdsline['Actual'], mode='lines', name='Actual',
                           line=dict(color='rgb(245, 130, 32)')),
                secondary_y=False
            )
            fig1.add_trace(
                go.Scatter(x=dfdsline['DATE'], y=dfdsline['Survey'], mode='lines', name='Survey',
                           line=dict(color='rgb(13, 45, 79)')),
                secondary_y=False
            )
            fig1.add_trace(
                go.Scatter(x=dfdsline['DATE'], y=dfdsline['tgt'], mode='lines', name=seltgt,
                           line=dict(dash='solid', color='rgb(0, 169, 206)')),
                secondary_y=True
            )
            fig1.update_layout(
                title_text="Actual vs. Survey",
                xaxis_title="Date",
                yaxis_title="value",
                yaxis2_title=seltgt,
                template='plotly_dark'
            )

            #chart_bar = dfdsbar.tail(36)
            chart_bar = dfdsbar.copy()

            count_hit1 = ((chart_bar['Diff'] * chart_bar['tgtchg']) > 0).sum() / (chart_bar['Diff'] != 0).sum()
            count_hit2 = ((chart_bar['Actual_Direc'] * chart_bar['tgtchg']) > 0).sum() / (chart_bar['Actual_Direc'] != 0).sum()
            hit1 = f"{count_hit1 * 100:.1f}%"
            hit2 = f"{count_hit2 * 100:.1f}%"

            if abs(dfdsbar['tgtchg'].min()) > abs(dfdsbar['tgtchg'].max()):
                y_min = dfdsbar['tgtchg'].min() - dfdsbar['tgtchg'].min() * 0.1
                y_max = abs(dfdsbar['tgtchg'].min()) + dfdsbar['tgtchg'].min() * 0.1
            else:
                y_min = dfdsbar['tgtchg'].max() * -1 - dfdsbar['tgtchg'].max() * 0.1
                y_max = dfdsbar['tgtchg'].max() + dfdsbar['tgtchg'].max() * 0.1

            fig2 = make_subplots(specs=[[{"secondary_y": True}]])
            fig2.add_trace(
                go.Bar(x=chart_bar['DATE'], y=chart_bar['Diff'], name='Actual - Survey',
                       marker=dict(color='rgb(13, 45, 79)', opacity=1, line=dict(width=0))),
                secondary_y=False,
            )
            fig2.add_trace(
                go.Bar(x=chart_bar['DATE'], y=chart_bar['tgtchg'], name=f"chg.{seltgt}",
                       marker=dict(color='rgb(245, 130, 32)', opacity=1, line=dict(width=0))),
                secondary_y=True,
            )
            fig2.update_yaxes(title_text="Actual > Survey: 1, else -1", secondary_y=False, range=[-1, 1])
            fig2.update_yaxes(title_text="chg.tgtchg", secondary_y=True, range=[y_min, y_max])
            fig2.update_layout(
                title_text=f"Actual vs. Survey: Hit({hit1})",
                xaxis_title="Date",
                template='plotly_dark',
                barmode='overlay',
                bargap=0,
                bargroupgap=0,
                legend=dict(
                    orientation='h',
                    yanchor='top',
                    y=1.1,
                    xanchor='center',
                    x=0.5
                )
            )

            fig3 = make_subplots(specs=[[{"secondary_y": True}]])
            fig3.add_trace(
                go.Bar(x=chart_bar['DATE'], y=chart_bar['Actual_Direc'], name='Actual_Direc',
                       marker=dict(color='rgb(13, 45, 79)', opacity=1, line=dict(width=0))),
                secondary_y=False,
            )
            fig3.add_trace(
                go.Bar(x=chart_bar['DATE'], y=chart_bar['tgtchg'], name=f"chg.{seltgt}",
                       marker=dict(color='rgb(245, 130, 32)', opacity=1, line=dict(width=0))),
                secondary_y=True,
            )
            fig3.update_yaxes(title_text="Actual +: 1, else -1", secondary_y=False, range=[-1, 1])
            fig3.update_yaxes(title_text="chg.tgtchg", secondary_y=True, range=[y_min, y_max])
            fig3.update_layout(
                title_text=f"[cf]Actual: Hit({hit2})",
                xaxis_title="Date",
                template='plotly_dark',
                barmode='overlay',
                bargap=0,
                bargroupgap=0,
                legend=dict(
                    orientation='h',
                    yanchor='top',
                    y=1.1,
                    xanchor='center',
                    x=0.5
                )
            )

            col1, col2 = st.columns([3, 1])
            with col1:

                st.plotly_chart(fig1)

                st.write("")
                html = """
                                    <style>
                                        .custom-text {
                                            line-height: 1.2; /* 행간을 줄이는 CSS 속성 */
                                        }
                                    </style>
                                    <div class="custom-text">
                                        <p>남색 막대가 +에 있으면 실제치>예상치, -에 있으면 실제치<예상치</p>
                                        <p>남색 막대 위에 주황색 막대가 있다면, 실제치가 예상치를 상회했을 때, Target 지표가 상승했음을 의미</p>                                    
                                    </div>
                                    """
                st.markdown(html, unsafe_allow_html=True)
                st.plotly_chart(fig2)

                st.write("")
                html = """
                                                    <style>
                                                        .custom-text {
                                                            line-height: 1.2; /* 행간을 줄이는 CSS 속성 */
                                                        }
                                                    </style>
                                                    <div class="custom-text">
                                                        <p>비교를 위해 실제 지표의 등락과 Target 지표의 방향일치성 확인</p>
                                                        <p>남색 막대 위에 주황색 막대가 있다면, 실제치가 상승했을 때, Target 지표가 상승했음을 의미</p>                                    
                                                    </div>
                                                    """
                st.markdown(html, unsafe_allow_html=True)
                st.plotly_chart(fig3)

    elif selected_main_menu == "모델전망 & Signal":
        if selected_sub_menu == "금리":

            dfw = pd.read_excel(model_path, sheet_name='Week')
            dfm = pd.read_excel(model_path, sheet_name='Month')

            st.title("Model Forecast - Duration")
            st.write("")

            html = """
                    <style>
                        .custom-text {
                            line-height: 1.2; /* 행간을 줄이는 CSS 속성 */
                        }
                    </style>
                    <div class="custom-text">
                        <p>1. 좌축: 주황색 영역이 +/- 이면, 금리 하락/상승 시그널이며, 없으면 중립시그널</p>
                        <p>2. 우축: 금요일 기준 시그널 발생 이후 1주간(월~월)의 실제 금리 등락폭</p>
                        <p>3. 주황색 영역과 남색 막대의 부호가 같으면(주황색 영역 위에 남색 막대가 있으면) 방향이 적중했음을 의미</p>
                        <p>4. 우측 테이블의 가장 하단값은 이번 주의 시그널(Actual값 없음)</p>
                    </div>
                    """
            st.markdown(html, unsafe_allow_html=True)

            st.write("")
            st.subheader("Duration Model1(선행변수모델) - Weekly")
            col1, col2 = st.columns([2, 1])
            with col1:
                chart_sig = dfw.tail(52)
                fig1 = make_subplots(specs=[[{"secondary_y": True}]])

                fig1.add_trace(
                    go.Bar(x=chart_sig['DATE'], y=chart_sig['Dur_Leading'], name='Dur_Leading',
                           marker=dict(color='rgb(245, 130, 32)', opacity=1, line=dict(width=0))),
                    secondary_y=False,
                )
                fig1.add_trace(
                    go.Bar(x=chart_sig['DATE'], y=chart_sig['Chg_Dur'], name='Chg_Dur',
                           marker=dict(color='rgb(13, 45, 79)', opacity=1, line=dict(width=0))),
                    secondary_y=True,
                )
                fig1.update_yaxes(range=[-1, 1], secondary_y=False, autorange='reversed', dtick=1)
                fig1.update_yaxes(range=[-0.5, 0.5], secondary_y=True)
                fig1.update_layout(
                    title_text="Duration Model1(선행변수모델) - Weekly",
                    xaxis_title="Date",
                    yaxis_title="Dur_Leading",
                    yaxis2_title="Chg_Dur",
                    template='plotly_dark',
                    barmode='overlay',
                    bargap=0,
                    bargroupgap=0,
                    legend=dict(
                        orientation='h',
                        yanchor='top',
                        y=1.1,
                        xanchor='center',
                        x=0.5
                    )
                )
                st.plotly_chart(fig1, use_container_width=True)
            with col2:
                recent_sig = dfw[['DATE', 'Dur_Leading', 'Act_Direc_Dur', 'Chg_Dur']].tail(10)
                html = recent_sig.to_html(index=False, border=0)
                last_row_style = '<style>table.dataframe tr:last-child { font-weight: bold; }</style>'
                html = last_row_style + html
                st.markdown(html, unsafe_allow_html=True)

            st.subheader("Duration Mdoel2(Boruta) - Weekly")
            col1, col2 = st.columns([2, 1])
            with col1:
                chart_sig = dfw.tail(52)
                fig2 = make_subplots(specs=[[{"secondary_y": True}]])

                fig2.add_trace(
                    go.Bar(x=chart_sig['DATE'], y=chart_sig['Dur_Boruta'], name='Dur_Boruta',
                           marker=dict(color='rgb(245, 130, 32)', opacity=1, line=dict(width=0))),
                    secondary_y=False,
                )
                fig2.add_trace(
                    go.Bar(x=chart_sig['DATE'], y=chart_sig['Chg_Dur'], name='Chg_Dur',
                           marker=dict(color='rgb(13, 45, 79)', opacity=1, line=dict(width=0))),
                    secondary_y=True,
                )
                fig2.update_yaxes(range=[-1, 1], secondary_y=False, autorange='reversed', dtick=1)
                fig2.update_yaxes(range=[-0.5, 0.5], secondary_y=True)
                fig2.update_layout(
                    title_text="Duration Model2(Boruta) - Weekly",
                    xaxis_title="Date",
                    yaxis_title="Dur_Boruta",
                    yaxis2_title="Chg_Dur",
                    template='plotly_dark',
                    barmode='overlay',
                    bargap=0,
                    bargroupgap=0,
                    legend=dict(
                        orientation='h',
                        yanchor='top',
                        y=1.1,
                        xanchor='center',
                        x=0.5
                    )
                )
                st.plotly_chart(fig2, use_container_width=True)
            with col2:
                recent_sig = dfw[['DATE', 'Dur_Boruta', 'Act_Direc_Dur', 'Chg_Dur']].tail(10)
                html = recent_sig.to_html(index=False, border=0)
                last_row_style = '<style>table.dataframe tr:last-child { font-weight: bold; }</style>'
                html = last_row_style + html
                st.markdown(html, unsafe_allow_html=True)

            st.subheader("Duration Mdoel3(Tree) - Monthly")
            col1, col2 = st.columns([2, 1])
            with col1:
                chart_sig = dfm.tail(36)
                fig3 = make_subplots(specs=[[{"secondary_y": True}]])

                fig3.add_trace(
                    go.Bar(x=chart_sig['DATE'], y=chart_sig['Duration_Monthly'], name='Dur_Tree',
                           marker=dict(color='rgb(245, 130, 32)', opacity=1, line=dict(width=0))),
                    secondary_y=False,
                )
                fig3.add_trace(
                    go.Bar(x=chart_sig['DATE'], y=chart_sig['Chg_Dur'], name='Chg_Dur',
                           marker=dict(color='rgb(13, 45, 79)', opacity=1, line=dict(width=0))),
                    secondary_y=True,
                )
                fig3.update_yaxes(range=[-1, 1], secondary_y=False, autorange='reversed', dtick=1)
                fig3.update_yaxes(range=[-1, 1], secondary_y=True)
                fig3.update_layout(
                    title_text="Duration Model3(Tree) - Monthly",
                    xaxis_title="Date",
                    yaxis_title="Dur_Tree",
                    yaxis2_title="Chg_Dur",
                    template='plotly_dark',
                    barmode='overlay',
                    bargap=0,
                    bargroupgap=0,
                    legend=dict(
                        orientation='h',
                        yanchor='top',
                        y=1.1,
                        xanchor='center',
                        x=0.5
                    )
                )
                st.plotly_chart(fig3, use_container_width=True)
            with col2:
                recent_sig = dfm[['DATE', 'Duration_Monthly', 'Act_Direc_Dur', 'Chg_Dur']].tail(10)
                html = recent_sig.to_html(index=False, border=0)
                last_row_style = '<style>table.dataframe tr:last-child { font-weight: bold; }</style>'
                html = last_row_style + html
                st.markdown(html, unsafe_allow_html=True)

        elif selected_sub_menu == "USIG스프레드":
            dfw = pd.read_excel(model_path, sheet_name='Week')
            dfm = pd.read_excel(model_path, sheet_name='Month')

            st.title("Model Forecast - Credit")
            st.write("")

            html = """
                    <style>
                        .custom-text {
                            line-height: 1.2;
                        }
                    </style>
                    <div class="custom-text">
                        <p>1. 좌축: 주황색 영역이 +/- 이면, 스프레드 축소/확대 시그널이며, 없으면 중립시그널</p>
                        <p>2. 우축: 금요일 기준 시그널 발생 이후 1주간(월~월)의 실제 스프레드 등락폭</p>
                        <p>3. 주황색 영역과 남색 막대의 부호가 같으면(주황색 영역 위에 남색 막대가 있으면) 방향이 적중했음을 의미</p>
                        <p>4. 우측 테이블의 가장 하단값은 이번 주의 시그널(Actual값 없음)</p>
                    </div>
                    """
            st.markdown(html, unsafe_allow_html=True)

            st.write("")
            st.subheader("Credit Model1 - Weekly")
            col1, col2 = st.columns([2, 1])
            with col1:
                chart_sig = dfw.tail(52)
                fig1 = make_subplots(specs=[[{"secondary_y": True}]])

                fig1.add_trace(
                    go.Bar(x=chart_sig['DATE'], y=chart_sig['Credit_1M'], name='Credit_1M',
                           marker=dict(color='rgb(245, 130, 32)', opacity=1, line=dict(width=0))),
                    secondary_y=False,
                )
                fig1.add_trace(
                    go.Bar(x=chart_sig['DATE'], y=chart_sig['Chg_Credit'], name='Chg_Credit',
                           marker=dict(color='rgb(13, 45, 79)', opacity=1, line=dict(width=0))),
                    secondary_y=True,
                )
                fig1.update_yaxes(range=[-1, 1], secondary_y=False, autorange='reversed', dtick=1)
                fig1.update_yaxes(range=[-0.2, 0.2], secondary_y=True)
                fig1.update_layout(
                    title_text="Credit Model1 - Weekly",
                    xaxis_title="Date",
                    yaxis_title="Credit_1M",
                    yaxis2_title="Chg_Credit",
                    template='plotly_dark',
                    barmode='overlay',
                    bargap=0,
                    bargroupgap=0,
                    legend=dict(
                        orientation='h',
                        yanchor='top',
                        y=1.1,
                        xanchor='center',
                        x=0.5
                    )
                )
                st.plotly_chart(fig1, use_container_width=True)
            with col2:
                recent_sig = dfw[['DATE', 'Credit_1M', 'Act_Direc_Credit', 'Chg_Credit']].tail(10)
                html = recent_sig.to_html(index=False, border=0)
                last_row_style = '<style>table.dataframe tr:last-child { font-weight: bold; }</style>'
                html = last_row_style + html
                st.markdown(html, unsafe_allow_html=True)

            st.subheader("Credit Model2 - Weekly")
            col1, col2 = st.columns([2, 1])
            with col1:
                chart_sig = dfw.tail(52)
                fig2 = make_subplots(specs=[[{"secondary_y": True}]])

                fig2.add_trace(
                    go.Bar(x=chart_sig['DATE'], y=chart_sig['Credit_3M'], name='Credit_3M',
                           marker=dict(color='rgb(245, 130, 32)', opacity=1, line=dict(width=0))),
                    secondary_y=False,
                )
                fig2.add_trace(
                    go.Bar(x=chart_sig['DATE'], y=chart_sig['Chg_Credit'], name='Chg_Credit',
                           marker=dict(color='rgb(13, 45, 79)', opacity=1, line=dict(width=0))),
                    secondary_y=True,
                )
                fig2.update_yaxes(range=[-1, 1], secondary_y=False, autorange='reversed', dtick=1)
                fig2.update_yaxes(range=[-0.2, 0.2], secondary_y=True)
                fig2.update_layout(
                    title_text="Credit Model2 - Weekly",
                    xaxis_title="Date",
                    yaxis_title="Credit_3M",
                    yaxis2_title="Chg_Credit",
                    template='plotly_dark',
                    barmode='overlay',
                    bargap=0,
                    bargroupgap=0,
                    legend=dict(
                        orientation='h',
                        yanchor='top',
                        y=1.1,
                        xanchor='center',
                        x=0.5
                    )
                )
                st.plotly_chart(fig2, use_container_width=True)
            with col2:
                recent_sig = dfw[['DATE', 'Credit_1M', 'Act_Direc_Credit', 'Chg_Credit']].tail(10)
                html = recent_sig.to_html(index=False, border=0)
                last_row_style = '<style>table.dataframe tr:last-child { font-weight: bold; }</style>'
                html = last_row_style + html
                st.markdown(html, unsafe_allow_html=True)

            st.subheader("Credit Model3 - Monthly")
            col1, col2 = st.columns([2, 1])
            with col1:
                chart_sig = dfm.tail(36)
                fig3 = make_subplots(specs=[[{"secondary_y": True}]])

                fig3.add_trace(
                    go.Bar(x=chart_sig['DATE'], y=chart_sig['Credit_Monthly'], name='Credit_Monthly',
                           marker=dict(color='rgb(245, 130, 32)', opacity=1, line=dict(width=0))),
                    secondary_y=False,
                )
                fig3.add_trace(
                    go.Bar(x=chart_sig['DATE'], y=chart_sig['Chg_Credit'], name='Chg_Credit',
                           marker=dict(color='rgb(13, 45, 79)', opacity=1, line=dict(width=0))),
                    secondary_y=True,
                )
                fig3.update_yaxes(range=[-1, 1], secondary_y=False, autorange='reversed', dtick=1)
                fig3.update_yaxes(range=[-0.5, 0.5], secondary_y=True)
                fig3.update_layout(
                    title_text="Credit Model3 - Monthly",
                    xaxis_title="Date",
                    yaxis_title="Credit_Monthly",
                    yaxis2_title="Chg_Credit",
                    template='plotly_dark',
                    barmode='overlay',
                    bargap=0,
                    bargroupgap=0,
                    legend=dict(
                        orientation='h',
                        yanchor='top',
                        y=1.1,
                        xanchor='center',
                        x=0.5
                    )
                )
                st.plotly_chart(fig3, use_container_width=True)
            with col2:
                recent_sig = dfm[['DATE', 'Credit_Monthly', 'Act_Direc_Credit', 'Chg_Credit']].tail(10)
                html = recent_sig.to_html(index=False, border=0)
                last_row_style = '<style>table.dataframe tr:last-child { font-weight: bold; }</style>'
                html = last_row_style + html
                st.markdown(html, unsafe_allow_html=True)

        elif selected_sub_menu == "USIG 추천종목":
            st.title("USIG 종목 Picking")

            df = pd.read_excel(usig_path, sheet_name='pick')

            dates_rev = df['DATE'].unique()[::-1]
            sel_dt = st.selectbox(
                "종목산출 기준일을 선택해 주세요(예. 7월말 산출 종목은 8월 포트폴리오를 위한 종목).",
                dates_rev,
                index=0
            )

            model_rev = list(df['Model'].unique())[::-1]
            sector_ = list(df['Sector'].unique())
            tenor_ = list(df['Tenor'].unique())
            rating_ = list(df['Rating'].unique())

            col1, col2, col3, col4 = st.columns(4)

            with col1:
                sel_model = st.selectbox("산출모델 선택", ['All'] + model_rev)
            with col2:
                sel_sector = st.selectbox("섹터", ['All'] + sector_)
            with col3:
                sel_tenor = st.selectbox("만기", ['All'] + tenor_)
            with col4:
                sel_rating = st.selectbox("등급", ['All'] + rating_)

            fdf1 = df[(df['DATE'] == pd.to_datetime(sel_dt))]

            fdf2 = fdf1
            if sel_model != 'All':
                fdf2 = fdf2[fdf2['Model'] == sel_model]
            if sel_sector != 'All':
                fdf2 = fdf2[fdf2['Sector'] == sel_sector]
            if sel_tenor != 'All':
                fdf2 = fdf2[fdf2['Tenor'] == sel_tenor]
            if sel_rating != 'All':
                fdf2 = fdf2[fdf2['Rating'] == sel_rating]

            selable_columns = [col for col in fdf2.columns if col != 'DATE']

            sel_columns = st.multiselect(
                "Column Select:",
                selable_columns,
                default=selable_columns
            )
            if sel_columns:
                st.dataframe(fdf2[sel_columns], hide_index=True)
                csvf = fdf2[sel_columns].to_csv(index=False)
                st.download_button(
                    label="Download CSV",
                    data=csvf,
                    file_name='usig_pick.csv',
                    mime='text/csv'
                )

            st.image(igimage_path, use_column_width=True, output_format='PNG')

        elif selected_sub_menu == "FX":
            st.title("FX Strategy by Transformer")
            st.write("!! 2024.08.26일. USDKRW의 USD 강세모델에서 USD 강세 시그널 발생 !!")

            def fxgenfig1(fxnm, selprob, chart_title, df_path=fx_path):
                df = pd.read_excel(df_path, sheet_name=fxnm)
                if fxnm == "KRWUSD":
                    fxnm = "USDKRW"
                    colnm = ['DATE', 'Prob0', 'Prob1', fxnm, 'fx_v', 'Conviction', 'FX_Long', 'Strategy']
                    df.columns = colnm
                df = df[df['DATE'].notna()]
                fdf = df[df['DATE'] >= pd.Timestamp('2021-01-04')]

                fdf.set_index('DATE', inplace=True)
                all_dates = pd.date_range(start=fdf.index.min(), end=fdf.index.max(), freq='D')
                fdf = fdf.reindex(all_dates, method='pad')

                fig1 = make_subplots(specs=[[{"secondary_y": True}]])
                fig1.add_trace(go.Bar(
                    x=fdf.index, y=fdf[selprob],
                    opacity=1,
                    showlegend=False,
                    marker=dict(
                        color='rgb(245, 130, 32)',
                        line=dict(width=0)
                    )
                ), secondary_y=False)
                fig1.add_trace(go.Bar(
                    x=fdf.index, y=fdf['Conviction'],
                    name='Conviction',
                    opacity=1,
                    marker=dict(
                        color='rgb(255, 217, 102)',
                        line=dict(width=0)
                    )
                ), secondary_y=False)
                fig1.add_trace(go.Scatter(
                    x=fdf.index, y=fdf[fxnm],
                    mode='lines', name=f'{fxnm}',
                    line=dict(width=4, color='rgb(13, 45, 79)'),
                    stackgroup=None
                ), secondary_y=True)

                fig1.update_layout(
                    barmode='overlay',
                    bargap=0.15,
                    title=chart_title,
                    xaxis_title='Date',
                    yaxis_title='Prob/Conviction',
                    yaxis2_title=f'{fxnm}',
                    legend=dict(
                        orientation='h',
                        yanchor='top',
                        y=1.1,
                        xanchor='center',
                        x=0.5
                    )
                )

                fig1.update_yaxes(range=[0, max(fdf['Prob1'].max(), fdf['Conviction'].max()) * 1.1], secondary_y=False)
                fig1.update_yaxes(range=[fdf[fxnm].min() * 0.9, fdf[fxnm].max() * 1.1], secondary_y=True)
                fig1.update_traces(opacity=1, selector=dict(type='bar'))

                fig2 = go.Figure()
                fig2.add_trace(go.Scatter(
                    x=fdf.index, y=fdf['FX_Long'],
                    mode='lines', name=f'{fxnm}',
                    yaxis='y1',
                    line=dict(width=4, color='rgb(245, 130, 32)')
                ))
                fig2.add_trace(go.Scatter(
                    x=fdf.index, y=fdf['Strategy'],
                    mode='lines', name='Strategy',
                    yaxis='y1',
                    line=dict(width=4, color='rgb(13, 45, 79)')
                ))
                fig2.update_layout(
                    xaxis_title='DATE',
                    yaxis_title='Return',
                    template='plotly_dark',
                    legend=dict(
                        orientation='h',
                        yanchor='top',
                        y=1.1,
                        xanchor='center',
                        x=0.5
                    )
                )
                return fig1, fig2


            def fxgenfig2(fxnm, chart_title, df_path=fx_path):
                df = pd.read_excel(df_path, sheet_name=fxnm)
                df = df[df['DATE'].notna()]
                fdf = df[df['DATE'] >= pd.Timestamp('2021-01-04')]

                fdf.set_index('DATE', inplace=True)
                all_dates = pd.date_range(start=fdf.index.min(), end=fdf.index.max(), freq='D')
                fdf = fdf.reindex(all_dates, method='pad')

                fig1 = make_subplots(specs=[[{"secondary_y": True}]])
                fig1.add_trace(go.Bar(
                    x=fdf.index, y=fdf['Prob1'],
                    opacity=1,
                    showlegend=False,
                    marker=dict(
                        color='rgb(245, 130, 32)',
                        line=dict(width=0)
                    )
                ), secondary_y=False)
                fig1.add_trace(go.Bar(
                    x=fdf.index, y=fdf['Conviction'],
                    name='Conviction',
                    opacity=1,
                    marker=dict(
                        color='rgb(255, 217, 102)',
                        line=dict(width=0)
                    )
                ), secondary_y=False)
                fig1.add_trace(go.Scatter(
                    x=fdf.index, y=fdf[fxnm],
                    mode='lines', name=f'{fxnm}',
                    line=dict(width=4, color='rgb(13, 45, 79)'),
                    stackgroup=None
                ), secondary_y=True)

                fig1.update_layout(
                    barmode='overlay',
                    bargap=0.15,
                    title=chart_title,
                    xaxis_title='Date',
                    yaxis_title='Prob/Conviction',
                    yaxis2_title=f'{fxnm}',
                    legend=dict(
                        orientation='h',
                        yanchor='top',
                        y=1.1,
                        xanchor='center',
                        x=0.5
                    )
                )

                fig1.update_yaxes(range=[0, max(fdf['Prob1'].max(), fdf['Conviction'].max()) * 1.1], secondary_y=False)
                fig1.update_yaxes(range=[fdf[fxnm].min() * 0.9, fdf[fxnm].max() * 1.1], secondary_y=True)
                fig1.update_traces(opacity=1, selector=dict(type='bar'))

                return fig1


            fig_USDKRW1, fig_USDKRW2 = fxgenfig1('USDKRW', 'Prob1', 'USDKRW: USD 강세 모델')
            fig_KRWUSD1, fig_KRWUSD2 = fxgenfig1('KRWUSD', 'Prob0', 'USDKRW: KRW 강세 모델')
            fig_USDEUR1 = fxgenfig2('USDEUR', 'USDEUR')
            fig_USDGBP1 = fxgenfig2('USDGBP', 'USDGBP')
            fig_USDCNY1 = fxgenfig2('USDCNY', 'USDCNY')
            fig_USDJPY1 = fxgenfig2('USDJPY', 'USDJPY')

            col1, col2, col3 = st.columns([2, 2, 1])
            with col1:
                st.plotly_chart(fig_USDKRW1)
                st.plotly_chart(fig_KRWUSD1)
                st.write("")
                st.write("")
                st.plotly_chart(fig_USDEUR1)
                st.plotly_chart(fig_USDCNY1)
            with col2:
                st.plotly_chart(fig_USDKRW2)
                st.plotly_chart(fig_KRWUSD2)
                st.write("")
                st.write("")
                st.plotly_chart(fig_USDGBP1)
                st.plotly_chart(fig_USDJPY1)

        elif selected_sub_menu == "RankingModel":
            st.title("Ranking Model Output")
            st.write("")

            html = """
                    <style>
                        .custom-text {
                            line-height: 1.2;
                        }
                    </style>
                    <div class="custom-text">
                        <p>막대그래프는 과거 3개월의 우선순위와 실제성과</p>
                        <p>따라서, 막대그래프가 우하향하는 경향이 강할수록 과거의 순위 예측이 적중했음을 의미</p>                        
                    </div>
                    """
            st.markdown(html, unsafe_allow_html=True)

            st.write("")
            st.subheader("I. Golbal Agg: Region")
            st.write("")

            df = pd.read_excel(allo_path, sheet_name='GAgg1')
            dates = df['DATE'].unique()[::-1]

            sel_latest = df[df['DATE'] == dates[0]]
            sel_latest['DATE'] = pd.to_datetime(sel_latest['DATE'])
            date_maxsel = sel_latest['DATE'].iloc[0]
            formatted_date = date_maxsel.strftime('%Y.%m')

            labelsl = [sel_latest.iloc[0]['Label01'], sel_latest.iloc[0]['Label02'], sel_latest.iloc[0]['Label03'],
                       sel_latest.iloc[0]['Label04'], sel_latest.iloc[0]['Label05'], sel_latest.iloc[0]['Label06']]
            probl = [sel_latest.iloc[0]['Fret01'], sel_latest.iloc[0]['Fret02'], sel_latest.iloc[0]['Fret03'],
                     sel_latest.iloc[0]['Fret04'], sel_latest.iloc[0]['Fret05'], sel_latest.iloc[0]['Fret06']]
            labels_df = pd.DataFrame({
                'Rank': labelsl, 'Score': probl
            })
            labels_df = labels_df.style.format({
                'Score': lambda x: f"{x:.2f}"
            })

            dates = dates[1:4]


            def getbardt(df, dtrow):
                sel_df = df[df['DATE'] == dates[dtrow]]
                labels = [sel_df.iloc[0]['Label01'], sel_df.iloc[0]['Label02'], sel_df.iloc[0]['Label03'],
                          sel_df.iloc[0]['Label04'], sel_df.iloc[0]['Label05'], sel_df.iloc[0]['Label06']]
                values = [sel_df.iloc[0]['Fret01'], sel_df.iloc[0]['Fret02'], sel_df.iloc[0]['Fret03'],
                          sel_df.iloc[0]['Fret04'], sel_df.iloc[0]['Fret05'], sel_df.iloc[0]['Fret06']]
                fig = go.Figure(data=[
                    go.Bar(
                        x=labels,
                        y=values,
                        marker_color=['rgb(245, 130, 32)', 'rgb(245, 130, 32)', 'rgb(245, 130, 32)',
                                      'rgb(245, 130, 32)', 'rgb(245, 130, 32)', 'rgb(245, 130, 32)']
                    )
                ])
                fig.update_layout(
                    title=f"{dates[dtrow]}",
                    xaxis_title="추천순위",
                    yaxis_title="실제성과",
                )
                fig.update_yaxes(range=[min(values) - (0.1 * min(values)), max(values) + (0.1 * max(values))])
                return fig


            fig1 = getbardt(df, 0)
            fig2 = getbardt(df, 1)
            fig3 = getbardt(df, 2)

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.write("")
                st.subheader(f"{formatted_date}월의 순위 예측")
                st.write("")
                st.dataframe(labels_df, use_container_width=True, hide_index=True)
            with col2:
                st.plotly_chart(fig1)
            with col3:
                st.plotly_chart(fig2)
            with col4:
                st.plotly_chart(fig3)

            st.write("")
            st.subheader("II. Golbal Agg: US Sector")
            st.write("")

            df = pd.read_excel(allo_path, sheet_name='GAgg2')
            dates = df['DATE'].unique()[::-1]

            sel_latest = df[df['DATE'] == dates[0]]
            sel_latest['DATE'] = pd.to_datetime(sel_latest['DATE'])
            date_maxsel = sel_latest['DATE'].iloc[0]
            formatted_date = date_maxsel.strftime('%Y.%m')

            labelsl = [sel_latest.iloc[0]['Label01'], sel_latest.iloc[0]['Label02'], sel_latest.iloc[0]['Label03']]
            probl = [sel_latest.iloc[0]['Fret01'], sel_latest.iloc[0]['Fret02'], sel_latest.iloc[0]['Fret03']]
            labels_df = pd.DataFrame({
                'Rank': labelsl, 'Score': probl
            })
            labels_df = labels_df.style.format({
                'Score': lambda x: f"{x:.2f}"
            })

            dates = dates[1:4]


            def getbardt(df, dtrow):
                sel_df = df[df['DATE'] == dates[dtrow]]
                labels = [sel_df.iloc[0]['Label01'], sel_df.iloc[0]['Label02'], sel_df.iloc[0]['Label03']]
                values = [sel_df.iloc[0]['Fret01'], sel_df.iloc[0]['Fret02'], sel_df.iloc[0]['Fret03']]
                fig = go.Figure(data=[
                    go.Bar(
                        x=labels,
                        y=values,
                        marker_color=['rgb(245, 130, 32)', 'rgb(245, 130, 32)', 'rgb(245, 130, 32)']
                    )
                ])
                fig.update_layout(
                    title=f"{dates[dtrow]}",
                    xaxis_title="추천순위",
                    yaxis_title="실제성과",
                )
                fig.update_yaxes(range=[min(values) - (0.1 * min(values)), max(values) + (0.1 * max(values))])
                return fig


            fig1 = getbardt(df, 0)
            fig2 = getbardt(df, 1)
            fig3 = getbardt(df, 2)

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.write("")
                st.subheader(f"{formatted_date}월의 순위 예측")
                st.write("")
                st.dataframe(labels_df, use_container_width=True, hide_index=True)
            with col2:
                st.plotly_chart(fig1)
            with col3:
                st.plotly_chart(fig2)
            with col4:
                st.plotly_chart(fig3)

            st.write("")
            st.subheader("III. Golbal Agg: US Corporate")
            st.write("")

            df = pd.read_excel(allo_path, sheet_name='GAgg3')
            dates = df['DATE'].unique()[::-1]

            sel_latest = df[df['DATE'] == dates[0]]
            sel_latest['DATE'] = pd.to_datetime(sel_latest['DATE'])
            date_maxsel = sel_latest['DATE'].iloc[0]
            formatted_date = date_maxsel.strftime('%Y.%m')

            labelsl = [sel_latest.iloc[0]['Label01'], sel_latest.iloc[0]['Label02'], sel_latest.iloc[0]['Label03'],
                       sel_latest.iloc[0]['Label04'], sel_latest.iloc[0]['Label05'], sel_latest.iloc[0]['Label06']]
            probl = [sel_latest.iloc[0]['Fret01'], sel_latest.iloc[0]['Fret02'], sel_latest.iloc[0]['Fret03'],
                     sel_latest.iloc[0]['Fret04'], sel_latest.iloc[0]['Fret05'], sel_latest.iloc[0]['Fret06']]
            labels_df = pd.DataFrame({
                'Rank': labelsl, 'Score': probl
            })
            labels_df = labels_df.style.format({
                'Score': lambda x: f"{x:.2f}"
            })

            dates = dates[1:4]


            def getbardt(df, dtrow):
                sel_df = df[df['DATE'] == dates[dtrow]]
                labels = [sel_df.iloc[0]['Label01'], sel_df.iloc[0]['Label02'], sel_df.iloc[0]['Label03'],
                          sel_df.iloc[0]['Label04'], sel_df.iloc[0]['Label05'], sel_df.iloc[0]['Label06']]
                values = [sel_df.iloc[0]['Fret01'], sel_df.iloc[0]['Fret02'], sel_df.iloc[0]['Fret03'],
                          sel_df.iloc[0]['Fret04'], sel_df.iloc[0]['Fret05'], sel_df.iloc[0]['Fret06']]
                fig = go.Figure(data=[
                    go.Bar(
                        x=labels,
                        y=values,
                        marker_color=['rgb(245, 130, 32)', 'rgb(245, 130, 32)', 'rgb(245, 130, 32)',
                                      'rgb(245, 130, 32)', 'rgb(245, 130, 32)', 'rgb(245, 130, 32)']
                    )
                ])
                fig.update_layout(
                    title=f"{dates[dtrow]}",
                    xaxis_title="추천순위",
                    yaxis_title="실제성과",
                )
                fig.update_yaxes(range=[min(values) - (0.1 * min(values)), max(values) + (0.1 * max(values))])
                return fig


            fig1 = getbardt(df, 0)
            fig2 = getbardt(df, 1)
            fig3 = getbardt(df, 2)

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.write("")
                st.subheader(f"{formatted_date}월의 순위 예측")
                st.write("")
                st.dataframe(labels_df, use_container_width=True, hide_index=True)
            with col2:
                st.plotly_chart(fig1)
            with col3:
                st.plotly_chart(fig2)
            with col4:
                st.plotly_chart(fig3)

            st.write("")
            st.subheader("IV. Golbal Agg: US Treasury")
            st.write("")

            df = pd.read_excel(allo_path, sheet_name='GAgg4')
            dates = df['DATE'].unique()[::-1]

            sel_latest = df[df['DATE'] == dates[0]]
            sel_latest['DATE'] = pd.to_datetime(sel_latest['DATE'])
            date_maxsel = sel_latest['DATE'].iloc[0]
            formatted_date = date_maxsel.strftime('%Y.%m')

            labelsl = [sel_latest.iloc[0]['Label01'], sel_latest.iloc[0]['Label02'], sel_latest.iloc[0]['Label03'],
                       sel_latest.iloc[0]['Label04'], sel_latest.iloc[0]['Label05'], sel_latest.iloc[0]['Label06'],
                       sel_latest.iloc[0]['Label07']]
            probl = [sel_latest.iloc[0]['Fret01'], sel_latest.iloc[0]['Fret02'], sel_latest.iloc[0]['Fret03'],
                     sel_latest.iloc[0]['Fret04'], sel_latest.iloc[0]['Fret05'], sel_latest.iloc[0]['Fret06'],
                     sel_latest.iloc[0]['Fret07']]
            labels_df = pd.DataFrame({
                'Rank': labelsl, 'Score': probl
            })
            labels_df = labels_df.style.format({
                'Score': lambda x: f"{x:.2f}"
            })

            dates = dates[1:4]


            def getbardt(df, dtrow):
                sel_df = df[df['DATE'] == dates[dtrow]]
                labels = [sel_df.iloc[0]['Label01'], sel_df.iloc[0]['Label02'], sel_df.iloc[0]['Label03'],
                          sel_df.iloc[0]['Label04'], sel_df.iloc[0]['Label05'], sel_df.iloc[0]['Label06'],
                          sel_df.iloc[0]['Label07']]
                values = [sel_df.iloc[0]['Fret01'], sel_df.iloc[0]['Fret02'], sel_df.iloc[0]['Fret03'],
                          sel_df.iloc[0]['Fret04'], sel_df.iloc[0]['Fret05'], sel_df.iloc[0]['Fret06'],
                          sel_df.iloc[0]['Fret07']]
                fig = go.Figure(data=[
                    go.Bar(
                        x=labels,
                        y=values,
                        marker_color=['rgb(245, 130, 32)', 'rgb(245, 130, 32)', 'rgb(245, 130, 32)',
                                      'rgb(245, 130, 32)', 'rgb(245, 130, 32)', 'rgb(245, 130, 32)',
                                      'rgb(245, 130, 32)']
                    )
                ])
                fig.update_layout(
                    title=f"{dates[dtrow]}",
                    xaxis_title="추천순위",
                    yaxis_title="실제성과",
                )
                fig.update_yaxes(range=[min(values) - (0.1 * min(values)), max(values) + (0.1 * max(values))])
                return fig


            fig1 = getbardt(df, 0)
            fig2 = getbardt(df, 1)
            fig3 = getbardt(df, 2)

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.write("")
                st.subheader(f"{formatted_date}월의 순위 예측")
                st.write("")
                st.dataframe(labels_df, use_container_width=True, hide_index=True)
            with col2:
                st.plotly_chart(fig1)
            with col3:
                st.plotly_chart(fig2)
            with col4:
                st.plotly_chart(fig3)

            st.write("")
            st.subheader("V. USIG Sector")
            st.write("")

            df = pd.read_excel(allo_path, sheet_name='USIGSector')
            dates = df['DATE'].unique()[::-1]

            sel_latest = df[df['DATE'] == dates[0]]
            sel_latest['DATE'] = pd.to_datetime(sel_latest['DATE'])
            date_maxsel = sel_latest['DATE'].iloc[0]
            formatted_date = date_maxsel.strftime('%Y.%m')

            labelsl = [sel_latest.iloc[0]['Label01'], sel_latest.iloc[0]['Label02'], sel_latest.iloc[0]['Label03'],
                       sel_latest.iloc[0]['Label04'], sel_latest.iloc[0]['Label05'], sel_latest.iloc[0]['Label06'],
                       sel_latest.iloc[0]['Label07']]
            probl = [sel_latest.iloc[0]['Fret01'], sel_latest.iloc[0]['Fret02'], sel_latest.iloc[0]['Fret03'],
                     sel_latest.iloc[0]['Fret04'], sel_latest.iloc[0]['Fret05'], sel_latest.iloc[0]['Fret06'],
                     sel_latest.iloc[0]['Fret07']]
            labels_df = pd.DataFrame({
                'Rank': labelsl, 'Prob(>BM)': probl
            })
            labels_df = labels_df.style.format({
                'Prob(>BM)': lambda x: f"{x * 100:.1f}%"
            })

            dates = dates[1:4]


            def getbardt(df, dtrow):
                sel_df = df[df['DATE'] == dates[dtrow]]
                labels = [sel_df.iloc[0]['Label01'], sel_df.iloc[0]['Label02'], sel_df.iloc[0]['Label03'],
                          sel_df.iloc[0]['Label04'], sel_df.iloc[0]['Label05'], sel_df.iloc[0]['Label06'],
                          sel_df.iloc[0]['Label07'], sel_df.iloc[0]['Label00']]
                values = [sel_df.iloc[0]['Fret01'], sel_df.iloc[0]['Fret02'], sel_df.iloc[0]['Fret03'],
                          sel_df.iloc[0]['Fret04'], sel_df.iloc[0]['Fret05'], sel_df.iloc[0]['Fret06'],
                          sel_df.iloc[0]['Fret07'], sel_df.iloc[0]['Fret00']]
                fig = go.Figure(data=[
                    go.Bar(
                        x=labels,
                        y=values,
                        marker_color=['rgb(245, 130, 32)', 'rgb(245, 130, 32)', 'rgb(245, 130, 32)',
                                      'rgb(245, 130, 32)', 'rgb(245, 130, 32)', 'rgb(245, 130, 32)',
                                      'rgb(245, 130, 32)', 'rgb(13, 45, 79)']
                    )
                ])
                fig.update_layout(
                    title=f"{dates[dtrow]}",
                    xaxis_title="추천순위",
                    yaxis_title="실제성과",
                )
                fig.update_yaxes(range=[min(values) - (0.1 * min(values)), max(values) + (0.1 * max(values))])
                return fig


            fig1 = getbardt(df, 0)
            fig2 = getbardt(df, 1)
            fig3 = getbardt(df, 2)

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.write("")
                st.subheader(f"{formatted_date}월의 순위 예측")
                st.write("")
                st.dataframe(labels_df, use_container_width=True, hide_index=True)
            with col2:
                st.plotly_chart(fig1)
            with col3:
                st.plotly_chart(fig2)
            with col4:
                st.plotly_chart(fig3)

        elif selected_sub_menu == "FDS":
            st.title("FDS Signal monitoring")

            nmcol1 = ['FDS A', 'FDS B', 'FDS C', 'FDS D', 'FDS E', 'FDS F', 'FDS G', 'FDS H']
            nmcol2 = ['Signal A', 'Signal B', 'Signal C', 'Signal D', 'Signal E', 'Signal F', 'Signal G', 'Signal H']
            nmcol3 = ['Threshold A', 'Threshold B', 'Threshold C', 'Threshold D', 'Threshold E', 'Threshold F',
                      'Threshold G', 'Threshold H']

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                selyr = st.number_input('StartFrom', value=2015)

            def fdschart(xnm, selyr=selyr):
                st.subheader(xnm)
                df = pd.read_excel(fds_path, sheet_name=xnm)
                df = df[df['DATE'].dt.year >= selyr]
                numsig = [col for col in df.columns if col.startswith('FDS')]
                numsigx = len(numsig)
                col1x = ['DATE'] + [xnm] + nmcol2[0:numsigx]
                df = df[col1x]

                fig = make_subplots(specs=[[{"secondary_y": True}]])
                fig.add_trace(
                    go.Scatter(x=df['DATE'], y=df[xnm], mode='lines', line=dict(color='rgb(245, 130, 32)'), name=xnm),
                    secondary_y=False
                )
                for col in df.columns[2:]:
                    fig.add_trace(
                        go.Bar(x=df['DATE'], y=df[col], name=col),
                        secondary_y=True
                    )
                fig.update_layout(
                    title_text=f"Fractal Dimension Signal of {xnm}",
                    xaxis_title='Date',
                    yaxis_title=xnm,
                    yaxis2_title="Signal",
                    barmode='overlay',
                    bargap=0,
                    bargroupgap=0
                )

                return fig

            col1, col2 = st.columns(2)
            with col1:
                fig_us10y = fdschart("US10Y")
                st.plotly_chart(fig_us10y)
            with col2:
                fig_usig = fdschart("USIGOAS")
                st.plotly_chart(fig_usig)

            col1, col2 = st.columns(2)
            with col1:
                fig_jpy = fdschart("USDJPY")
                st.plotly_chart(fig_jpy)
            with col2:
                fig_krw = fdschart("USDKRW")
                st.plotly_chart(fig_krw)

            col1, col2 = st.columns(2)
            with col1:
                fig_inr = fdschart("INRKRW")
                st.plotly_chart(fig_inr)

    authenticator.logout('Logout', 'sidebar')

elif authentication_status is False:
    st.error('Username or password is incorrect')
elif authentication_status is None:
    st.warning('Please enter your username and password')

