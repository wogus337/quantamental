

import streamlit as st
import pandas as pd
import openpyxl
import plotly.graph_objs as go
from plotly.subplots import make_subplots

# 1차 메뉴 옵션
main_menu_options = ["Market", "국면판단", "유사국면", "모델전망 & Signal", "시나리오"]

# 1차 메뉴 선택
selected_main_menu = st.sidebar.selectbox("Select a Main Menu", main_menu_options)

# 2차 메뉴 옵션과 해당되는 서브 메뉴 정의
if selected_main_menu == "Market":
    st.sidebar.markdown("### Market Options")
    sub_menu_options = ["Chart", "Descriptive"]

elif selected_main_menu == "국면판단":
    st.sidebar.markdown("### 국면판단 Options")
    sub_menu_options = ["Economic Cycle", "Credit Cycle"]

elif selected_main_menu == "유사국면":
    st.sidebar.markdown("### 유사국면 Options")
    sub_menu_options = ["유사국면분석", "유사국면2"]

elif selected_main_menu == "모델전망 & Signal":
    st.sidebar.markdown("### 모델전망 & Signal Options")
    sub_menu_options = ["예측종합", "금리예측", "USIG 스프레드 예측", "장단기 스프레드 예측", "FX"]

elif selected_main_menu == "시나리오":
    st.sidebar.markdown("### 시나리오 Options")
    sub_menu_options = ["금리", "스프레드"]

selected_sub_menu = st.sidebar.selectbox("Select a Sub Menu", sub_menu_options)

if selected_main_menu == "Market":
    if selected_sub_menu == "Chart":
        st.title("MAGI")
        file_path = "data/streamlit_24.xlsx"
        selected_sheet = "P1_Raw"  # 원하는 시트 이름을 지정합니다.

        try:
            df = pd.read_excel(file_path, sheet_name=selected_sheet)

            if 'DATE' in df.columns:
                df['DATE'] = pd.to_datetime(df['DATE'])
            else:
                st.error("시트에 'DATE' 열이 없습니다.")
                st.stop()

            columns = [col for col in df.columns if col != 'DATE']
            col1, col2 = st.columns(2)

            with col1:
                selected_column1 = st.selectbox("Series1", ['선택 없음'] + columns)
            with col2:
                selected_column2 = st.selectbox("Series2", ['선택 없음'] + columns)


            def convert_df_to_csv(df):
                return df.to_csv(index=False).encode('utf-8')


            # 선택된 열의 데이터만 추출
            if selected_column1 != '선택 없음' and selected_column2 != '선택 없음' and selected_column1 != selected_column2:
                data_to_download = df[['DATE', selected_column1, selected_column2]]
                csv_data = convert_df_to_csv(data_to_download)
                st.download_button(
                    label="Data Download(CSV)",
                    data=csv_data,
                    file_name='timeseries_data.csv',
                    mime='text/csv'
                )
            elif ((selected_column1 != '선택 없음' and selected_column2 == '선택 없음') or
                  (selected_column1 != '선택 없음' and selected_column1 == selected_column2)):
                data_to_download = df[['DATE', selected_column1]]
                csv_data = convert_df_to_csv(data_to_download)
                st.download_button(
                    label="Data Download(CSV)",
                    data=csv_data,
                    file_name='timeseries_data.csv',
                    mime='text/csv'
                )
            else:
                pass

            if selected_column1 != '선택 없음':
                # 첫 번째 플롯: 선택된 열의 시계열 차트
                fig1 = go.Figure()

                fig1.add_trace(go.Scatter(x=df['DATE'], y=df[selected_column1], name=selected_column1, mode='lines'))

                fig1.update_layout(
                    title=f"{selected_column1}",
                    xaxis_title='Date',
                    yaxis_title=selected_column1,
                    template='plotly_dark'
                )

                st.plotly_chart(fig1, use_container_width=True)

                if selected_column2 != '선택 없음' and selected_column1 != selected_column2:
                    fig2 = go.Figure()
                    fig2.add_trace(
                        go.Scatter(x=df['DATE'], y=df[selected_column1], name=selected_column1, mode='lines'))
                    fig2.add_trace(go.Scatter(x=df['DATE'], y=df[selected_column2], name=selected_column2, mode='lines',
                                              yaxis='y2'))

                    df['difference'] = df[selected_column1] - df[selected_column2]
                    fig3 = go.Figure()
                    fig3.add_trace(go.Scatter(x=df['DATE'], y=df['difference'], name='Diff(1-2)', mode='lines',
                                              line=dict(color='orange')))

                    fig2.update_layout(
                        title=f"{selected_column1} & {selected_column2}",
                        xaxis_title='Date',
                        yaxis_title=selected_column1,
                        yaxis2=dict(
                            title=selected_column2,
                            overlaying='y',
                            side='right'
                        ),
                        template='plotly_dark',
                        legend=dict(
                            orientation='h',
                            yanchor='top',
                            y=1.1,
                            xanchor='center',  # 범례의 x축 앵커를 가운데로
                            x=0.5
                        )
                    )

                    fig3.update_layout(
                        title=f"Spr({selected_column1}-{selected_column2})",
                        xaxis_title='Date',
                        yaxis_title='Diff(1-2)',
                        template='plotly_dark'
                    )

                    st.plotly_chart(fig2, use_container_width=True)
                    st.plotly_chart(fig3, use_container_width=True)

    elif selected_sub_menu == "Descriptive":
        st.title("Descriptive")
        st.write("Descriptive")

elif selected_main_menu == "국면판단":
    if selected_sub_menu == "Economic Cycle":
        st.title("Economic Cycle")
        st.write("Economic Cycle")
    elif selected_sub_menu == "Credit Cycle":
        st.title("Credit Cycle")
        st.write("Credit Cycle")

elif selected_main_menu == "유사국면":
    if selected_sub_menu == "유사국면분석":
        st.title("유사국면분석")
        st.write("유사국면분석")
    elif selected_sub_menu == "유사국면2":
        st.title("유사국면2")
        st.write("유사국면2")

elif selected_main_menu == "모델전망 & Signal":
    if selected_sub_menu == "예측종합":
        st.title("예측종합")
        st.write("예측종합")
    elif selected_sub_menu == "금리예측":
        st.title("금리예측")
        st.write("금리예측")
    elif selected_sub_menu == "USIG 스프레드 예측":
        st.title("USIG 스프레드 예측")
        st.write("USIG 스프레드 예측")
    elif selected_sub_menu == "장단기 스프레드 예측":
        st.title("장단기 스프레드 예측")
        st.write("장단기 스프레드 예측")
    elif selected_sub_menu == "FX":
        st.title("FX")
        st.write("FX")

elif selected_main_menu == "시나리오":
    if selected_sub_menu == "금리":
        st.title("금리")
        st.write("금리")
    elif selected_sub_menu == "스프레드":
        st.title("스프레드")
        st.write("스프레드")
