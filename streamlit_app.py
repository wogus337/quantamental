

import streamlit as st
import pandas as pd
import openpyxl
import plotly.graph_objs as go
from plotly.subplots import make_subplots


# 페이지 제목
st.title("MAGI")

# 엑셀 파일 경로 지정
file_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24.xlsx"

# 특정 시트 지정 (예: 'Sheet1')
selected_sheet = "P1_Raw"  # 원하는 시트 이름을 지정합니다.

try:
    # 선택된 시트에서 데이터프레임 읽기
    df = pd.read_excel(file_path, sheet_name=selected_sheet)

    # Date 열이 존재하는지 확인하고, datetime 타입으로 변환
    if 'DATE' in df.columns:
        df['DATE'] = pd.to_datetime(df['DATE'])
    else:
        st.error("시트에 'DATE' 열이 없습니다.")
        st.stop()

    # 사용자에게 선택할 열을 제공, 'date' 열은 제외
    columns = [col for col in df.columns if col != 'DATE']
    col1, col2 = st.columns(2)

    with col1:
        selected_column1 = st.selectbox("Series1", ['선택 없음'] + columns)
    with col2:
        selected_column2 = st.selectbox("Series2", ['선택 없음'] + columns)

    # 데이터 다운로드 버튼
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
            # 두 번째 플롯: 두 열의 시계열 차트
            fig2 = go.Figure()

            fig2.add_trace(
                go.Scatter(x=df['DATE'], y=df[selected_column1], name=selected_column1, mode='lines'))
            fig2.add_trace(go.Scatter(x=df['DATE'], y=df[selected_column2], name=selected_column2, mode='lines',
                                      yaxis='y2'))

            # 두 열의 차이 계산
            df['difference'] = df[selected_column1] - df[selected_column2]

            # 세 번째 플롯: 두 열의 차이 시계열 차트
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


except FileNotFoundError:
    st.error("파일을 찾을 수 없습니다. 경로를 확인하세요.")
except Exception as e:
    st.error(f"파일을 열 수 없습니다: {e}")
