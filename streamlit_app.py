

import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import plotly.graph_objs as go
from plotly.subplots import make_subplots
import datetime

st.set_page_config(layout="wide")

#image_path = r"D:\Anaconda_envs\streamlit\pycharmprj\miraeasset.png"
image_path = "images/miraeasset.png"

st.sidebar.image(image_path, use_column_width=True, output_format='PNG')
st.sidebar.write("")
st.sidebar.title("QIS")
main_menu_options = ["Market", "국면판단", "유사국면", "모델전망 & Signal", "시나리오"]
selected_main_menu = st.sidebar.selectbox("Select a Main Menu", main_menu_options)

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

        st.title("Chart")
        file_path = "data/streamlit_24.xlsx"
        #file_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24.xlsx"
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

            if selected_column1 != '선택 없음':
                with col1:
                    sdate = df['DATE'].min().strftime('%Y/%m/%d')
                    edate = df['DATE'].max().strftime('%Y/%m/%d')
                    st.header(f"Date: {sdate} ~ {edate}")
                    start_date = st.date_input("Start", min_value=df['DATE'].min(), max_value=df['DATE'].max(), value=df['DATE'].min())
                    st.write("")
                with col2:
                    st.header("")
                    end_date = st.date_input("End", min_value=df['DATE'].min(), max_value=df['DATE'].max(), value=df['DATE'].max())
                    st.write("")
                fdf = df[(df['DATE'] >= pd.to_datetime(start_date)) & (df['DATE'] <= pd.to_datetime(end_date))]
                st.subheader(f"{selected_column1}")
                recent_data1 = fdf[['DATE', selected_column1]].tail(5)
                recent_data1.set_index('DATE', inplace=True)
                st.dataframe(recent_data1, use_container_width=True)

                fig1 = go.Figure()
                fig1.add_trace(go.Scatter(x=fdf['DATE'], y=fdf[selected_column1], name=selected_column1, mode='lines'))
                fig1.update_layout(
                    xaxis_title='Date',
                    yaxis_title=selected_column1,
                    template='plotly_dark'
                )

                st.plotly_chart(fig1, use_container_width=True)

                if selected_column2 != '선택 없음' and selected_column1 != selected_column2:

                    with col1:
                        st.write("")
                        st.write("")
                        st.write("")
                        st.write("")
                        st.write("")
                        st.write("")
                        st.subheader(f"{selected_column1} & {selected_column2}")
                        recent_data2 = fdf[['DATE', selected_column1, selected_column2]].tail(5)
                        recent_data2.set_index('DATE', inplace=True)
                        st.dataframe(recent_data2, use_container_width=True)

                        fig2 = go.Figure()
                        fig2.add_trace(
                            go.Scatter(x=fdf['DATE'], y=fdf[selected_column1], name=selected_column1, mode='lines'))
                        fig2.add_trace(go.Scatter(x=fdf['DATE'], y=fdf[selected_column2], name=selected_column2, mode='lines',
                                                  yaxis='y2'))

                        fig2.update_layout(
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
                        st.plotly_chart(fig2, use_container_width=True)

                    with col2:
                        selectr = st.radio("Relative:",
                                           ["Spread", "Ratio"]
                        )

                        if selectr == "Spread":
                            fdf['rel'] = fdf[selected_column1] - fdf[selected_column2]
                            st.subheader(f"Spr({selected_column1}-{selected_column2})")
                        elif selectr == "Ratio":
                            fdf['rel'] = fdf[selected_column1]/fdf[selected_column2]
                            st.subheader(f"Ratio({selected_column1}/{selected_column2})")

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
            if selected_column1 != '선택 없음' and selected_column2 != '선택 없음' and selected_column1 != selected_column2:
                data_to_download = fdf[['DATE', selected_column1, selected_column2, 'rel']]
                csv_data = convert_df_to_csv(data_to_download)
                st.download_button(
                    label="Data Download(CSV)",
                    data=csv_data,
                    file_name='timeseries_data.csv',
                    mime='text/csv'
                )
            elif ((selected_column1 != '선택 없음' and selected_column2 == '선택 없음') or
                  (selected_column1 != '선택 없음' and selected_column1 == selected_column2)):
                data_to_download = fdf[['DATE', selected_column1]]
                csv_data = convert_df_to_csv(data_to_download)
                st.download_button(
                    label="Data Download(CSV)",
                    data=csv_data,
                    file_name='timeseries_data.csv',
                    mime='text/csv'
                )
            else:
                pass
        except FileNotFoundError:
            st.error("파일을 찾을 수 없습니다. 경로를 확인하세요.")
        except Exception as e:
            st.error(f"파일을 열 수 없습니다: {e}")

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

        simfile_path = "data/streamlit_24_sim.xlsx"
        #simfile_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24_sim.xlsx"

        df_raw = pd.read_excel(simfile_path, sheet_name='RawdataSim')
        df_simdt = pd.read_excel(simfile_path, sheet_name='siminfo')

        df_raw['DATE'] = pd.to_datetime(df_raw['DATE'])
        df_simdt['SDATE'] = pd.to_datetime(df_simdt['SDATE'])
        df_simdt['EDATE'] = pd.to_datetime(df_simdt['EDATE'])
        df_simdt['SDATE_SIM'] = pd.to_datetime(df_simdt['SDATE_SIM'])
        df_simdt['EDATE_SIM'] = pd.to_datetime(df_simdt['EDATE_SIM'])

        sel_edt = st.selectbox("분석기준일을 선택하면 해당기준일에 산출한 유사국면 리스트가 생성됩니다.:", df_simdt['EDATE'].unique())

        sel_df = df_simdt[df_simdt['EDATE'] == sel_edt]
        st.table(sel_df)

        fil_dt = df_simdt[df_simdt['EDATE'] == sel_edt]['EDATE_SIM']
        sel_simdt = st.selectbox("산출된 유사국면을 선택하면 정보가 표시됩니다.:", fil_dt)

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            selectafter = st.radio("After:", ["20D", "40D", "60D"])

        if selectafter == "20D":
            numafter = 20
        elif selectafter == "40D":
            numafter = 40
        else:
            numafter = 60

        sdate = pd.to_datetime(df_simdt[(df_simdt['EDATE'] == sel_edt) & (df_simdt['EDATE_SIM'] == sel_simdt)]['SDATE'].values[0])
        edate = pd.to_datetime(df_simdt[(df_simdt['EDATE'] == sel_edt) & (df_simdt['EDATE_SIM'] == sel_simdt)]['EDATE'].values[0])
        sdate_sim = pd.to_datetime(df_simdt[(df_simdt['EDATE'] == sel_edt) & (df_simdt['EDATE_SIM'] == sel_simdt)]['SDATE_SIM'].values[0])
        edate_sim = pd.to_datetime(df_simdt[(df_simdt['EDATE'] == sel_edt) & (df_simdt['EDATE_SIM'] == sel_simdt)]['EDATE_SIM'].values[0])

        lenk = df_simdt[(df_simdt['EDATE'] == sel_edt) & (df_simdt['EDATE_SIM'] == sel_simdt)]['LEN_K'].values[0]
        lens = df_simdt[(df_simdt['EDATE'] == sel_edt) & (df_simdt['EDATE_SIM'] == sel_simdt)]['LEN_S'].values[0]
        numaftsim = np.round((lens / lenk) * numafter).astype(int)
        #lenf = max(lenk, lens)

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
                dfb['dindex'] = np.round(np.linspace(1, len(dfa)+1, num=len(dfb))).astype(int)
                dfb = dfb.rename(columns={colnm: f"{colnm}_sim"})
                dfc = pd.merge(dfa, dfb, on='dindex', how='left')
            else:
                dfa['dindex'] = range(1, len(dfa) + 1)
                indices = np.round(np.linspace(1, len(dfb)+1, num=len(dfa))).astype(int)
                indices = np.clip(indices, 0, len(dfb) - 1)
                dfb = dfb.iloc[indices].reset_index(drop=True)  # 인덱스 리셋
                dfb['dindex'] = range(1, len(dfb) + 1)
                dfb = dfb.rename(columns={colnm: f"{colnm}_sim"})
                dfc = pd.merge(dfa, dfb, on='dindex', how='left')

            dfc_a = dfc.iloc[0:120]
            dfc_b = dfc.iloc[120:]
            dfc_a[colnm] = dfc_a[colnm].interpolate()
            dfc_a[f"{colnm}_sim"] = dfc_a[f"{colnm}_sim"].interpolate()
            dfc = pd.concat([dfc_a, dfc_b])

            fig = go.Figure()
            fig.add_trace(go.Scatter(x=dfc['dindex'], y=dfc[colnm], mode='lines',
                                     name=f'{colnm} ({sdate.date()} ~ {edate.date()})'))
            fig.add_trace(go.Scatter(x=dfc['dindex'], y=dfc[f'{colnm}_sim'], mode='lines',
                                     name=f'{colnm}_sim ({sdate_sim.date()} ~ {edate_sim.date()})', yaxis='y2'))

            specificx = 120
            fig.add_vline(x=specificx, line_width=1, line_dash="dash", line_color="red")

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
            st.plotly_chart(fig_US10Y)
            st.plotly_chart(fig_DXY)
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
