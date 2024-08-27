

import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import plotly.graph_objs as go
from plotly.subplots import make_subplots
import datetime

st.set_page_config(layout="wide")

series_path = "data/streamlit_24.xlsx"
cylfile_path = "data/streamlit_24_cycle.xlsx"
simfile_path = "data/streamlit_24_sim.xlsx"
fx_path = "data/streamlit_24_fx.xlsx"
model_path = "data/streamlit_24_signal.xlsx"
image_path = "images/miraeasset.png"

# series_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24.xlsx"
# cylfile_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24_cycle.xlsx"
# simfile_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24_sim.xlsx"
# fx_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24_fx.xlsx"
# model_path = r"\\172.16.130.210\채권운용부문\FMVC\Monthly QIS\making_files\SC_2408\streamlit_24_signal.xlsx"
# image_path = r"D:\Anaconda_envs\streamlit\pycharmprj\miraeasset.png"

st.sidebar.image(image_path, use_column_width=True, output_format='PNG')
st.sidebar.write("")
st.sidebar.title("QIS")
#main_menu_options = ["Market", "국면", "유사국면", "모델전망 & Signal", "Allocation", "시나리오"]
main_menu_options = ["Market", "국면", "유사국면", "모델전망 & Signal"]
selected_main_menu = st.sidebar.selectbox("Select a Main Menu", main_menu_options)

if selected_main_menu == "Market":
    sub_menu_options = ["ChartBoard", "Relative"]

elif selected_main_menu == "국면":
    sub_menu_options = ["Economic Cycle", "Credit Cycle"]

elif selected_main_menu == "유사국면":
    sub_menu_options = ["유사국면분석"]

#elif selected_main_menu == "시나리오":
#    sub_menu_options = ["금리", "스프레드"]

elif selected_main_menu == "모델전망 & Signal":
    sub_menu_options = ["금리", "USIG스프레드", "FX"]

#elif selected_main_menu == "Allocation":
#    sub_menu_options = ["Region", "US_Sector", "USIG_Sector"]

selected_sub_menu = st.sidebar.selectbox("Select a Sub Menu", sub_menu_options)

if selected_main_menu == "Market":
    if selected_sub_menu == "ChartBoard":

        st.title("Chart Borad")

        df = pd.read_excel(series_path, sheet_name='P1_Raw')
        selecpr = st.radio("", ["1M", "3M", "6M", "1Y", "3Y", "5Y", "10Y"], horizontal=True)
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

        def chartgen(vname, rawdf=df):
            xdf = rawdf[['DATE', vname]]
            fdf = xdf[(xdf['DATE'] >= pd.to_datetime(sdate)) & (xdf['DATE'] <= pd.to_datetime(edate))]
            fdf = fdf.dropna()
            fig1 = go.Figure()
            fig1.add_trace(go.Scatter(x=fdf['DATE'], y=fdf[vname], name=vname, mode='lines'))
            fig1.update_layout(
                xaxis_title='Date',
                yaxis_title=vname,
                template='plotly_dark'
            )
            return fig1

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            fig_US10Y = chartgen('USGG10YR')
            st.subheader("USGG10YR")
            st.plotly_chart(fig_US10Y)

            fig_USIGOAS = chartgen('USIG_OAS')
            st.subheader("USIG_OAS")
            st.plotly_chart(fig_USIGOAS)

            fig_MSCIA = chartgen('MSCI_ACWI')
            st.subheader("MSCI_ACWI")
            st.plotly_chart(fig_MSCIA)

            fig_OIL = chartgen('OIL')
            st.subheader("WTI")
            st.plotly_chart(fig_OIL)

        with col2:
            fig_US2Y = chartgen('USGG2YR')
            st.subheader("USGG2YR")
            st.plotly_chart(fig_US2Y)

            fig_USHYOAS = chartgen('USHY_OAS')
            st.subheader("USHY_OAS")
            st.plotly_chart(fig_USHYOAS)

            fig_MSCIE = chartgen('MSCI_EM')
            st.subheader("MSCI_EM")
            st.plotly_chart(fig_MSCIE)

            fig_DXY = chartgen('DXY')
            st.subheader("Dollar Index")
            st.plotly_chart(fig_DXY)

        with col3:
            fig_US5Y = chartgen('USGG5YR')
            st.subheader("USGG5YR")
            st.plotly_chart(fig_US5Y)

            fig_SP500 = chartgen('SP500')
            st.subheader("S&P500")
            st.plotly_chart(fig_SP500)

            fig_KOSPI = chartgen('KOSPI')
            st.subheader("KOSPI")
            st.plotly_chart(fig_KOSPI)

        with col4:
            fig_US30Y = chartgen('USGG30YR')
            st.subheader("USGG30YR")
            st.plotly_chart(fig_US30Y)

            fig_NASDAQ = chartgen('NASDAQ')
            st.subheader("NASDAQ")
            st.plotly_chart(fig_NASDAQ)

            fig_GOLD = chartgen('GOLD')
            st.subheader("GOLD")
            st.plotly_chart(fig_GOLD)

if selected_main_menu == "Market":
    if selected_sub_menu == "Relative":

        st.title("Relative")

        try:
            df = pd.read_excel(series_path, sheet_name='P1_Raw')

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

elif selected_main_menu == "국면":
    if selected_sub_menu == "Economic Cycle":
        st.title("Economic Cycle")

        tseries = pd.read_excel(series_path, sheet_name='P1_Raw')
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
            line=dict(width=4, color='red')
        ))
        fig.add_trace(go.Scatter(
            x=fdf['DATE'], y=fdf[selected_column1],
            mode='lines', name=selected_column1,
            yaxis='y2',
            line=dict(width=4, color='blue')
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
            autosize=True
        )

        st.plotly_chart(fig)

    elif selected_sub_menu == "Credit Cycle":
        st.title("Credit Cycle")

        tseries = pd.read_excel(series_path, sheet_name='P1_Raw')
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
            line=dict(width=4, color='red')
        ))
        fig.add_trace(go.Scatter(
            x=fdf['DATE'], y=fdf[selected_column1],
            mode='lines', name=selected_column1,
            yaxis='y2',
            line=dict(width=4, color='blue')
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
            autosize=True
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

            dfc[f"{colnm}_sim"] = dfc[f"{colnm}_sim"].interpolate()
            dfc_a = dfc.iloc[0:120]
            dfc_b = dfc.iloc[120:]
            dfc_a[colnm] = dfc_a[colnm].interpolate()
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

    elif selected_sub_menu == "유사국면2":
        st.title("유사국면2")
        st.write("유사국면2")

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
                    <p>3. 주황색 영역과 남색 막대의 부호가 같으면 방향이 적중했음을 의미</p>
                    <p>4. 우측 테이블의 가장 하단값은 이번 주의 시그널(Actual값 없음)</p>
                </div>
                """
        st.markdown(html, unsafe_allow_html=True)

        st.write("")
        st.subheader("Duration Mdoel1(선행변수모델) - Weekly")
        col1, col2 = st.columns([3, 1])
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
                title_text="Duration Mdoel1(선행변수모델) - Weekly",
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
        col1, col2 = st.columns([3, 1])
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

    elif selected_sub_menu == "USIG스프레드":
        dfw = pd.read_excel(model_path, sheet_name='Week')
        dfm = pd.read_excel(model_path, sheet_name='Month')

        st.title("Model Forecast - Credit")
        st.write("")

        html = """
                <style>
                    .custom-text {
                        line-height: 1.2; /* 행간을 줄이는 CSS 속성 */
                    }
                </style>
                <div class="custom-text">
                    <p>1. 좌축: 주황색 영역이 +/- 이면, 스프레드 축소/확대 시그널이며, 없으면 중립시그널</p>
                    <p>2. 우축: 금요일 기준 시그널 발생 이후 1주간(월~월)의 실제 스프레드 등락폭</p>
                    <p>3. 주황색 영역과 남색 막대의 부호가 같으면 방향이 적중했음을 의미</p>
                    <p>4. 우측 테이블의 가장 하단값은 이번 주의 시그널(Actual값 없음)</p>
                </div>
                """
        st.markdown(html, unsafe_allow_html=True)

        st.write("")
        st.subheader("Credit Mdoel1 - Weekly")
        col1, col2 = st.columns([3, 1])
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

        st.subheader("Credit Mdoel2 - Weekly")
        col1, col2 = st.columns([3, 1])
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

    elif selected_sub_menu == "FX":
        st.title("FX Strategy by Transformer")

        def fxgenfig1(xrange, fxnm, selprob, chart_title, df_path=fx_path):
            df = pd.read_excel(df_path, sheet_name='fx', usecols=xrange, skiprows=0)
            colnm = ['DATE', 'Prob0', 'Prob1', fxnm, 'fx_v', 'Conviction', 'FX_Long', 'Strategy']
            df.columns = colnm
            df = df[df['DATE'].notna()]
            fdf = df[df['DATE'] >= pd.Timestamp('2021-01-04')]

            fdf.set_index('DATE', inplace=True)
            all_dates = pd.date_range(start=fdf.index.min(), end=fdf.index.max(), freq='D')
            fdf = fdf.reindex(all_dates).ffill()

            fig1 = go.Figure()
            fig1.add_trace(go.Bar(
                x=fdf.index, y=fdf[selprob],
                yaxis='y2',
                opacity=0.4,
                showlegend=False,
                marker=dict(line=dict(width=0))
            ))
            fig1.add_trace(go.Bar(
                x=fdf.index, y=fdf['Conviction'],
                name='Conviction',
                yaxis='y2',
                opacity=0.4,
                marker=dict(line=dict(width=0))
            ))
            fig1.add_trace(go.Scatter(
                x=fdf.index, y=fdf[fxnm],
                mode='lines', name=f'{fxnm}',
                yaxis='y1',
                line=dict(width=4, color='red')
            ))

            fig1.update_layout(
                title=chart_title,
                xaxis=dict(title='DATE'),
                yaxis=dict(
                    title=fxnm,
                    side='left'
                ),
                yaxis2=dict(
                    title='Conviction',
                    overlaying='y',
                    side='right'
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
                )
            )

            fig2 = go.Figure()
            fig2.add_trace(go.Scatter(
                x=fdf.index, y=fdf['FX_Long'],
                mode='lines', name=f'{fxnm}',
                yaxis='y1',
                line=dict(width=4, color='red')
            ))
            fig2.add_trace(go.Scatter(
                x=fdf.index, y=fdf['Strategy'],
                mode='lines', name='Strategy',
                yaxis='y1',
                line=dict(width=4, color='blue')
            ))
            fig2.update_layout(
                xaxis_title='DATE',
                yaxis_title='Return',
                template='plotly_dark',
                legend=dict(
                    orientation='h',
                    yanchor='top',
                    y=1.1,
                    xanchor='center',  # 범례의 x축 앵커를 가운데로
                    x=0.5
                )
            )
            return fig1, fig2

        def fxgenfig2(xrange, fxnm, chart_title, df_path=fx_path):
            df = pd.read_excel(df_path, sheet_name='fx', usecols=xrange, skiprows=0)
            colnm = ['DATE', 'Prob0', 'Prob1', fxnm, 'fx_v', 'Conviction']
            df.columns = colnm
            df = df[df['DATE'].notna()]
            fdf = df[df['DATE'] >= pd.Timestamp('2021-01-04')]

            fdf.set_index('DATE', inplace=True)
            all_dates = pd.date_range(start=fdf.index.min(), end=fdf.index.max(), freq='D')
            fdf = fdf.reindex(all_dates).ffill()

            fig1 = go.Figure()
            fig1.add_trace(go.Bar(
                x=fdf.index, y=fdf['Prob1'],
                yaxis='y2',
                opacity=0.4,
                showlegend=False,
                marker=dict(line=dict(width=0))
            ))
            fig1.add_trace(go.Bar(
                x=fdf.index, y=fdf['Conviction'],
                name='Conviction',
                yaxis='y2',
                opacity=0.4,
                marker=dict(line=dict(width=0))
            ))
            fig1.add_trace(go.Scatter(
                x=fdf.index, y=fdf[fxnm],
                mode='lines', name=f'{fxnm}',
                yaxis='y1',
                line=dict(width=4, color='red')
            ))

            fig1.update_layout(
                title=chart_title,
                xaxis=dict(title='DATE'),
                yaxis=dict(
                    title=fxnm,
                    side='left'
                ),
                yaxis2=dict(
                    title='Conviction',
                    overlaying='y',
                    side='right'
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
                )
            )
            return fig1

        fig_USDKRW1, fig_USDKRW2 = fxgenfig1('B:I', 'USDKRW', 'Prob1', 'USDKRW: USD 강세 모델')
        fig_KRWUSD1, fig_KRWUSD2 = fxgenfig1('O:V', 'USDKRW', 'Prob0', 'USDKRW: KRW 강세 모델')
        fig_USDEUR1 = fxgenfig2('AJ:AO', 'USDEUR', 'USDEUR')
        fig_USDGBP1 = fxgenfig2('AQ:AV', 'USDGBP', 'EURUSD')
        fig_USDCNY1 = fxgenfig2('AX:BC', 'USDCNY', 'USDCNY')
        fig_USDJPY1 = fxgenfig2('BE:BJ', 'USDJPY', 'USDJPY')

        col1, col2 = st.columns(2)
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


elif selected_main_menu == "시나리오":
    if selected_sub_menu == "금리":
        st.title("금리")
        st.write("금리")
    elif selected_sub_menu == "스프레드":
        st.title("스프레드")
        st.write("스프레드")
