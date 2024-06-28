import pandas as pd
import plotly.express as px 
import streamlit as st 
import plotly.graph_objects as go 
from plotly.subplots import make_subplots 
from openpyxl import load_workbook
import subprocess


file_address = "Tunnels_PBS_QBR.xlsx"

wb = load_workbook(file_address, data_only=True)
ws = wb["Pillar C Tunnels"]


# Read Excel data into a DataFrame
df_totals = pd.read_excel(file_address,
                          engine="openpyxl",
                          sheet_name="Pillar C Tunnels")
df_QBR = pd.read_excel(file_address, engine="openpyxl", sheet_name="QBR")

df_24_25 = pd.read_excel(file_address, engine="openpyxl", sheet_name="2024_25")
df_24_25_totals = pd.read_excel(file_address, engine="openpyxl", sheet_name="2024_25_TOTALS")
df_25_26 = pd.read_excel(file_address, engine="openpyxl", sheet_name="2025_26")
df_25_26_totals = pd.read_excel(file_address, engine="openpyxl", sheet_name="2025_26_TOTALS")
df_in_month = pd.read_excel(file_address, engine="openpyxl", sheet_name="In-month")
df_in_month_totals = pd.read_excel(file_address, engine="openpyxl", sheet_name="In-month_TOTALS")
df_YTD = pd.read_excel(file_address, engine="openpyxl", sheet_name="YTD")
df_YTD_totals = pd.read_excel(file_address, engine="openpyxl", sheet_name="YTD_TOTALS")

df_in_month_value = df_totals.at[89, 'Jun-24']
df_in_month_QBR = df_QBR.at[89, 'Jun-24']
df_in_month_value_result = df_in_month_value / 1000
df_in_month_QBR_result = df_in_month_QBR / 1000
formatted_df_in_month_value = f"£{df_in_month_value_result:,.0f}k"

previous_in_mounth_value = 2265.34770
df_in_month_delta = df_in_month_value_result - df_in_month_QBR_result
formatted_df_in_month_delta = f"£{df_in_month_delta:,.0f}k"


df_YTD_Pillar = df_totals.at[89, 'YTD Actuals']
df_YTD_QBR = df_QBR.at[89, 'YTD Actuals']
df_YTD_value = df_YTD_Pillar - df_YTD_QBR
df_YTD_value_result = df_YTD_value/1000
df_YTD_Pillar_result = df_YTD_Pillar/1000
formatted_df_YTD_value = f"£{df_YTD_Pillar_result:,.0f}k"
df_YTD_delta = df_YTD_value_result 
formatted_df_YTD_delta = f"-£{df_YTD_delta:,.0f}k"


df_24_25_value = df_totals.at[89, '2024/25']
df_24_25_value_result = df_24_25_value/1000
formatted_df_24_25_value = f"£{df_24_25_value_result:,.0f}k"

df_24_25_delta = df_24_25_totals.at[0, 'Variance']
df_24_25_delta_result = df_24_25_delta/1000
formatted_df_24_25_delta = f"£{df_24_25_delta_result:,.0f}k"

df_25_26_value = df_totals.at[89, '2025/26']
df_25_26_value_result = df_25_26_value/1000
formatted_df_25_26_value = f"£{df_25_26_value_result:,.0f}k"

df_25_26_delta = df_25_26_totals.at[0, 'Variance']
df_25_26_delta_result = df_25_26_delta/1000
formatted_df_25_26_delta = f"£{df_25_26_delta_result:,.0f}k"

df_EAC_value = df_totals.at[89, 'EAC (Current)']
df_EAC_value_result = df_EAC_value / 1000
formatted_df_EAC_value = f"£{df_EAC_value_result:,.0f}k"


st.set_page_config(page_title="T&A Variances Dashboard",
                   page_icon=":bar_chart:",
                   layout="wide")

st.title(':green-background[T&A Cost Forecast Variances Dashboard]')
st.markdown('#### LTC Tunnels and Approaches variances against the QBR')
st.markdown('''##### Month-end Reporting: :red[QBR June-24]''')

# Sidebar filters
# df_EAC= df_totals.groupby('Account').sum()
# df_EAC = df_EAC[['EAC (Current)']]
# df_EAC = df_EAC.sort_values('EAC (Current)', ascending=False)

#df_24_25.loc['Total'] = df_24_25.sum()
#FY_24_25_value = df_24_25[["FY 2024_25 In-month"]]


with st.container(border=True):
    st.subheader(':blue-background[Current Position to QBR]')
    a1, a2, a3, a4 = st.columns(4)
    with a1:
        tile = a1.container(height=120)
        tile.metric(label="**In-month**", value=formatted_df_in_month_value, delta=formatted_df_in_month_delta, label_visibility="visible")
    with a2:
        tile = a2.container(height=120)
        tile.metric(label="**Year To Date**", value=formatted_df_YTD_value, delta=formatted_df_YTD_delta)
        # tile.metric(label="EAC to QBR",
        #                 value=formatted_df_EAC_value,
        #                 delta="£0",
        #                 delta_color="off")
    with a3:
        tile = a3.container(height=120)
        tile.metric(label="**FY 24/25**", value=formatted_df_24_25_value, delta=formatted_df_24_25_delta)
    with a4:
        tile = a4.container(height=120)
        tile.metric(label="**FY 25/26**",
                        value=formatted_df_25_26_value,
                        delta=formatted_df_25_26_delta)



with st.container(border=True):
    st.subheader(':blue-background[In-month]')
    tab_b1, tab_b2, tab_b3 = st.tabs(["**Totals**", "**Control Account**", "**Table and Commentary**"])
    with tab_b1:
        fig_in_month_totals = px.histogram(
            df_in_month_totals,
            x="Account",
            y=["Jun-24 In-month", "Jun-24 QBR", "Variance"],
            title="Totals",
            barmode="group", text_auto='.2s')
        st.plotly_chart(fig_in_month_totals, use_container_width=True, theme="streamlit")
    with tab_b2:
        fig_in_month = px.histogram(
            df_in_month,
            x="Account",
            y=["Jun-24 In-month", "Jun-24 QBR", "Variance"],
            title="Control Account",
            barmode="group", text_auto='.2s')
        st.plotly_chart(fig_in_month, use_container_width=True, theme="streamlit")
    with tab_b3:
        b1, b2 = st.columns([0.5, 0.5])
        with b1:
            st.dataframe(df_in_month, height=460)
        with b2:
            st.markdown('##### In-month Variance Comments')
            st.markdown("**Overall In-month variance £533k:**")
            st.markdown("1. OCS (£72k) - IVL PFA Removal to re-commence in July 2024.")
            st.markdown("2. Stage 1 Scope £427k - Forecast based on cost and In-month actualised to forecasted by DP supplier.")
            st.markdown("3. Inflation £95k - Inflation rate calculation adjusted from 8.09%  to 10.52% plus 2% a year and consequential changes from Stage 1 Scope.")
            st.markdown("4. NR VAT £82k - Consequential changes from above items.")


# with st.container(border=True):
#     st.subheader(':blue-background[In-month]')
#     b1, b2 = st.columns([0.6, 0.4], gap="large")
#     with b1:
#         fig_in_month_totals = px.histogram(
#             df_in_month_totals,
#             x="Account",
#             y=["Jun-24 In-month", "Jun-24 QBR", "Variance"],
#             title="Totals",
#             barmode="group", text_auto='.2s')
#         st.plotly_chart(fig_in_month_totals, use_container_width=True, theme="streamlit")
#     with b2:
#         st.markdown('##### In-month Variance Comments')
#         st.markdown("1. OCS (£178k) - Reprofile of Network Rail £3k and IVL costs (£181k) reprofiled to JUN24.")
#         st.markdown("2. Stage 1 Scope (£111k) - In-month underspend.")
#         st.markdown("3. Inflation £27k - Inflation rate calculation adjusted to 10.52% plus 2% a year and consequential changes from Stage 1 Scope.")
#         st.markdown("4. Stage 2 Unlet Works (£30k) - UKPN non-contestable works £30k reprofiled from MAY to SEP24.")
#         st.markdown("5. NR VAT (£54k) - Consequential changes from above items.")
#     c1, c2 = st.columns([0.6, 0.4])
#     with c1:
#         fig_in_month = px.histogram(
#             df_in_month,
#             x="Account",
#             y=["Jun-24 In-month", "Jun-24 QBR", "Variance"],
#             title="Control Account",
#             barmode="group", text_auto='.2s')
#         st.plotly_chart(fig_in_month, use_container_width=True, theme="streamlit")
#     with c2:
#         st.dataframe(df_in_month)



with st.container(border=True):
    st.subheader(':blue-background[Year To Date]')
    tab_c1, tab_c2, tab_c3 = st.tabs(["**Totals**", "**Control Account**", "**Table and Commentary**"])
    with tab_c1:
        fig_YTD_totals = px.histogram(
        df_YTD_totals,
        x="Account",
        y=["YTD In-month", "YTD QBR", "Variance"],
        title="Totals",
        barmode="group", text_auto='.2s')
        st.plotly_chart(fig_YTD_totals, use_container_width=True)
    with tab_c2:
        fig_YTD = px.histogram(
        df_YTD,    
        x="Account",    
        y=["YTD In-month", "YTD QBR", "Variance"],
        title="Control Account",    
        barmode="group", text_auto='.2s')    
        st.plotly_chart(fig_YTD, use_container_width=True)            
    with tab_c3:
        e1, e2 = st.columns([0.5, 0.5])
        with e1:
            st.dataframe(df_YTD, height=460)
        with e2:
            st.markdown('##### YTD Variance Comments')
            st.markdown("**Overall YTD variance £331k:**")
            st.markdown('1. OCS (£247k) - Reprofile of Network Rail £9k and IVL PFA Removal (£256k) to re-commence in July 2024.')
            st.markdown('2. Stage 1 Scope £385k - Forecast based on cost and In-month actualised to forecasted by DP supplier £427k offset by May24 YTD (£42k).')
            st.markdown('3. Inflation £172k - Inflation rate calculation adjusted from 8.09% to 10.52% plus 2% a year and consequential changes from Stage 1 Scope.')
            st.markdown('4. Stage 2 Unlet Works (£30k) - UKPN non-contestable works £30k reprofiled from May to September 2024.')
            st.markdown('5. NR VAT £51k - Consequential changes from above items.')


with st.container(border=True):
    st.subheader(':blue-background[FY 2024_25]')
    tab_f1, tab_f2, tab_f3 = st.tabs(["**Totals**", "**Control Account**", "**Table and Commentary**"])
    with tab_f1:
        fig_24_25_totals = px.histogram(
            df_24_25_totals,
            x="Account",
            y=["FY 2024_25 In-month", "FY 2024_25 QBR", "Variance"],
            title="Totals",
            barmode="group", text_auto='.2s')
        st.plotly_chart(fig_24_25_totals, use_container_width=True)
    with tab_f2:
        fig_24_25 = px.histogram(
                df_24_25,
                x="Account",
                y=["FY 2024_25 In-month", "FY 2024_25 QBR", "Variance"],
                title="Control Account",
                barmode="group", text_auto='.2s')
        st.plotly_chart(fig_24_25, use_container_width=True)
    with tab_f3:
        g1, g2 = st.columns([0.5, 0.5])
        with g1:
            st.dataframe(df_24_25, height=460)
        with g2:
            st.markdown('##### FY 24/25 Variance Comments')
            st.markdown("**Overall FY 24/25 variance £3,243k:**")
            st.markdown('1. OCS £24k - Reprofile of Network Rail costs')
            st.markdown('2. Stage 1 Scope £1,873k - £1,200k BMJV Phase 1 de-risking transfer of costs from FY25/26 and £673k from Trends of CHP and NP access bridge')
            st.markdown('3. Inflation £842k - Inflation rate calculation adjusted from 8.09% to 10.52% plus 2% a year and consequential changes from Stage 1 Scope')
            st.markdown('4. NR VAT £504k - Consequential changes from above items')


with st.container(border=True):
    st.subheader(':blue-background[FY 2025_26]')
    tab_h1, tab_h2, tab_h3 = st.tabs(["**Totals**", "**Control Account**", "**Table and Commentary**"])
    with tab_h1:
        #st.markdown('### FY 2024_25 Chart')
        fig_25_26_totals = px.histogram(
            df_25_26_totals,
            x="Account",
            y=["FY 2025_26 In-month", "FY 2025_26 QBR", "Variance"],
            title="Totals",
            barmode="group", text_auto='.2s')
        st.plotly_chart(fig_25_26_totals, use_container_width=True)
    with tab_h2:
        #st.markdown('### FY 2025_26 Chart')
        fig_25_26 = px.histogram(
            df_25_26,
            x="Account",
            y=["FY 2025_26 In-month", "FY 2025_26 QBR", "Variance"],
            title="Control Account",
            barmode="group", text_auto='.2s')
        st.plotly_chart(fig_25_26, use_container_width=True)
    with tab_h3:
        i1, i2 = st.columns([0.5, 0.5])
        with i1:
            st.dataframe(df_25_26, height=460)
        with i2:
            #st.markdown('### FY 2025_26 Table')
            st.markdown('##### FY 25/26 Variance Comments')
            st.markdown("**Overall FY 25/26 variance £27,834k:**")
            st.markdown('1. Stage 1 Scope (£5,453) - (5,886k) BMJV Phase 1 de-risking transfer of costs to FY24/25 and FY26/27 and £433k from Trends of CHP and NP access bridge')
            st.markdown('2. Stage 2 Unlet Works £25,926k - £5,160k Early Works-Utilities & Programme de-risking £5,160k and £20,766k from Trends')
            st.markdown('3. Inflation £3,036k - Inflation rate calculation adjusted from 8.09% to 10.52% plus 2% a year')
            st.markdown('4. NR VAT £4,326k - Consequential changes from above items')


#st.page_link("https://lowerthamescrossing.sharepoint.com/:x:/s/Prism/ERTBK6KjBdNIr4BN87Ny8AYBFWv9f64o9h2zRGjWBosGkg?e=y7vqFc", label="Tunnels Pillar Table", icon="🌎")
st.link_button("Go to Tunnels Pillar Sheet :sunglasses:", "https://lowerthamescrossing.sharepoint.com/:x:/s/HECommercialReporting/Ed4JH7-26lVFlKYWE7I3wLQBAOQ3qL55-aOTxonYyMQTKA?e=xKahZj", type="primary")

# Execute the external Python script TDR.py
def run_external_script():
    try:
        subprocess.run(["python", r"C:\Users\marcelo.defranceschi\OneDrive - Mace\Desktop\Projetos Python\PYTHON_files\Oracle_forecast_by_task2.py"], check=True)
        st.write("Python script executed successfully!")
    except subprocess.CalledProcessError:
        st.write("Error executing python script. Please check the script and try again.")

# Create a Streamlit button to run the external script
if st.button("Run python script"):
    run_external_script()


# with st.container(border=True):
#     st.markdown('##### Comments FY 25_26')
#     st.markdown('Explain variances here')
#     st.markdown('And here too')

# with st.expander("### Variances explanation", expanded=False):
#     st.markdown('Variances are explained below: :sunglasses:')
#     tile1 = st.container(height=60)
#     tile1.markdown('1-aaaaaaaaaaaaa')
#     st.markdown('2-bbbbbbbbbbbb')

# with st.container(border=True):
#     b1, b2 = st.columns(2)
#     with b1:
#         st.markdown('### EAC by Account')
#         st.bar_chart(df_EAC, color="#EA832E")
#         #fig1 = px.bar(df, x=Projetos, y=["EAC (Current)"])
#         #st.plotly_chart(fig1)
#     with b2:
#         st.markdown('### Pie Chart')
#         fig_pie = px.pie(df_totals,
#                          values=["EAC (Current)"],
#                          names=["Account"])
#         st.plotly_chart(fig_pie, use_container_width=True)



# Original table to be visible
#with st.container(border=True):
    #st.markdown('##### Tunnels Pillar Table')
    #st.data_editor(df_totals)


#fig2 = px.pie(df, values=["EAC (Current)"], names=["Projects"])
#st.plotly_chart(fig2)

#st.text_input(label="Comments Input")





# import plotly.graph_objects as go
# from plotly.subplots import make_subplots
# import pandas as pd

# Create sample data (replace with your actual data)
# df = pd.DataFrame({
#     'Days': [1, 2, 3, 4, 5],
#     'Perc_Cases': [15, 6, 5, 4, 3],
#     'Count_Cases': [1, 4, 2, 5, 4]
# })

# # Set up the Plotly figure with subplots
# fig = make_subplots(1, 2)
# # Add the bar trace to the first subplot
# fig.add_trace(go.Bar(x=df['Days'], y=df['Count_Cases'], name="Absolute Cases", marker_color='green', opacity=0.5), row=1, col=1)
# # Add the line trace to the first subplot
# fig.add_trace(go.Scatter(x=df['Days'], y=df['Perc_Cases'], mode='lines+markers', name="Percentage Cases", marker_color='crimson'), row=1, col=1)
# # Customize the layout
# fig.update_layout(
#     title_text='COVID-19 Cases',
#     yaxis=dict(range=[0, 100], side='right'))
# #Show the plot
# st.plotly_chart(fig, use_container_width=True)





#python -m streamlit run Tunnels_app.py
