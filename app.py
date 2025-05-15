import streamlit as st
import pandas as pd
import io
import plotly.express as px

st.set_page_config(layout="wide", page_title="MTTD/MTTR Dashboard")

st.title("📊 MTTD & MTTR Dashboard")

uploaded_file = st.file_uploader("📁 Upload Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Convert date columns
    date_cols = ['Date/Time Opened', 'Responded Date/Time', 'Service Restored Date', 'Opened Date']
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors='coerce')

    # Add Month-Year column
    df['Month-Year'] = df['Service Restored Date'].dt.strftime('%b-%Y')

    # Priority settings
    priority_order = ['P1', 'P2', 'P3', 'P4', 'P5']
    priority_colors = {
        'P1': '#EF553B',  # red
        'P2': '#636EFA',  # blue
        'P3': '#00CC96',  # green
        'P4': '#AB63FA',  # purple
        'P5': '#FFA15A',  # orange
    }
    df['Priority'] = pd.Categorical(df['Priority'], categories=priority_order, ordered=True)

    # Calculate Final MTD and MTR
    def calc_mtd(row):
        if str(row['Case Origin']).strip().lower() in ['internal call logging', 'web']:
            return 0
        if pd.notnull(row['Responded Date/Time']) and pd.notnull(row['Date/Time Opened']):
            return (row['Responded Date/Time'] - row['Date/Time Opened']).total_seconds() / 60
        return None

    def calc_mtr(row):
        if pd.notnull(row['Service Restored Date']) and pd.notnull(row['Date/Time Opened']):
            return (row['Service Restored Date'] - row['Date/Time Opened']).total_seconds() / 60
        return None

    df['Final MTD'] = df.apply(calc_mtd, axis=1)
    df['Final MTR'] = df.apply(calc_mtr, axis=1)

    # Filters
    months = sorted(df['Month-Year'].dropna().unique(), key=lambda x: pd.to_datetime(x, format='%b-%Y'))
    priorities = [p for p in priority_order if p in df['Priority'].unique()]

    with st.sidebar:
        st.header("📅 Filters")
        selected_months = st.multiselect("Select Month-Year", months, default=months)
        selected_priorities = st.multiselect("Select Priorities", priorities, default=priorities)

    # Apply filters
    filtered_df = df[df['Month-Year'].isin(selected_months) & df['Priority'].isin(selected_priorities)]

    # Summary generator
    def generate_avg_summary(df, metric):
        summary = df.groupby(['Month-Year', 'Priority'])[metric].mean().reset_index()
        summary['Priority'] = pd.Categorical(summary['Priority'], categories=priority_order, ordered=True)
        summary = summary.sort_values(by=['Month-Year', 'Priority'], key=lambda x: x.map(lambda y: pd.to_datetime(y, format='%b-%Y')) if x.name == 'Month-Year' else x)
        return summary

    mtd_avg = generate_avg_summary(filtered_df, 'Final MTD')
    mtr_avg = generate_avg_summary(filtered_df, 'Final MTR')

    # Display tabs
    tab1, tab2, tab3 = st.tabs(["📈 MTTD Summary", "📉 MTTR Summary", "📁 Main Data"])

    with tab1:
        st.subheader("📈 Monthly Avg MTTD")
        mtd_pivot = mtd_avg.pivot(index='Month-Year', columns='Priority', values='Final MTD')
        mtd_pivot = mtd_pivot.loc[mtd_pivot.index.sort_values(key=lambda x: pd.to_datetime(x, format='%b-%Y'))]
        st.dataframe(mtd_pivot.style.format("{:.2f}"), use_container_width=True)
        fig = px.line(
            mtd_avg,
            x='Month-Year',
            y='Final MTD',
            color='Priority',
            markers=True,
            title="MTTD Over Time by Priority",
            color_discrete_map=priority_colors
        )
        fig.update_layout(
            xaxis_title='Month-Year',
            yaxis_title='Average MTTD (minutes)',
            hovermode='x unified',
            legend_title='Priority'
        )
        fig.update_traces(mode='lines+markers')
        st.plotly_chart(fig, use_container_width=True)

    with tab2:
        st.subheader("📉 Monthly Avg MTTR")
        mtr_pivot = mtr_avg.pivot(index='Month-Year', columns='Priority', values='Final MTR')
        mtr_pivot = mtr_pivot.loc[mtr_pivot.index.sort_values(key=lambda x: pd.to_datetime(x, format='%b-%Y'))]
        st.dataframe(mtr_pivot.style.format("{:.2f}"), use_container_width=True)
        fig = px.line(
            mtr_avg,
            x='Month-Year',
            y='Final MTR',
            color='Priority',
            markers=True,
            title="MTTR Over Time by Priority",
            color_discrete_map=priority_colors
        )
        fig.update_layout(
            xaxis_title='Month-Year',
            yaxis_title='Average MTTR (minutes)',
            hovermode='x unified',
            legend_title='Priority'
        )
        fig.update_traces(mode='lines+markers')
        st.plotly_chart(fig, use_container_width=True)

    with tab3:
        st.subheader("📁 Filtered Main Data")
        st.dataframe(filtered_df, use_container_width=True)

    # Full summary for export
    def full_summary(df, metric, label):
        grouped = df.groupby(['Month-Year', 'Priority']).agg(
            COUNT=(metric, 'count'),
            SUM=(metric, 'sum'),
            AVG=(metric, 'mean')
        ).reset_index()
        result = []
        for m in sorted(grouped['Month-Year'].unique(), key=lambda x: pd.to_datetime(x, format='%b-%Y')):
            row = {'Month-Year': m}
            month_data = grouped[grouped['Month-Year'] == m]
            for p in priority_order:
                match = month_data[month_data['Priority'] == p]
                if not match.empty:
                    row[f'{p} (COUNT)'] = int(match['COUNT'].values[0])
                    row[f'{p} (SUM {label})'] = match['SUM'].values[0]
                    row[f'{p} (AVG {label})'] = match['AVG'].values[0]
                else:
                    row[f'{p} (COUNT)'] = 0
                    row[f'{p} (SUM {label})'] = 0.0
                    row[f'{p} (AVG {label})'] = 0.0
            result.append(row)
        summary_df = pd.DataFrame(result)
        summary_df.set_index('Month-Year', inplace=True)
        avg_row = summary_df.mean(numeric_only=True).to_frame().T
        avg_row.index = ['AVG']
        return pd.concat([summary_df, avg_row])

    mtd_summary_full = full_summary(filtered_df, 'Final MTD', 'MTTD')
    mtr_summary_full = full_summary(filtered_df, 'Final MTR', 'MTTR')

    # Excel export
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        filtered_df.to_excel(writer, sheet_name='Main Data', index=False)
        mtd_summary_full.to_excel(writer, sheet_name='MTTD Summary')
        mtr_summary_full.to_excel(writer, sheet_name='MTTR Summary')

    st.download_button(
        label="📥 Download Full Excel Report",
        data=output.getvalue(),
        file_name="MTTD_MTTR_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
