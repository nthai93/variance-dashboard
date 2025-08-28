# ================================================================
# 📊 Variance Dashboard – Plan vs Actual (Fixed Version)
# Streamlit dashboard with Gantt, Pareto, Pie, Heatmap, Export HTML
# ================================================================

# %% 🔹 Import Libraries
import os, warnings, io, base64
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from PIL import Image as PILImage

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
import base64

def img_to_base64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

# %% 🔹 Helper: Convert Matplotlib or Uploaded Image to Base64 <img>
def fig_to_base64_img(fig, dpi=120):
    if fig is None:
        return ""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight")
    encoded = base64.b64encode(buf.getvalue()).decode("utf-8")
    return f"<img class='chart' src='data:image/png;base64,{encoded}'/>"

def file_to_base64_img(uploaded_file):
    if uploaded_file is None:
        return ""
    data = uploaded_file.getvalue()
    encoded = base64.b64encode(data).decode("utf-8")
    return f"<img class='chart' src='data:image/png;base64,{encoded}'/>"

# %% 🔹 Streamlit Page Setup
st.set_page_config(page_title="Variance Dashboard", layout="wide")
st.write("✅ App started successfully!")

# %% 🔹 Upload Excel + Gantt PNG
uploaded_excel = st.file_uploader("📤 Upload file summary_plan_vs_actual.xlsx", type=["xlsx"])
uploaded_png = st.file_uploader("📤 Upload Gantt chart (optional, .png)", type=["png"])

# %% 🔹 Load Excel
if uploaded_excel is not None:
    try:
        df = pd.read_excel(io.BytesIO(uploaded_excel.getvalue()))
        st.success("✅ Excel data loaded successfully")
    except Exception as e:
        st.error(f"⚠️ Could not read Excel file: {e}")
        st.stop()
else:
    st.error("⚠️ Please upload an Excel file to view the dashboard")
    st.stop()

# %% 🔹 Clean & Rename Columns
df.columns = df.columns.str.strip()
rename_map = {
    "Ngày": "Date", "Máy": "Machine", "Mã SP": "ItemCode",
    "Trễ thời gian": "Delay (min)", "Cảnh báo": "Alert",
    "Nguyên nhân": "Reason", "Ghi chú": "Note"
}
df.rename(columns=rename_map, inplace=True)

# Parse datetime columns
for col in ["Plan Start", "Actual Start"]:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], format="%H:%M", errors="coerce")

# Parse date column
if "Date" in df.columns:
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date

# Drop duplicate columns if already included
cols_to_drop = ["Date", "Machine", "ItemCode", "Delay (min)", "Alert", "Reason", "Note"]
df.drop(columns=[c for c in cols_to_drop if c in df.columns.duplicated()], errors="ignore", inplace=True)

# %% 🔹 Compute KPIs
total_jobs = len(df)
num_alerts = df["Alert"].notna().sum() if "Alert" in df.columns else 0
alert_rate = round(num_alerts / total_jobs * 100, 2) if total_jobs else 0

# %% 🔹 Overview Section
st.title("📊 Variance Dashboard – Plan vs Actual")
with st.expander("📈 Overview", expanded=True):
    total_machines = df["Machine"].nunique() if "Machine" in df.columns else 0
    alert_machines = df[df["Alert"].notna()]["Machine"].nunique() if "Alert" in df.columns else 0
    machine_alert_rate = round(alert_machines / total_machines * 100, 2) if total_machines else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("🧾 Total Jobs", total_jobs)
    col2.metric("🚨 Jobs with Alert", num_alerts)
    col3.metric("⚠️ Alert Rate", f"{alert_rate}%")
    col4.metric("🏭 Active Machines", total_machines)

    col5, col6 = st.columns(2)
    col5.metric("💥 Machines with Alert", alert_machines)
    col6.metric("📊 Machines Alert Rate", f"{machine_alert_rate}%")

    if "Delay (min)" in df.columns and not df.empty:
        worst_job = df.loc[df["Delay (min)"].idxmax()]
        st.markdown(f"""
        ### 🛑 Job with Max Delay
        - ⏰ **{worst_job['Delay (min)']} minutes**
        - 🏭 Machine: **{worst_job.get('Machine', 'N/A')}**
        - 🧾 ItemCode: **{worst_job.get('ItemCode', 'N/A')}**
        - 📆 Date: **{worst_job.get('Date', 'N/A')}**
        """)

        # Threshold Alerts
        threshold_100 = df[df["Delay (min)"] > 100].shape[0]
        threshold_200 = df[df["Delay (min)"] > 200].shape[0]
        st.markdown(f"""
        ### 📊 Jobs exceeding thresholds
        - 🟥 Delay > 200 min: **{threshold_200}**
        - 🟧 Delay > 100 min: **{threshold_100}**
        """)

    # Top Reasons + Delay by Machine
    if "Reason" in df.columns:
        top_reasons = df[df["Reason"].notna()]["Reason"].value_counts().head(3)
        st.markdown("### 🧠 Top 3 Reasons:")
        for i, (reason, count) in enumerate(top_reasons.items(), start=1):
            st.markdown(f"- {i}. **{reason}** – {count} jobs")
    else:
        top_reasons = pd.Series(dtype=int)

    if "Delay (min)" in df.columns and "Machine" in df.columns:
        delay_by_machine = (
            df[df["Delay (min)"] > 0].groupby("Machine")["Delay (min)"].sum()
            .sort_values(ascending=False).head(3)
        )
        st.markdown("### 🏭 Top 3 Machines with Highest Total Delay:")
        st.dataframe(delay_by_machine.reset_index().rename(columns={"Delay (min)": "Total Delay"}))
    else:
        delay_by_machine = pd.Series(dtype=int)

# %% 🔹 Alert Job Details
with st.expander("📋 Jobs with Alerts", expanded=True):
    alert_df = df[df["Alert"].notna()] if "Alert" in df.columns else pd.DataFrame()
    show_cols = ["Date", "Machine", "ItemCode", "Delay (min)", "Reason", "Note"]
    show_cols = [col for col in show_cols if col in df.columns]

    if not alert_df.empty:
        for reason in top_reasons.index:
            st.markdown(f"**🔸 {reason}** – {len(alert_df[alert_df['Reason'] == reason])} jobs:")
            st.dataframe(alert_df[alert_df["Reason"] == reason][show_cols], use_container_width=True)

        for mach in delay_by_machine.index:
            st.markdown(f"**🔸 {mach}** – {len(alert_df[alert_df['Machine'] == mach])} jobs:")
            st.dataframe(alert_df[alert_df["Machine"] == mach][show_cols], use_container_width=True)

# %% 🔹 Gantt Chart Display
with st.expander("📉 Gantt Chart"):
    if uploaded_png:
        try:
            image = PILImage.open(io.BytesIO(uploaded_png.getvalue()))
            st.image(image, caption="Gantt Chart", use_column_width=True)
        except Exception as e:
            st.error(f"⚠️ Could not display PNG: {e}")
    else:
        st.warning("⚠️ Please upload gantt_chart.png")

# %% 🔹 Pareto + Pie Chart
fig_pareto, fig_pie, fig_heatmap = None, None, None
with st.expander("🧁 Reason Analysis"):
    if "Reason" in df.columns and df["Reason"].notna().any():
        count_by_reason = df["Reason"].value_counts()
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("📊 Pareto Chart")
            
            # Tạo figure và trục
            fig_pareto, ax1 = plt.subplots(figsize=(10, 5))  # 👉 tăng ngang cho nhãn dài

            # Dữ liệu
            values = count_by_reason.values
            labels = [l.replace(" ", "\n") for l in count_by_reason.index]
            cumulative = np.cumsum(values) / sum(values) * 100

            # Vẽ cột
            ax1.bar(labels, values, color="#69b3a2", alpha=0.8)
            ax1.set_ylabel("Số lượng")

            # Vẽ đường cộng dồn
            ax2 = ax1.twinx()
            ax2.plot(labels, cumulative, color="red", marker="o", linewidth=2)
            ax2.set_ylabel("Tỷ lệ cộng dồn (%)")

            # Sửa nhãn trục x
            ax1.set_xticks(range(len(labels)))
            ax1.set_xticklabels(labels, rotation=40, ha="center", fontsize=10)

            # Căn layout cho không bị cắt nhãn
            fig_pareto.tight_layout()

            # Hiển thị
            st.pyplot(fig_pareto)


        with col2:
            st.subheader("🧁 Pie Chart")
            fig_pie, ax = plt.subplots(figsize=(6, 6))
            labels = [f"{r} ({v} jobs)" for r, v in zip(count_by_reason.index, count_by_reason.values)]
            explode = [0.1 if v == max(values) else 0 for v in values]

            # 🎨 Dùng bảng màu tab20
            colors = plt.cm.tab20(np.linspace(0, 1, len(values)))

            ax.pie(
                values,
                labels=labels,
                explode=explode,
                startangle=90,
                colors=colors,
                autopct='%1.1f%%',
                textprops={'fontsize': 10}
            )
            st.pyplot(fig_pie)


# %% 🔹 Heatmap
with st.expander("🧭 Heatmap by Machine × Date"):
    if {"Machine", "Date", "Delay (min)"}.issubset(df.columns):
        # Dùng mean để thể hiện xu hướng sớm/trễ
        heat_data = df.pivot_table(
            index="Machine",
            columns="Date",
            values="Delay (min)",
            aggfunc="mean"   # giữ logic âm/dương
        )

        if not heat_data.empty:
            fig_heatmap, ax = plt.subplots(figsize=(10, 6))
            sns.heatmap(
                heat_data,
                cmap="RdYlGn_r",   # đỏ = trễ, xanh = sớm, vàng = gần đúng giờ
                center=0,          # mốc 0 ở giữa
                annot=True,
                fmt=".0f",
                ax=ax
            )
            ax.set_title("Heatmap: Mean Delay (min) by Machine × Date", fontsize=14)
            st.pyplot(fig_heatmap)

        # 📌 Thêm bảng thống kê min / max / std để hiểu rõ hơn
        stats = df.pivot_table(
            index="Machine",
            columns="Date",
            values="Delay (min)",
            aggfunc=["min", "max", "std"]
        )
        st.dataframe(stats)
# %% 🔹 Scatterplot Start vs Duration Delay
# %% 🔹 Scatterplot Start vs Duration Delay
# with st.expander("📊 Scatterplot: Start Delay vs Duration Delay", expanded=True):
#     if {"Trễ bắt đầu", "Delay (min)"}.issubset(df.columns):

#         # Ưu tiên hue = Alert nếu có, ngược lại hue = Machine
#         hue_col = "Alert" if "Alert" in df.columns else "Machine"

#         fig_scatter, ax = plt.subplots(figsize=(8,6))
#         sns.scatterplot(
#             data=df,
#             x="Trễ bắt đầu",     # giữ nguyên vì chưa rename
#             y="Delay (min)",     # sau rename từ "Trễ thời gian"
#             hue=hue_col,
#             style="Machine" if "Machine" in df.columns else None,
#             ax=ax,
#             s=80
#         )

#         # Vẽ trục chuẩn (0,0)
#         ax.axhline(0, color="black", linewidth=1)
#         ax.axvline(0, color="black", linewidth=1)

#         ax.set_xlabel("Trễ bắt đầu (phút)")
#         ax.set_ylabel("Trễ thời gian (phút)")   # nhãn vẫn giữ tiếng Việt cho dễ đọc
#         ax.set_title("Start Delay vs Duration Delay", fontsize=14)

#         st.pyplot(fig_scatter)


# %% 🔹 Export Dashboard to HTML
report_date = df["Date"].dropna().mode()[0] if "Date" in df.columns else "N/A"
report_date_str = pd.to_datetime(report_date).strftime("%d/%m/%Y")
logo_base64 = img_to_base64("TKMB.jpg")   # vì file nằm chung folder
logo_html = f'<img src="data:image/jpg;base64,{logo_base64}" class="company-logo">'

# %% 🔹 Export Dashboard to HTML
if st.button("📥 Export FULL Dashboard to HTML"):
    html_template = """<!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <title>Variance Dashboard</title>
    <style>
        body {{
            font-family: "Source Sans Pro", sans-serif;
            margin: 2rem;
            background: #f0f2f6;
            color: #333;
        }}

        h1, h2, h3 {{
            color: #222;
        }}

        .block-box {{
            border: none;  /* bỏ border mảnh */
            border-radius: 12px;
            background: #fff;
            padding: 20px;
            margin: 24px 0;
            box-shadow: 0 8px 20px rgba(0,0,0,0.12); /* 💥 shadow mạnh hơn */
            transition: all 0.3s ease-in-out;
        }}
        .metric-box {{
            display: inline-block;
            width: 22%;
            background: #fff;
            margin: 10px;
            padding: 14px;
            border-radius: 10px;
            border: none;
            box-shadow: 0 10px 28px rgba(0,0,0,0.20);  /* 💥 shadow dày hơn */
            transition: all 0.3s ease-in-out;
        }}
        
        .metric-box:hover {{
            transform: translateY(-4px);
            bbox-shadow: 0 10px 28px rgba(0,0,0,0.20); /* hover nổi hẳn */
        }}
        
        .company-logo {{
            position: absolute;
            top: 20px;
            right: 30px;
            width: 15%;
            height: auto;
            z-index: 1000;
        }}

        .metric-box h3 {{
            font-size: 14px;
            margin-bottom: 4px;
        }}

        .metric-box p {{
            font-size: 20px;
            font-weight: 600;
            color: #3f8efc;
            margin: 0;
        }}

        img.chart {{
            max-width: 100%;
            margin: 20px auto;
            display: block;
            border-radius: 6px;
            border: 1px solid #ccc;
        }}

        ul {{
            line-height: 1.6;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 10px auto;
            font-size: 14px;
            background: white;
            text-align: center;
        }}

        th, td {{
            border: 1px solid #ccc;
            padding: 6px 10px;
            vertical-align: middle;
        }}

        th {{
            background-color: #e9ecef;
            font-weight: bold;
        }}

        .table-block {{
            border: 1px solid #aaa;
            border-radius: 8px;
            background: #fff;
            padding: 12px;
            margin-bottom: 24px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }}

        .block-box {{
            border: 1px solid #aaa;
            border-radius: 8px;
            background: #fff;
            padding: 16px;
            margin: 20px 0;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }}

@media print {{
    body {{
        background: white !important;
        color: black !important;
        margin: 0 auto !important;     /* ✅ auto center */
        padding: 0.5cm !important;     /* ✅ padding đồng đều 4 phía */
        width: 95% !important;         /* ✅ co nhẹ để chừa đều trái–phải */
        max-width: 95% !important;
        overflow: visible !important;
    }}

    /* KPI block */
    .metric-box {{
        display: inline-block;
        width: 30% !important;      
        margin: 6px !important;
        padding: 10px !important;
        font-size: 14px !important;
        border: 1px solid #000 !important;
        box-shadow: none !important;
        vertical-align: top;
        box-sizing: border-box !important;
    }}

    .metric-box p {{
        font-size: 18px !important;
        font-weight: bold;
    }}

    /* Section box */
    .block-box, .table-block {{
        width: 100% !important;
        box-sizing: border-box !important;
        padding: 14px !important;
        margin: 15px 0 !important;
        border: 1px solid #000 !important;
        background: #fff !important;
        box-shadow: none !important;
        page-break-inside: avoid;
        break-inside: avoid;
    }}

    /* Table */
    table {{
        font-size: 13px !important;
        width: 100% !important;
        border-collapse: collapse;
    }}

    th, td {{
        border: 1px solid #000 !important;
        padding: 6px 8px !important;
    }}

    /* Chart */
    img.chart {{
        width: 100% !important;     
        max-height: 80vh !important;   /* ✅ ép vừa trang, không bị nhảy */
        object-fit: contain !important;
        margin: 10px auto !important;
        display: block;
        page-break-inside: avoid;
        break-inside: avoid;
    }}

    /* Heatmap riêng → luôn xuống trang mới */
    .heatmap-block {{
        page-break-before: always;
        page-break-inside: avoid;
    }}

    /* Trang giấy */
    @page {{
        size: A4 landscape;
        margin: 0.5cm;
    }}
}}



            h1, h2, h3 {{
                font-size: 22px !important;
                margin-top: 10px !important;
                margin-bottom: 8px !important;
                page-break-after: avoid;
                break-after: avoid;
            }}

            .company-logo {{
                width: 250px !important;
            }}
        }}

    </style>
    </head>
    <body>
    {logo_html}
    <h1>📊 Variance Report – Plan vs Actual</h1>
    <h3>📅 Report Date: {report_date_str}</h3>

    <div class="block-box">
    <h2>📈 Overview</h2>
    <div class='metric-box'><h3>🧾 Total Jobs</h3><p>{total_jobs}</p></div>
    <div class='metric-box'><h3>🚨 Jobs with Alert</h3><p>{num_alerts}</p></div>
    <div class='metric-box'><h3>⚠️ Alert Rate</h3><p>{alert_rate:.1f}%</p></div>
    <div class='metric-box'><h3>🏭 Active Machines</h3><p>{total_machines}</p></div>
    <div class='metric-box'><h3>💥 Machines with Alert</h3><p>{alert_machines}</p></div>
    <div class='metric-box'><h3>📊 Machine Alert Rate</h3><p>{machine_alert_rate:.1f}%</p></div>
    </div>

    {details_html}

    {alert_details_html}

    <div class="block-box">
    <h2>📉 Gantt Chart</h2>
    {gantt_html}
    </div>

    <div class="block-box">
    <h2>📊 Pareto Chart</h2>
    {pareto_html}
    </div>

    <div class="block-box">
    <h2>🧁 Pie Chart</h2>
    {pie_html}
    </div>

    <div class="block-box">
    <h2>🌡 Heatmap</h2>
    {heatmap_html}
    </div>

    </body>
    </html>
    """
    

    # Convert ảnh và figure
    gantt_html = file_to_base64_img(uploaded_png)
    pareto_html = fig_to_base64_img(fig_pareto) if fig_pareto else ""
    pie_html = fig_to_base64_img(fig_pie) if fig_pie else ""
    heatmap_html = fig_to_base64_img(fig_heatmap) if fig_heatmap else ""

    # KPI
    total_jobs = int(total_jobs) if pd.notnull(total_jobs) else 0
    num_alerts = int(num_alerts) if pd.notnull(num_alerts) else 0
    alert_rate = float(alert_rate) if pd.notnull(alert_rate) else 0
    total_machines = int(df["Machine"].nunique()) if "Machine" in df.columns else 0
    alert_machines = int(df.loc[df["Alert"].notna(), "Machine"].nunique()) if "Alert" in df.columns and "Machine" in df.columns else 0
    machine_alert_rate = 100 * alert_machines / total_machines if total_machines else 0

    # =============== DETAILS BLOCK ===============
    details_html = ""
    if "Delay (min)" in df.columns and not df.empty:
        worst_job = df.loc[df["Delay (min)"].idxmax()]
        details_html += f"""
        <div class='block-box'>
        <h2>🛑 Job with Max Delay</h2>
        <ul>
          <li>⏰ <b>{worst_job['Delay (min)']} minutes</b></li>
          <li>🏭 Machine: <b>{worst_job.get('Machine', 'N/A')}</b></li>
          <li>🧾 ItemCode: <b>{worst_job.get('ItemCode', 'N/A')}</b></li>
          <li>📆 Date: <b>{worst_job.get('Date', 'N/A')}</b></li>
        </ul>
        </div>
        """

    delay_100 = df[df["Delay (min)"] > 100].shape[0]
    delay_200 = df[df["Delay (min)"] > 200].shape[0]
    details_html += f"""
    <div class='block-box'>
    <h2>📊 Jobs exceeding delay thresholds</h2>
    <ul>
      <li>🟥 Delay > 200 min: <b>{delay_200}</b></li>
      <li>🟧 Delay > 100 min: <b>{delay_100}</b></li>
    </ul>
    </div>
    """

    if "Reason" in df.columns:
        top_reasons = df[df["Reason"].notna()]["Reason"].value_counts().head(3)
        details_html += "<div class='block-box'><h2>🧠 Top 3 Reasons</h2><ul>"
        for reason, count in top_reasons.items():
            details_html += f"<li>{reason}: <b>{count} jobs</b></li>"
        details_html += "</ul></div>"

    if "Delay (min)" in df.columns and "Machine" in df.columns:
        delay_by_machine = (
            df[df["Delay (min)"] > 0]
            .groupby("Machine")["Delay (min)"]
            .sum()
            .sort_values(ascending=False)
            .head(3)
        )
        details_html += "<div class='block-box'><h2>🏭 Top 3 Machines by Total Delay</h2><ul>"
        for mach, mins in delay_by_machine.items():
            details_html += f"<li>{mach}: <b>{int(mins)} min</b></li>"
        details_html += "</ul></div>"

    # =============== ALERT DETAILS BLOCK ===============
    alert_df = df[df["Alert"].notna()] if "Alert" in df.columns else pd.DataFrame()
    show_cols = [col for col in ["Date", "Machine", "ItemCode", "Delay (min)", "Reason", "Note"] if col in df.columns]
    alert_details_html = "<div class='block-box'><h2>📋 Jobs with Alerts</h2>"

    if not alert_df.empty and "Reason" in df.columns:
        alert_details_html += "<h3>🧠 Breakdown by Top 3 Reasons:</h3>"
        for reason in top_reasons.index:
            filtered = alert_df[alert_df["Reason"] == reason]
            alert_details_html += f"<h4>🔸 {reason} – {len(filtered)} jobs:</h4>"
            alert_details_html += filtered[show_cols].to_html(index=False, border=1, justify="center")

    if not alert_df.empty and "Machine" in df.columns:
        alert_details_html += "<h3>🏭 Breakdown by Top 3 Machines:</h3>"
        for mach in delay_by_machine.index:
            filtered = alert_df[alert_df["Machine"] == mach]
            alert_details_html += f"<h4>🔸 {mach} – {len(filtered)} jobs:</h4>"
            alert_details_html += filtered[show_cols].to_html(index=False, border=1, justify="center")

    alert_details_html += "</div>"

    # =============== EXPORT ====================
     # Fill data
    filled_html = html_template.format(
        report_date_str=report_date_str,
        total_jobs=total_jobs,
        num_alerts=num_alerts,
        alert_rate=alert_rate,
        total_machines=total_machines,
        alert_machines=alert_machines,
        machine_alert_rate=machine_alert_rate,
        details_html=details_html,
        alert_details_html=alert_details_html,
        gantt_html=gantt_html,
        pareto_html=pareto_html,
        pie_html=pie_html,
        heatmap_html=heatmap_html,
        logo_html=logo_html
    )


    with open("dashboard_snapshot.html", "w", encoding="utf-8") as f:
        f.write(filled_html)

    with open("dashboard_snapshot.html", "rb") as f:
        st.download_button("📥 Download HTML File", f, file_name="dashboard_snapshot.html", mime="text/html")


