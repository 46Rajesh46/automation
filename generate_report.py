"""
FIX 5: Plotly HTML report generator + Outlook email sender.

Reads TO / CC email addresses from control.xlsx Main sheet (same config section
as URL / username / password) so you never need to edit this file.

Add these rows to control.xlsx → Main sheet (key/value columns):
  Email To  |  manager@hospital.com;doctor@hospital.com
  Email CC  |  admin@hospital.com

Requirements: pip install plotly xlrd openpyxl
"""
import os
import glob
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.io as pio
from datetime import datetime, timedelta
from win32com.client import dynamic

BASE_DIR   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Zonal Zodiac Automation", "Input Data")
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
CTRL_PATH  = os.path.join(OUTPUT_DIR, "control.xlsx")


def append_run_log(message):
    log_path = os.path.join(OUTPUT_DIR, f"run_log_{datetime.now().strftime('%Y-%m-%d')}.txt")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {message}\n")


def load_email_config():
    """Read Email To and Email CC from control.xlsx Main sheet config section."""
    try:
        raw = pd.read_excel(CTRL_PATH, sheet_name="Main", header=None)
        cfg = {}
        for i in range(len(raw)):
            key = str(raw.iloc[i, 0]).strip().lower()
            val = raw.iloc[i, 1]
            if pd.isna(val):
                val = ""
            cfg[key] = str(val).strip()
        return cfg.get("email to", cfg.get("email_to", "")), cfg.get("email cc", cfg.get("email_cc", ""))
    except Exception as e:
        append_run_log(f"[REPORT] Could not read email config: {e}")
        return "", ""


def read_excel_any(filepath):
    """Use xlrd for .xls (old format), openpyxl for .xlsx."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".xls":
        return pd.read_excel(filepath, sheet_name=0, engine="xlrd")
    return pd.read_excel(filepath, sheet_name=0, engine="openpyxl")


def get_unit(filename):
    return os.path.basename(filename).split("_")[0]


# ---------------------------------------------------------------------------
# Revenue Service Level — use "Net Amount" column
# ---------------------------------------------------------------------------
def build_revenue_chart(fig, row, col):
    folder = os.path.join(BASE_DIR, "Revenue_Service_Level")
    files = glob.glob(os.path.join(folder, "*.xlsx")) + glob.glob(os.path.join(folder, "*.xls"))
    if not files:
        return

    frames = []
    for f in files:
        try:
            df = read_excel_any(f)
            # Prefer "Net Amount"; fall back to "Bill Amt" or first numeric column
            for col_try in ["Net Amount", "Bill Amt", "Itm Amt"]:
                if col_try in df.columns:
                    df = df[["_unit_placeholder", col_try]].rename(columns={col_try: "Value"}) if False else df
                    frames.append(pd.DataFrame({
                        "_unit": get_unit(f),
                        "Value": pd.to_numeric(df[col_try], errors="coerce").fillna(0)
                    }))
                    break
        except Exception as e:
            print(f"[WARN] Revenue file {os.path.basename(f)}: {e}")

    if not frames:
        return
    combined = pd.concat(frames, ignore_index=True)
    s = combined.groupby("_unit")["Value"].sum().reset_index().sort_values("Value", ascending=False)

    fig.add_trace(go.Bar(
        x=s["_unit"], y=s["Value"], name="Revenue",
        marker_color="#1565C0",
        text=s["Value"].apply(lambda v: f"₹{v/100000:.1f}L"),
        textposition="outside",
    ), row=row, col=col)
    fig.update_yaxes(title_text="Net Amount (₹)", row=row, col=col)
    fig.update_xaxes(title_text="Unit", row=row, col=col)


# ---------------------------------------------------------------------------
# OPD Foot Falls — count rows per unit (each row = 1 patient visit)
# also show total Item amount as secondary info
# ---------------------------------------------------------------------------
def build_opd_chart(fig, row, col):
    folder = os.path.join(BASE_DIR, "OPD_Foot_Falls_New")
    files = glob.glob(os.path.join(folder, "*.xlsx")) + glob.glob(os.path.join(folder, "*.xls"))
    if not files:
        return

    rows_list = []
    for f in files:
        try:
            df = read_excel_any(f)
            unit = get_unit(f)
            # Each row = 1 patient visit; use Qty if present, else count rows
            if "Qty" in df.columns:
                count = pd.to_numeric(df["Qty"], errors="coerce").fillna(1).sum()
            else:
                count = len(df)
            rows_list.append({"_unit": unit, "Count": count})
        except Exception as e:
            print(f"[WARN] OPD file {os.path.basename(f)}: {e}")

    if not rows_list:
        return
    s = pd.DataFrame(rows_list).groupby("_unit")["Count"].sum().reset_index().sort_values("Count", ascending=False)

    fig.add_trace(go.Bar(
        x=s["_unit"], y=s["Count"], name="OPD Footfall",
        marker_color="#2E7D32",
        text=s["Count"].apply(lambda v: f"{int(v):,}"),
        textposition="outside",
    ), row=row, col=col)
    fig.update_yaxes(title_text="Patient Visits", row=row, col=col)
    fig.update_xaxes(title_text="Unit", row=row, col=col)


# ---------------------------------------------------------------------------
# Census Report — malformed headers; columns are: SNO | Date | Ward | Count
# Group by unit (from filename), sum bed count column (index 3)
# ---------------------------------------------------------------------------
def build_census_chart(fig, row, col):
    folder = os.path.join(BASE_DIR, "Census_Report")
    files = glob.glob(os.path.join(folder, "*.xlsx")) + glob.glob(os.path.join(folder, "*.xls"))
    if not files:
        return

    rows_list = []
    for f in files:
        try:
            # header=None because the first row contains data, not proper headers
            df = pd.read_excel(f, sheet_name=0, header=None, engine="openpyxl"
                               if f.endswith(".xlsx") else None)
            if f.endswith(".xls"):
                df = pd.read_excel(f, sheet_name=0, header=None, engine="xlrd")
            # Column layout: 0=SNO, 1=DateSerial, 2=WardName, 3=BedCount
            if df.shape[1] >= 4:
                count = pd.to_numeric(df.iloc[:, 3], errors="coerce").fillna(0).sum()
            elif df.shape[1] >= 2:
                count = pd.to_numeric(df.iloc[:, -1], errors="coerce").fillna(0).sum()
            else:
                count = 0
            rows_list.append({"_unit": get_unit(f), "Count": count})
        except Exception as e:
            print(f"[WARN] Census file {os.path.basename(f)}: {e}")

    if not rows_list:
        return
    s = pd.DataFrame(rows_list).groupby("_unit")["Count"].sum().reset_index()
    s = s[s["Count"] > 0]

    fig.add_trace(go.Pie(
        labels=s["_unit"], values=s["Count"], name="Census",
        hole=0.45,
        textinfo="label+value",
        texttemplate="%{label}<br>%{value:.0f} beds",
        marker=dict(colors=["#1E88E5","#43A047","#FB8C00","#E53935",
                             "#8E24AA","#00ACC1","#F4511E","#6D4C41"]),
    ), row=row, col=col)


# ---------------------------------------------------------------------------
# Sales Day Book — use "Net Amount Sales W/Tax" or "Gross Amount"
# ---------------------------------------------------------------------------
def build_sales_chart(fig, row, col):
    folder = os.path.join(BASE_DIR, "Sales_Day_Book_Excel")
    files = glob.glob(os.path.join(folder, "*.xlsx")) + glob.glob(os.path.join(folder, "*.xls"))
    if not files:
        return

    frames = []
    for f in files:
        try:
            df = read_excel_any(f)
            for col_try in ["Net Amount Sales W/Tax", "Gross Amount", "Net Amount", "Amount"]:
                if col_try in df.columns:
                    frames.append(pd.DataFrame({
                        "_unit": get_unit(f),
                        "Value": pd.to_numeric(df[col_try], errors="coerce").fillna(0)
                    }))
                    break
        except Exception as e:
            print(f"[WARN] Sales file {os.path.basename(f)}: {e}")

    if not frames:
        return
    combined = pd.concat(frames, ignore_index=True)
    s = combined.groupby("_unit")["Value"].sum().reset_index().sort_values("Value", ascending=False)

    fig.add_trace(go.Bar(
        x=s["_unit"], y=s["Value"], name="Sales",
        marker_color="#E65100",
        text=s["Value"].apply(lambda v: f"₹{v/100000:.1f}L"),
        textposition="outside",
    ), row=row, col=col)
    fig.update_yaxes(title_text="Net Sales (₹)", row=row, col=col)
    fig.update_xaxes(title_text="Unit", row=row, col=col)


# ---------------------------------------------------------------------------
# Main: build and save HTML
# ---------------------------------------------------------------------------
def generate_html_report():
    report_date = (datetime.now() - timedelta(days=1)).strftime("%d-%b-%Y")
    generated_at = datetime.now().strftime("%d-%b-%Y %I:%M %p")
    title = f"Zonal Daily Dashboard — {report_date}"

    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=(
            "Revenue by Unit (Net Amount)",
            "OPD Footfall by Unit",
            "Census — Bed Occupancy by Unit",
            "Sales by Unit (Net Amount W/Tax)",
        ),
        specs=[[{"type": "bar"}, {"type": "bar"}],
               [{"type": "pie"}, {"type": "bar"}]],
        vertical_spacing=0.20,
        horizontal_spacing=0.12,
    )

    build_revenue_chart(fig, row=1, col=1)
    build_opd_chart(fig, row=1, col=2)
    build_census_chart(fig, row=2, col=1)
    build_sales_chart(fig, row=2, col=2)

    fig.update_layout(
        title=dict(
            text=f"{title}<br><sup style='color:gray'>Generated {generated_at}</sup>",
            font=dict(size=22, family="Segoe UI, Arial"),
            x=0.5,
        ),
        height=950,
        showlegend=False,
        paper_bgcolor="#F0F4F8",
        plot_bgcolor="#FFFFFF",
        font=dict(family="Segoe UI, Arial", size=12, color="#333333"),
        margin=dict(t=120, b=70, l=70, r=70),
    )

    # Uniform bar styling
    fig.update_traces(
        selector=dict(type="bar"),
        marker_line_width=0,
        opacity=0.9,
    )

    out_name = f"dashboard_report_{datetime.now().strftime('%Y-%m-%d')}.html"
    out_path = os.path.join(OUTPUT_DIR, out_name)
    # include_plotlyjs=True embeds JS — works offline, no internet needed
    pio.write_html(fig, file=out_path, full_html=True, include_plotlyjs=True)
    print(f"[OK] HTML report saved: {out_path}")
    append_run_log(f"[REPORT] HTML dashboard saved: {out_path}")
    return out_path, title


def send_report_email(report_path, subject, to_addr, cc_addr):
    if not to_addr:
        print("[WARN] 'Email To' not set in control.xlsx — skipping email.")
        print("       Add row:  Email To  |  recipient@hospital.com")
        append_run_log("[REPORT] Email skipped — 'Email To' not in control.xlsx.")
        return False

    report_date = (datetime.now() - timedelta(days=1)).strftime("%d-%b-%Y")
    body_html = (
        f"Dear All,<br><br>"
        f"Please find attached the <b>Zonal Daily Dashboard</b> for <b>{report_date}</b>.<br><br>"
        f"Open the attached <b>.html file</b> in any browser (Chrome, Edge, Firefox) "
        f"to view interactive charts. Hover over bars for exact values.<br><br>"
        f"Regards"
    )

    try:
        outlook = dynamic.DumbDispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = to_addr
        if cc_addr:
            mail.CC = cc_addr
        mail.Subject = subject
        mail.HTMLBody = body_html
        mail.Attachments.Add(os.path.abspath(report_path))
        mail.Send()
        print(f"[OK] Email sent to: {to_addr}")
        append_run_log(f"[REPORT] Email sent to={to_addr} cc={cc_addr}")
        return True
    except Exception as e:
        print(f"[WARN] Email send failed: {e}")
        append_run_log(f"[REPORT] Email failed: {e}")
        return False


def main():
    to_addr, cc_addr = load_email_config()
    report_path, subject = generate_html_report()
    send_report_email(report_path, subject, to_addr, cc_addr)


if __name__ == "__main__":
    main()
