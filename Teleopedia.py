import os
from dotenv import load_dotenv
import streamlit as st
import pandas as pd
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import formataddr
import datetime
from io import BytesIO
import time
import base64
import streamlit.components.v1 as components

# Load environment variables from .env file
load_dotenv()

# =============================================
# ============ CONFIGURATION SECTION ==========
# =============================================
# 🔧 UPDATE THESE VALUES FOR YOUR COMPANY

# --- Company Branding ---
COMPANY_NAME = "Teleopedia Communications Pvt. Ltd"
COMPANY_TAGLINE = "Communication Easily"
LOGO_URL = "https://teleopedia.com/wp-content/uploads/2022/10/logo-dark.png"
PRIMARY_COLOR = "#004080"
SECONDARY_COLOR = "#008080"

# --- Contact Information ---
CONTACT_PHONE = "+91-7013785049"
CONTACT_EMAIL = "nadeem@teleopedia.com"
COMPANY_WEBSITE = "https://teleopedia.com/"

# --- Email Server Settings ---
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587

# --- Email Content ---
EMAIL_SUBJECT = "Business Proposal for Strategic Collaboration with Teleopedia Communications"

# ← Customize email body template (HTML format supported) ↓
EMAIL_BODY_TEMPLATE = """
<img src='{LOGO_URL}' style='width:120px;'><br><br>

Dear <b>{name}</b>,<br><br>

    I’m reaching out on behalf of <b>Teleopedia Communications Pvt. Ltd</b>, a trusted provider of next-gen enterprise communication and digital marketing solutions.<br><br>

    In today’s fast-paced digital world, having <b>scalable, efficient, and reliable communication systems</b> is critical for customer engagement and growth. At <b>Teleopedia Communications</b>, we specialize in <b>SMS, Voice, WhatsApp, Email, and RCS services</b> tailored to businesses like yours.<br><br>

    We would love the opportunity to work with your team and enhance your business communication strategy. Here’s what we offer:<br><br>

    <div style="line-height: 1.8;">
      ✅ <b>Bulk SMS & Transactional Alerts</b><br>
      ✅ <b>WhatsApp for Business API Integration</b><br>
      ✅ <b>Automated Voice Solutions (IVR, Blasts)</b><br>
      ✅ <b>Email Marketing & Campaign Automation</b><br>
      ✅ <b>RCS Messaging with Interactive Capabilities</b><br>
      ✅ <b>High-Speed SMPP Connectivity</b><br>
    </div><br>

    Our platform supports <b>30M+ messages per month</b> and serves clients across <b>30+ cities in India</b>. We offer <b>24/7 support, real-time analytics</b>, and custom integrations to help you get the best out of every customer interaction.<br><br>

    I would be happy to schedule a brief call to discuss how <b>Teleopedia Communications</b> can support your goals and deliver measurable impact.<br><br>

    <img src='{LOGO_URL}' width='1' height='1' style='display:none;'><br>

    Regards,<br>
    <b>Nadeem Ahmad || Teleopedia Communications</b> 📲<br>
    📞 <b>+91 7013785049</b> | 📧 <b>nadeem@teleopedia.com</b> | 🌐 <a href='https://teleopedia.com/'><b>teleopedia.com</b></a><br><br>

    <hr style="border: none; border-top: 1px solid #ddd;">
    <div style="font-size: 12px; color: #999; text-align: center;">
      © {CURRENT_YEAR} Teleopedia Communications | <a href="https://teleopedia.com/" style="color: #999;">https://teleopedia.com/</a>
"""

# --- Footer ---
FOOTER_TEXT = f"© {datetime.datetime.now().year} {COMPANY_NAME}"
FOOTER_LINK = "https://teleopedia.com/"

# =============================================
# ======= END OF CONFIGURATION SECTION ========
# =============================================

# --- Page Setup ---
st.set_page_config(
    page_title=f"{COMPANY_NAME} Email Tool",
    page_icon="📧",
    layout="centered"
)

# --- Custom CSS ---
st.markdown(f"""
    <style>
    .main {{
        background-color: #f5f9ff;
    }}
    .stButton > button {{
        background-color: {PRIMARY_COLOR};
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 10px 20px;
        transition: all 0.3s;
    }}
    .stButton > button:hover {{
        background-color: {PRIMARY_COLOR}CC;
        transform: scale(1.02);
    }}
    .stDownloadButton > button {{
        background-color: {SECONDARY_COLOR};
        color: white;
        font-weight: bold;
        border-radius: 6px;
        padding: 8px 16px;
        transition: all 0.3s;
    }}
    .stDownloadButton > button:hover {{
        background-color: {SECONDARY_COLOR}CC;
        transform: scale(1.02);
    }}
    .progress-container {{
        background-color: #e0e0e0;
        border-radius: 10px;
        height: 20px;
        margin: 15px 0;
    }}
    .progress-bar {{
        background: linear-gradient(90deg, #4CAF50, #8BC34A);
        border-radius: 10px;
        height: 100%;
        transition: width 0.5s;
    }}
    .success-box {{
        background-color: #e8f5e9;
        border-left: 5px solid #4CAF50;
        padding: 15px;
        border-radius: 5px;
        margin: 15px 0;
    }}
    .metric-box {{
        background-color: #e3f2fd;
        border-radius: 10px;
        padding: 15px;
        margin: 15px 0;
        text-align: center;
    }}
    .status-badge {{
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 12px;
        font-weight: bold;
    }}
    .status-sent {{
        background-color: #e8f5e9;
        color: #2e7d32;
    }}
    .status-failed {{
        background-color: #ffebee;
        color: #c62828;
    }}
    </style>
""", unsafe_allow_html=True)

# --- Header ---
st.markdown(f"""
    <div style='text-align: center;'>
        <img src="{LOGO_URL}" width="260">
    </div>
    <h2 style='text-align: center; color: {PRIMARY_COLOR};'>{COMPANY_NAME} - Bulk Email Tool</h2>
    <p style='text-align: center; color: #666;'>{COMPANY_TAGLINE}</p>
    <hr>
""", unsafe_allow_html=True)

# --- File Uploads ---
uploaded_file = st.file_uploader("📄 Upload Excel with 'Name' and 'Email' columns", type=["xlsx"])
attachment = st.file_uploader("📎 Optional Attachment")

delivery_report = []

# --- Recipient Preview ---
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        if 'Email' not in df.columns or 'Name' not in df.columns:
            st.error("❌ Excel file must contain 'Name' and 'Email' columns.")
        else:
            st.success(f"✅ Loaded {len(df)} recipients.")
            st.dataframe(df[['Name', 'Email']].head(5))
    except Exception as e:
        st.error(f"❌ Error reading Excel file: {str(e)}")

# --- Email Sending Functionality ---
if st.button("🚀 Send Emails Now", key="send_emails"):
    if not uploaded_file:
        st.error("Please upload a valid Excel file.")
    else:
        # Validate environment variables
        if not EMAIL_USER or not EMAIL_PASSWORD:
            st.error("❌ Email Error: Missing EMAIL_USER or EMAIL_PASSWORD. Please check your .env file or Streamlit Cloud secrets.")
        else:
            try:
                st.info("🔌 Connecting to email server...")
                context = ssl.create_default_context()
                server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
                server.starttls(context=context)  # Enable TLS for Gmail
                server.login(EMAIL_USER, EMAIL_PASSWORD)
                st.success("✅ Connected to email server successfully!")

                progress_bar = st.empty()
                status_text = st.empty()
                progress_container = st.empty()

                total_emails = len(df)
                success_count = 0

                for i, (_, row) in enumerate(df.iterrows()):
                    # Update progress
                    progress = (i + 1) / total_emails
                    progress_bar.markdown(f"""
                    <div class="progress-container">
                        <div class="progress-bar" style="width: {progress * 100}%"></div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    status_text.info(f"📨 Sending email {i+1} of {total_emails}...")

                    recipient = row['Email']
                    name = row['Name']
                    personalized_body = EMAIL_BODY_TEMPLATE.format(
                        name=name,
                        LOGO_URL=LOGO_URL,
                        CURRENT_YEAR=datetime.datetime.now().year
                    )

                    msg = MIMEMultipart()
                    msg['From'] = formataddr((COMPANY_NAME, EMAIL_USER))
                    msg['To'] = recipient
                    msg['Subject'] = EMAIL_SUBJECT

                    msg.attach(MIMEText(personalized_body, 'html'))

                    if attachment is not None:
                        attachment.seek(0)
                        file_data = attachment.read()
                        file_name = attachment.name
                        part = MIMEApplication(file_data, Name=file_name)
                        part['Content-Disposition'] = f'attachment; filename="{file_name}"'
                        msg.attach(part)

                    try:
                        server.sendmail(EMAIL_USER, recipient, msg.as_string())
                        delivery_report.append({
                            "Name": name,
                            "Email": recipient,
                            "Status": "✅ Sent",
                            "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        })
                        success_count += 1
                    except Exception as e:
                        delivery_report.append({
                            "Name": name,
                            "Email": recipient,
                            "Status": f"❌ Failed: {str(e)[:50]}",
                            "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        })

                server.quit()
                
                # Clear progress elements
                progress_bar.empty()
                status_text.empty()
                progress_container.empty()
                
                # Celebration!
                st.balloons()
                st.success(f"🎉 Successfully processed {total_emails} emails!")
                
                # Success metrics
                success_rate = (success_count / total_emails) * 100
                st.markdown(f"""
                <div class="metric-box">
                    <h3>Delivery Metrics</h3>
                    <p style="font-size: 24px; margin: 5px 0;">Success Rate: <b>{success_rate:.1f}%</b></p>
                    <p style="font-size: 16px; margin: 5px 0;">✅ Successful: {success_count} | ❌ Failed: {total_emails - success_count}</p>
                </div>
                """, unsafe_allow_html=True)

                # Generate and auto-download report
                report_df = pd.DataFrame(delivery_report)
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                excel_filename = f"{COMPANY_NAME.replace(' ', '_')}_report_{timestamp}.xlsx"
                excel_buffer = BytesIO()
                report_df.to_excel(excel_buffer, index=False, engine='openpyxl')
                excel_buffer.seek(0)
                
                # Create download link
                b64 = base64.b64encode(excel_buffer.getvalue()).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{excel_filename}" id="auto-download">Download Report</a>'
                st.markdown(href, unsafe_allow_html=True)
                
                # JavaScript to trigger download
                components.html("""
                <script>
                function triggerDownload() {
                    var link = document.getElementById('auto-download');
                    if (link) {
                        link.click();
                    }
                }
                setTimeout(triggerDownload, 1000);
                </script>
                """, height=0)
                
                # Manual download button as fallback
                st.download_button(
                    label="💾 Click here if download didn't start automatically",
                    data=excel_buffer,
                    file_name=excel_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Email Error: {str(e)}")

# --- Delivery Report Section ---
if delivery_report:
    st.subheader("📊 Delivery Analytics")
    report_df = pd.DataFrame(delivery_report)
    
    # Format status badges
    def format_status(status):
        if "Sent" in status:
            return f"<span class='status-badge status-sent'>{status}</span>"
        else:
            return f"<span class='status-badge status-failed'>{status}</span>"
    
    report_df['Status'] = report_df['Status'].apply(format_status)

    # Filter options
    col1, col2 = st.columns(2)
    with col1:
        status_filter = st.selectbox("Filter by Status", options=["All"] + ["✅ Sent", "❌ Failed"])
    with col2:
        show_count = st.slider("Show Records", 5, 50, 10)

    if status_filter != "All":
        filtered_df = report_df[report_df['Status'].str.contains(status_filter)]
    else:
        filtered_df = report_df

    st.markdown(filtered_df.head(show_count).to_html(escape=False), unsafe_allow_html=True)

# --- Footer ---
st.markdown(f"""
---
<p style='text-align: center; font-size: 12px;'>
{FOOTER_TEXT} | <a href="{FOOTER_LINK}">Visit Website</a>
</p>
""", unsafe_allow_html=True)