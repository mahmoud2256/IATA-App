"""
IATA BSP Converter PRO - Fixed Streamlit Edition
Developed by Mahmoud Amin
No fitz / PyMuPDF dependency - uses only pdfplumber + openpyxl
"""
import streamlit as st
import io, re, os
import pandas as pd

st.set_page_config(
    page_title="IATA BSP Converter PRO | Mahmoud Amin",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    header {visibility: hidden;}
    .stApp { background-color: #0F172A; }
    .header-box {
        background: linear-gradient(135deg, #1E3A8A 0%, #1E40AF 50%, #3B82F6 100%);
        padding: 40px 30px; border-radius: 20px; text-align: center;
        margin-bottom: 30px; box-shadow: 0 20px 40px rgba(59,130,246,0.3);
    }
    .header-title { font-size: 42px; font-weight: 900; color: white; }
    .header-sub { font-size: 14px; color: #BFDBFE; margin-top: 8px; letter-spacing: 3px; }
    .header-dev { font-size: 13px; color: #93C5FD; margin-top: 10px; font-style: italic; }
    .card { background: #1E293B; padding: 25px; border-radius: 15px; border: 1px solid #334155; }
    .stButton>button { background: linear-gradient(135deg, #10B981, #059669); color: white; font-weight: 900; font-size: 18px; padding: 18px; border-radius: 12px; border: none; width: 100%; }
    .footer { text-align: center; color: #64748B; padding: 30px; font-size: 12px; margin-top: 40px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-box">
    <div class="header-title">✈️ IATA / BSP CONVERTER</div>
    <div class="header-sub">FCAGBILLDET PDF → EXCEL • PROFESSIONAL EDITION v3.1 FIXED</div>
    <div class="header-dev">Developed by Mahmoud Amin • 2026 • Streamlit Cloud Ready</div>
</div>
""", unsafe_allow_html=True)

# Imports with safe check
try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except Exception as e:
    HAS_PDFPLUMBER = False
    st.error(f"pdfplumber not installed: {e}")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    HAS_OPENPYXL = True
except Exception as e:
    HAS_OPENPYXL = False
    st.error(f"openpyxl not installed: {e}")

def parse_pdf_simple(pdf_bytes):
    """Simple robust parser using only pdfplumber - no fitz"""
    tickets = []
    full_text = ""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                try:
                    txt = page.extract_text() or ""
                    full_text += txt + "\n"
                    # Also try tables
                    tables = page.extract_tables() or []
                    for tbl in tables:
                        for row in tbl:
                            if row:
                                line = " | ".join([str(c or "") for c in row])
                                # Ticket pattern
                                m = re.search(r'(\d{3})[-\s]?(\d{10,13})', line)
                                if m:
                                    tickets.append({
                                        "doc_no": f"{m.group(1)}-{m.group(2)}",
                                        "raw": line[:250],
                                        "source": "table"
                                    })
                except:
                    continue

        # If no tickets from tables, scan full text lines
        if len(tickets) == 0:
            for line in full_text.split("\n"):
                m = re.search(r'(\d{3})[-\s]?(\d{10,13})', line)
                if m:
                    # Filter out lines that are too short or look like dates
                    if len(line.strip()) > 15:
                        tickets.append({
                            "doc_no": f"{m.group(1)}-{m.group(2)}",
                            "raw": line.strip()[:250],
                            "source": "text"
                        })
        
        # Fallback: if still no tickets, create raw lines for debugging
        if len(tickets) == 0 and len(full_text) > 100:
            # Split into meaningful lines
            lines = [l.strip() for l in full_text.split("\n") if len(l.strip()) > 20][:200]
            for l in lines:
                tickets.append({
                    "doc_no": "EXTRACTED",
                    "raw": l[:250],
                    "source": "raw_text"
                })
                
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
        import traceback
        st.code(traceback.format_exc())
    
    return tickets, full_text

def create_excel_file(tickets, full_text, filename):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "BSP-ISSUES"
    
    # Styles
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    
    headers = ["#", "Document No / Ticket", "Extracted Line", "Source"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    
    for i, t in enumerate(tickets, 1):
        ws.append([i, t.get("doc_no",""), t.get("raw",""), t.get("source","")])
    
    # Auto width
    try:
        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 22
        ws.column_dimensions['C'].width = 90
        ws.column_dimensions['D'].width = 12
    except:
        pass
    
    # Summary sheet
    ws2 = wb.create_sheet("FULL_TEXT")
    ws2.append(["Full extracted text from PDF (first 30000 chars)"])
    ws2.append([full_text[:30000]])
    
    ws3 = wb.create_sheet("INFO")
    ws3.append(["Field", "Value"])
    ws3.append(["Developed by", "Mahmoud Amin"])
    ws3.append(["Original PDF", filename])
    ws3.append(["Tickets/Lines Found", len(tickets)])
    ws3.append(["Version", "v3.1 Fixed - No fitz dependency"])
    
    wb.save(output)
    output.seek(0)
    return output

# Layout
col1, col2 = st.columns([2, 1], gap="large")

with col1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<p style="color:#3B82F6; font-weight:800; font-size:12px; letter-spacing:2px;">STEP 1</p><h3 style="color:white; margin-top:0;">📄 Select PDF File</h3>', unsafe_allow_html=True)
    uploaded = st.file_uploader("Upload EG_FCAGBILLDET PDF", type=["pdf"], label_visibility="collapsed")
    
    st.markdown('<div style="height:20px"></div><p style="color:#10B981; font-weight:800; font-size:12px; letter-spacing:2px;">STEP 2</p><h3 style="color:white; margin-top:0;">⚙️ Convert & Download</h3>', unsafe_allow_html=True)
    
    if uploaded:
        st.success(f"✓ Selected: {uploaded.name} ({uploaded.size//1024} KB)")
        pdf_bytes = uploaded.getvalue()
        
        with st.spinner("Parsing PDF with pdfplumber..."):
            tickets, full_text = parse_pdf_simple(pdf_bytes)
        
        if tickets:
            st.metric("Lines / Tickets Found", len(tickets))
            df = pd.DataFrame(tickets[:20])
            st.dataframe(df, use_container_width=True)
            
            excel_data = create_excel_file(tickets, full_text, uploaded.name)
            
            st.download_button(
                "⚡ GENERATE & DOWNLOAD EXCEL NOW",
                data=excel_data,
                file_name=uploaded.name.replace(".pdf","").replace(".PDF","") + "_CONVERTED.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            st.balloons()
            st.success(f"✅ Done! Developed by Mahmoud Amin - {len(tickets)} records converted")
        else:
            st.warning("No data extracted - PDF might be scanned image or empty")
            if full_text:
                st.text_area("Extracted text", full_text[:2000], height=200)
    else:
        st.info("👆 Upload your EG_FCAGBILLDET_*.PDF file to start")
    
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<p style="color:#3B82F6; font-weight:800;">FILE INFO</p><h4 style="color:white;">Fixed Version v3.1</h4>', unsafe_allow_html=True)
    st.markdown("""
    <div style="background:#334155; padding:15px; border-radius:10px; color:#F1F5F9; font-size:13px; line-height:1.6;">
    <b>✅ What's Fixed:</b><br>
    • Removed fitz/PyMuPDF dependency<br>
    • Uses only pdfplumber<br>
    • Works on Streamlit Cloud<br>
    • No NameError<br><br>
    <b>Requirements.txt:</b><br>
    streamlit<br>pdfplumber<br>openpyxl<br>pandas<br><br>
    <b>GitHub:</b> Put both files in root:<br>
    - streamlit_app.py<br>
    - requirements.txt (without PyMuPDF)
    </div>
    """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="footer">© 2026 Mahmoud Amin - IATA BSP Converter PRO v3.1 Fixed - Developed for Travel Agencies</div>', unsafe_allow_html=True)
