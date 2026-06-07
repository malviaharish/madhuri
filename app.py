import streamlit as st
import os
import time
import urllib.parse
import uuid
import zipfile
import shutil
import datetime
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright
from playwright_stealth import stealth


# 1. Page Config (Must be first)
st.set_page_config(
    page_title="Capture Studio — Search & Export Tool",
    page_icon="📸",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Playwright Auto-installer (Cached)
try:
    import subprocess
    @st.cache_resource
    def install_playwright_browsers():
        subprocess.run(["playwright", "install", "chromium"], check=True)
    install_playwright_browsers()
except Exception as e:
    pass

# 3. Clean CSS styling (Only fonts and button gradients to ensure light/dark mode compatibility)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700&family=Inter:wght@400;500;600&display=swap');
    
    html, body, [class*="css"], .stApp {
        font-family: 'Inter', sans-serif;
    }
    h1, h2, h3, h4, h5, h6 {
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600 !important;
        letter-spacing: -0.02em !important;
    }
    
    /* Clean gradient for primary buttons */
    button[kind="primary"] {
        background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%) !important;
        border: none !important;
        border-radius: 8px !important;
        color: white !important;
        font-weight: 600 !important;
        padding: 10px 24px !important;
        box-shadow: 0 4px 12px rgba(99, 102, 241, 0.2) !important;
        transition: all 0.3s ease !important;
    }
    button[kind="primary"]:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 16px rgba(168, 85, 247, 0.35) !important;
    }
</style>
""", unsafe_allow_html=True)

# Initialize Session States
if "terms_input" not in st.session_state:
    st.session_state.terms_input = "triclosan suture\naspirin cardiovascular\nibuprofen inflammation"
if "logs" not in st.session_state:
    st.session_state.logs = []
if "outputs" not in st.session_state:
    st.session_state.outputs = []
if "job_id" not in st.session_state:
    st.session_state.job_id = None
if "is_running" not in st.session_state:
    st.session_state.is_running = False

# Workspace paths
WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
JOBS_DIR = os.path.join(WORKSPACE_DIR, "jobs")
os.makedirs(JOBS_DIR, exist_ok=True)

def append_log(message, log_type="info"):
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    emoji = "ℹ️" if log_type == "info" else "✅" if log_type == "success" else "⚠️" if log_type == "warning" else "❌"
    log_entry = f"[{timestamp}] {emoji} {message}"
    st.session_state.logs.append(log_entry)
    print(log_entry)

def clean_page_content(page, engine):
    selectors = [
        "header", "footer", "#search-sidebar", ".ncbi-header", ".m-footer", 
        "#gb", "#footer", ".Appheader", "#sfdiv", "#b_header", "#b_footer", ".b_footer",
        "[id*='cookie']", "[class*='cookie']", "[id*='consent']", "[class*='consent']",
        "[id*='banner']", "[class*='banner']", ".cookie-banner", "#onetrust-consent-sdk"
    ]
    selector_str = ", ".join(selectors)
    js_cleanup = f"""
    try {{
        const elementsToHide = document.querySelectorAll("{selector_str}");
        elementsToHide.forEach(el => {{ if (el) el.style.display = 'none'; }});
        const mainContent = document.querySelector('main');
        if (mainContent) {{
            mainContent.style.width = '100%';
            mainContent.style.margin = '0';
            mainContent.style.padding = '20px';
        }}
        const containers = document.querySelectorAll('.container, #main, #content');
        containers.forEach(c => {{ if (c) {{ c.style.maxWidth = '100%'; c.style.width = '100%'; }} }});
    }} catch (e) {{}}
    """
    try:
        page.evaluate(js_cleanup)
    except Exception:
        pass

def get_search_url(engine, query, custom_template=None):
    encoded = urllib.parse.quote(query)
    if engine == 'pmc':
        return f"https://pmc.ncbi.nlm.nih.gov/search/?term={encoded}"
    elif engine == 'pubmed':
        return f"https://pubmed.ncbi.nlm.nih.gov/?term={encoded}"
    elif engine == 'google':
        return f"https://www.google.com/search?q={encoded}"
    elif engine == 'scholar':
        return f"https://scholar.google.com/scholar?q={encoded}"
    elif engine == 'bing':
        return f"https://www.bing.com/search?q={encoded}"
    elif engine == 'duckduckgo':
        return f"https://duckduckgo.com/?q={encoded}"
    elif engine == 'custom' and custom_template:
        return custom_template.replace("{query}", encoded)
    return f"https://pmc.ncbi.nlm.nih.gov/search/?term={encoded}"


# Header Styling
st.markdown("""
<div style="background-color:rgba(124, 58, 237, 0.08); border-left: 5px solid #8b5cf6; padding: 20px; border-radius: 12px; margin-bottom: 25px;">
    <h2 style="margin: 0; color:#8b5cf6;">📸 Capture Studio</h2>
    <p style="margin: 5px 0 0 0; font-size: 0.95rem; opacity: 0.85;">
        Batch search query screenshot and PDF document exporter powered by Playwright.
    </p>
</div>
""", unsafe_allow_html=True)

# Main 2-Column Layout
col_left, col_right = st.columns([1, 2], gap="large")

# --- Left Column: Inputs & Settings ---
with col_left:
    st.subheader("⚙️ Settings & Inputs")
    
    with st.container(border=True):
        st.markdown("**1. Select Engine & Viewport**")
        engine = st.selectbox("Search Engine", ["pmc", "pubmed", "google", "scholar", "bing", "duckduckgo", "custom"])
        custom_url = ""
        if engine == "custom":
            custom_url = st.text_input("Custom URL (use {query} placeholder)", value="https://example.com/search?q={query}")
            
        viewport = st.selectbox("Viewport Size", ["1920x1080", "1280x720", "1536x864", "1024x768"])
        export_format = st.selectbox("Export Format", ["screenshot", "screenshot_full", "pdf", "both"])
        
        with st.expander("Advanced Options", expanded=False):
            clean_clutter = st.checkbox("Remove page clutter (headers, ads)", value=True)
            dismiss_banners = st.checkbox("Auto-dismiss cookie prompts", value=True)
            delay = st.slider("Render delay (seconds)", min_value=1, max_value=10, value=3)
            proxy_url = st.text_input(
                "Proxy Server (Optional)", 
                key="proxy_url", 
                help="Specify a proxy to bypass blocks (e.g. http://ip:port or http://user:pass@ip:port)"
            )


    st.markdown("---")
    
    with st.container(border=True):
        st.markdown("**2. Input Queries**")
        
        # Excel Upload
        uploaded_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
        if uploaded_file is not None:
            try:
                wb = load_workbook(uploaded_file, read_only=True)
                sheet = wb.active
                excel_terms = []
                for row in sheet.iter_rows(values_only=True):
                    if row and row[0] is not None:
                        val = str(row[0]).strip()
                        if val.lower() not in ["search term", "search terms", "query", "term", "queries", ""]:
                            excel_terms.append(val)
                wb.close()
                if excel_terms:
                    st.session_state.terms_input = "\n".join(excel_terms)
                    st.success(f"✓ Loaded {len(excel_terms)} queries from Excel!")
                else:
                    st.warning("No queries found in the Excel sheet.")
            except Exception as e:
                st.error(f"Error parsing Excel: {e}")

        # Text Area editor
        terms_text = st.text_area(
            "Queries (one per line):", 
            value=st.session_state.terms_input,
            height=250,
            help="Type, paste, or edit search terms directly here. Every line is treated as a separate search query."
        )
        
        # Sync back to state
        st.session_state.terms_input = terms_text
        active_terms = [t.strip() for t in terms_text.split("\n") if t.strip()]

# --- Right Column: Running & Outputs ---
with col_right:
    st.subheader("🚀 Automation Dashboard")
    
    # KPI metrics dashboard
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Queries", len(active_terms))
    m2.metric("Captured Items", len(st.session_state.outputs))
    m3.metric("Status", "Running ⚡" if st.session_state.is_running else "Idle 💤")

    # Start button (Uses type="primary" which triggers the gradient override)
    st.markdown("<br>", unsafe_allow_html=True)
    start_btn = st.button("▶️ Start Capture Job", disabled=st.session_state.is_running or not active_terms, type="primary", use_container_width=True)

    progress_bar = st.progress(0)
    progress_text = st.empty()

    # Log viewer
    st.markdown("### 📋 Activity Logs")
    logs_container = st.empty()
    if st.session_state.logs:
        logs_container.code("\n".join(st.session_state.logs))
    else:
        logs_container.info("No active log history. Start a job to see real-time updates.")

    # Captured Tab Panels
    st.markdown("### 🖼️ Export Gallery")
    output_tabs = st.tabs(["Screenshots", "PDF Documents", "ZIP Package"])

    # ZIP TAB
    with output_tabs[2]:
        if "zip_path" in st.session_state and os.path.exists(st.session_state.zip_path):
            with open(st.session_state.zip_path, "rb") as f:
                st.download_button(
                    label="📥 Download All Results (ZIP)",
                    data=f,
                    file_name=st.session_state.zip_name,
                    mime="application/zip",
                    use_container_width=True
                )
        else:
            st.info("Outputs zip file will be generated here once the job completes successfully.")

    # Run Job Process
    if start_btn:
        st.session_state.is_running = True
        st.session_state.logs = []
        st.session_state.outputs = []
        st.session_state.job_id = str(uuid.uuid4())
        job_dir = os.path.join(JOBS_DIR, st.session_state.job_id)
        os.makedirs(job_dir, exist_ok=True)
        
        append_log(f"Starting browser capture. Viewport: {viewport}, Engine: {engine}", "info")
        logs_container.code("\n".join(st.session_state.logs))
        
        width, height = map(int, viewport.split('x'))
        total_terms = len(active_terms)
        
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(
                    headless=True,
                    args=["--disable-blink-features=AutomationControlled"]
                )
                user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
                context = browser.new_context(
                    viewport={"width": width, "height": height},
                    user_agent=user_agent,
                    locale="en-US",
                    timezone_id="America/New_York",
                    extra_http_headers={
                        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                        "Accept-Language": "en-US,en;q=0.9",
                    }
                )
                
                # navigator webdriver footprint mask
                context.add_init_script("""
                    Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
                    window.navigator.chrome = { runtime: {}, loadTimes: () => {}, csi: () => {}, app: {} };
                """)
                
                page = context.new_page()
                
                for idx, term in enumerate(active_terms):
                    progress_percent = int(((idx) / total_terms) * 100)
                    progress_bar.progress(progress_percent)
                    progress_text.text(f"Processing term {idx + 1} of {total_terms}: '{term}'")
                    
                    append_log(f"[{idx+1}/{total_terms}] Navigating to: '{term}'", "info")
                    logs_container.code("\n".join(st.session_state.logs))
                    
                    url = get_search_url(engine, term, custom_url)
                    safe_chars = "".join([c if c.isalnum() or c in ".-_" else "_" for c in term]).strip("_")
                    
                    try:
                        page.goto(url, wait_until="domcontentloaded", timeout=25000)
                        time.sleep(delay)
                        
                        if dismiss_banners:
                            page.evaluate("""
                                const buttons = Array.from(document.querySelectorAll('button, a')).filter(el => {
                                    const t = el.innerText.toLowerCase();
                                    return t.includes('accept all') || t.includes('agree') || t.includes('accept cookies') || t.includes('i accept');
                                });
                                if (buttons.length > 0) buttons[0].click();
                            """)
                            time.sleep(0.5)
                        
                        if clean_clutter:
                            clean_page_content(page, engine)
                            time.sleep(0.5)
                        
                        outputs_generated = []
                        
                        # Screenshot Capture
                        if export_format in ["screenshot", "both"]:
                            filename = f"{idx+1}_screenshot_{safe_chars}.png"
                            filepath = os.path.join(job_dir, filename)
                            page.screenshot(path=filepath, full_page=False)
                            st.session_state.outputs.append({"type": "screenshot", "path": filepath, "name": filename, "term": term})
                            outputs_generated.append("Screenshot")
                            
                        # Full Page Screenshot Capture
                        if export_format == "screenshot_full":
                            filename = f"{idx+1}_screenshot_full_{safe_chars}.png"
                            filepath = os.path.join(job_dir, filename)
                            page.screenshot(path=filepath, full_page=True)
                            st.session_state.outputs.append({"type": "screenshot", "path": filepath, "name": filename, "term": term})
                            outputs_generated.append("Full Screenshot")

                        # PDF Capture
                        if export_format in ["pdf", "both"]:
                            filename = f"{idx+1}_document_{safe_chars}.pdf"
                            filepath = os.path.join(job_dir, filename)
                            page.pdf(path=filepath, format="A4", print_background=True)
                            st.session_state.outputs.append({"type": "pdf", "path": filepath, "name": filename, "term": term})
                            outputs_generated.append("PDF")
                            
                        append_log(f"✓ Successfully captured {', '.join(outputs_generated)} for '{term}'", "success")
                        
                    except Exception as page_err:
                        append_log(f"✗ Error processing '{term}': {page_err}", "error")
                    
                    logs_container.code("\n".join(st.session_state.logs))
                
                browser.close()
                
                # Bundle files into a ZIP
                zip_filename = f"export_job_{st.session_state.job_id}.zip"
                zip_path = os.path.join(job_dir, zip_filename)
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for out in st.session_state.outputs:
                        zipf.write(out["path"], out["name"])
                
                st.session_state.zip_path = zip_path
                st.session_state.zip_name = zip_filename
                
                progress_bar.progress(100)
                progress_text.text("Finished processing all search queries!")
                append_log("Job completed successfully. ZIP file is ready for download.", "success")
                logs_container.code("\n".join(st.session_state.logs))
                
        except Exception as job_err:
            append_log(f"Critical Automation Runner Error: {job_err}", "error")
            logs_container.code("\n".join(st.session_state.logs))
        
        st.session_state.is_running = False
        st.rerun()

    # RENDER SCREENSHOT TAB
    with output_tabs[0]:
        screenshots = [o for o in st.session_state.outputs if o["type"] == "screenshot"]
        if screenshots:
            # Display grid
            cols_gal = st.columns(2)
            for s_idx, s in enumerate(screenshots):
                col_gal = cols_gal[s_idx % 2]
                if os.path.exists(s["path"]):
                    col_gal.image(s["path"], caption=f"{s['term']}")
                    with open(s["path"], "rb") as sf:
                        col_gal.download_button(
                            label=f"Download Image",
                            data=sf,
                            file_name=s["name"],
                            mime="image/png",
                            key=f"dl_img_{s_idx}"
                        )
        else:
            st.info("No screenshots generated yet. Start a capture job.")

    # RENDER PDF TAB
    with output_tabs[1]:
        pdfs = [o for o in st.session_state.outputs if o["type"] == "pdf"]
        if pdfs:
            for p_idx, p in enumerate(pdfs):
                if os.path.exists(p["path"]):
                    col_pdf_name, col_pdf_dl = st.columns([3, 1])
                    col_pdf_name.markdown(f"📄 **{p['term']}** — `{p['name']}`")
                    with open(p["path"], "rb") as pf:
                        col_pdf_dl.download_button(
                            label="Download PDF",
                            data=pf,
                            file_name=p["name"],
                            mime="application/pdf",
                            key=f"dl_pdf_{p_idx}"
                        )
        else:
            st.info("No PDF files generated yet. Start a capture job.")
