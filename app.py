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

# App layout & theme configuration (Must be the first Streamlit command)
st.set_page_config(
    page_title="Capture Studio — Search & Export Tool",
    page_icon="📸",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Auto-install Playwright browser if running in a cloud environment
try:
    import subprocess
    @st.cache_resource
    def install_playwright_browsers():
        subprocess.run(["playwright", "install", "chromium"], check=True)
    install_playwright_browsers()
except Exception as e:
    pass

# Initialize Session States
if "terms" not in st.session_state:
    st.session_state.terms = ["triclosan suture", "aspirin cardiovascular", "ibuprofen inflammation"]
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
    elif engine == 'custom' and custom_template:
        return custom_template.replace("{query}", encoded)
    return f"https://pmc.ncbi.nlm.nih.gov/search/?term={encoded}"

# Header Styling
st.markdown("""
<div style="background-color:#1e1e2f; padding:20px; border-radius:10px; margin-bottom:25px; border-left: 5px solid #ff4b4b">
    <h1 style="color:white; margin:0;">📸 Capture Studio — Search & Export</h1>
    <p style="color:#a3a3c2; margin:5px 0 0 0;">Batch screenshot and PDF exporter for detailed web audits powered by Playwright.</p>
</div>
""", unsafe_allow_html=True)


# Layout Setup
col_left, col_right = st.columns([1, 2])

# Sidebar Settings
st.sidebar.title("⚙️ Control Settings")
engine = st.sidebar.selectbox("Search Engine", ["pmc", "pubmed", "google", "scholar", "custom"])
custom_url = ""
if engine == "custom":
    custom_url = st.sidebar.text_input("Custom URL Template", value="https://example.com/search?q={query}")

export_format = st.sidebar.selectbox("Export Format", ["screenshot", "screenshot_full", "pdf", "both"])
viewport = st.sidebar.selectbox("Viewport Size", ["1920x1080", "1280x720", "1536x864", "1024x768"])
clean_clutter = st.sidebar.checkbox("Remove Clutter (Header/Footer/Banners)", value=True)
dismiss_banners = st.sidebar.checkbox("Auto-Dismiss Cookie Banners", value=True)
delay = st.sidebar.slider("Render Delay (seconds)", min_value=1, max_value=10, value=3)

# --- Left Column: Search Terms Management ---
with col_left:
    st.subheader("📝 Search terms")
    
    # Excel Upload
    uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])
    if uploaded_file is not None:
        try:
            wb = load_workbook(uploaded_file, read_only=True)
            sheet = wb.active
            terms = []
            for row in sheet.iter_rows(values_only=True):
                if row and row[0] is not None:
                    val = str(row[0]).strip()
                    if val.lower() not in ["search term", "search terms", "query", "term", "queries", ""]:
                        terms.append(val)
            wb.close()
            if terms:
                st.session_state.terms = terms
                st.success(f"Loaded {len(terms)} terms from Excel!")
            else:
                st.warning("No search terms found in sheet.")
        except Exception as e:
            st.error(f"Error parsing Excel file: {e}")

    # Manual input
    new_term = st.text_input("Add manual term:")
    if st.button("➕ Add Term") and new_term.strip():
        st.session_state.terms.append(new_term.strip())
        st.rerun()

    # Terms List
    st.markdown("### Active Terms List")
    if st.session_state.terms:
        # Clear list button
        if st.button("🗑️ Clear All Terms"):
            st.session_state.terms = []
            st.rerun()

        for idx, term in enumerate(st.session_state.terms):
            col_t_text, col_t_del = st.columns([4, 1])
            col_t_text.markdown(f"**{idx + 1}.** `{term}`")
            if col_t_del.button("❌", key=f"del_{idx}"):
                st.session_state.terms.pop(idx)
                st.rerun()
    else:
        st.info("No active search terms. Type one above or upload an Excel sheet.")

# --- Right Column: Capture Dashboard ---
with col_right:
    st.subheader("🚀 Automation Dashboard")
    
    # Run Buttons
    col_run_1, col_run_2 = st.columns([1, 1])
    start_btn = col_run_1.button("▶️ Start Capture Job", disabled=st.session_state.is_running or not st.session_state.terms, use_container_width=True)
    
    # Progress Indicators
    progress_bar = st.progress(0)
    progress_text = st.empty()
    
    # Log box container
    st.markdown("### 📋 Activity Logs")
    logs_container = st.empty()
    
    # Gallery Output Sections
    st.markdown("### 🖼️ Generated Outputs")
    output_tabs = st.tabs(["Screenshots", "PDF Documents"])
    
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
        total_terms = len(st.session_state.terms)
        
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
                
                # navigator footprint mask
                context.add_init_script("""
                    Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
                    window.navigator.chrome = { runtime: {}, loadTimes: () => {}, csi: () => {}, app: {} };
                """)
                
                page = context.new_page()
                
                for idx, term in enumerate(st.session_state.terms):
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
                progress_text.text("Finished processing all search terms!")
                append_log("Job completed successfully. ZIP file is ready for download.", "success")
                logs_container.code("\n".join(st.session_state.logs))
                
        except Exception as job_err:
            append_log(f"Critical Automation Runner Error: {job_err}", "error")
            logs_container.code("\n".join(st.session_state.logs))
        
        st.session_state.is_running = False
        st.rerun()

    # Render Active Logs
    if st.session_state.logs:
        logs_container.code("\n".join(st.session_state.logs))

    # ZIP Download Button
    if "zip_path" in st.session_state and os.path.exists(st.session_state.zip_path):
        with open(st.session_state.zip_path, "rb") as f:
            st.download_button(
                label="📥 Download All Results (ZIP)",
                data=f,
                file_name=st.session_state.zip_name,
                mime="application/zip",
                use_container_width=True
            )

    # Render gallery items in tabs
    with output_tabs[0]:
        screenshots = [o for o in st.session_state.outputs if o["type"] == "screenshot"]
        if screenshots:
            # Display grid
            cols_gal = st.columns(3)
            for s_idx, s in enumerate(screenshots):
                col_gal = cols_gal[s_idx % 3]
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
            st.info("No screenshots generated yet.")
            
    with output_tabs[1]:
        pdfs = [o for o in st.session_state.outputs if o["type"] == "pdf"]
        if pdfs:
            for p_idx, p in enumerate(pdfs):
                if os.path.exists(p["path"]):
                    col_pdf_name, col_pdf_dl = st.columns([4, 1])
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
            st.info("No PDF files generated yet.")
