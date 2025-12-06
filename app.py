import streamlit as st
import os
import sys
import tempfile
import re
import html
import json
import ast
from datetime import datetime

# --- THƯ VIỆN XỬ LÝ FILE (Đã loại bỏ docx2pdf) ---
from pypdf import PdfWriter, PdfReader
import pandas as pd
from PIL import Image as PILImage, ImageDraw, ImageFont

# --- THƯ VIỆN REPORTLAB ---
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as ReportLabImage, XPreformatted
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.fonts import addMapping
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from xhtml2pdf import pisa

# ==============================================================================
# 1. HỆ THỐNG CẤU HÌNH
# ==============================================================================
st.set_page_config(page_title="PDF Tool Pro V25", page_icon="📄", layout="wide")

# CSS giao diện
st.markdown("""
<style>
    .main { background-color: #f8f9fa; }
    h1 {
        background: -webkit-linear-gradient(45deg, #2c3e50, #3498db);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: bold;
        padding-bottom: 10px;
    }
    .stButton>button {
        background: linear-gradient(90deg, #3498db, #00f2fe);
        color: white;
        border: none;
        border-radius: 10px;
        font-weight: bold;
        transition: 0.3s;
    }
    .stButton>button:hover {
        transform: scale(1.02);
        box-shadow: 0 4px 15px rgba(52, 152, 219, 0.4);
    }
    .stFileUploader {
        border: 2px dashed #3498db;
        border-radius: 10px;
        padding: 20px;
        background-color: white;
    }
</style>
""", unsafe_allow_html=True)

def register_vietnamese_font():
    font_path = "Roboto-Regular.ttf"
    font_name = "VietFont"
    if not os.path.exists(font_path): return "Helvetica", "" 
    try:
        pdfmetrics.registerFont(TTFont(font_name, font_path))
        pdfmetrics.registerFont(TTFont(f'{font_name}-Bold', font_path))
        pdfmetrics.registerFont(TTFont(f'{font_name}-Italic', font_path))
        pdfmetrics.registerFont(TTFont(f'{font_name}-BoldItalic', font_path))
        addMapping(font_name, 0, 0, font_name)
        addMapping(font_name, 1, 0, f'{font_name}-Bold')
        addMapping(font_name, 0, 1, f'{font_name}-Italic')
        addMapping(font_name, 1, 1, f'{font_name}-BoldItalic')
        return font_name, ""
    except Exception: return "Helvetica", ""

# ==============================================================================
# 2. LOGIC BACKEND
# ==============================================================================
class PDFProcessor:
    PAGE_WIDTH, PAGE_HEIGHT = A4
    MARGIN_X = 1.5 * cm
    MARGIN_Y = 1.5 * cm
    AVAILABLE_WIDTH = PAGE_WIDTH - (2 * MARGIN_X)

    @staticmethod
    def read_text_file_safe(filepath):
        encodings = ['utf-8', 'utf-8-sig', 'cp1252', 'latin-1', 'utf-16']
        for enc in encodings:
            try:
                with open(filepath, 'r', encoding=enc) as f:
                    content = f.read()
                if content: return content
            except Exception: continue
        return ""

    @staticmethod
    def _clean_markdown_link(text): return re.sub(r'\[(.*?)\]\(.*?\)', r'\1', text)

    @staticmethod
    def _is_diagram_line(line):
        chars = ['|', '+', '┌', '└', '├', '─', '│']; line = line.strip()
        return False if not line else any(line.startswith(c) for c in chars)

    @staticmethod
    def _create_image_flowable(image_path):
        try:
            pil_img = PILImage.open(image_path); img_w, img_h = pil_img.size; aspect = img_h / float(img_w)
            safe_width = PDFProcessor.AVAILABLE_WIDTH; safe_height = PDFProcessor.PAGE_HEIGHT - (2 * PDFProcessor.MARGIN_Y)
            final_width = img_w; final_height = img_h
            if img_w > safe_width: final_width = safe_width; final_height = final_width * aspect
            if final_height > safe_height: final_height = safe_height; final_width = final_height / aspect
            return ReportLabImage(image_path, width=final_width, height=final_height)
        except: return None

    @staticmethod
    def _process_markdown_structure(text, styles, font_name):
        flowables = []; lines = text.split('\n')
        common_settings = {'fontName': font_name, 'wordWrap': 'CJK', 'splitLongWords': 1}
        style_normal = ParagraphStyle('VN_Normal', parent=styles['Normal'], fontSize=11, leading=14, spaceAfter=6, alignment=TA_JUSTIFY, **common_settings)
        style_h1 = ParagraphStyle('VN_H1', parent=styles['Heading1'], fontSize=15, leading=18, spaceBefore=12, spaceAfter=8, textColor='#000000', keepWithNext=True, **common_settings)
        style_h2 = ParagraphStyle('VN_H2', parent=styles['Heading2'], fontSize=13, leading=16, spaceBefore=10, spaceAfter=6, textColor='#222222', keepWithNext=True, **common_settings)
        style_list = ParagraphStyle('VN_List', parent=styles['Normal'], fontSize=11, leading=14, leftIndent=15, spaceAfter=4, **common_settings)
        style_code_wrap = ParagraphStyle('VN_Code_Wrap', parent=styles['Normal'], fontSize=10, leading=12, spaceBefore=6, spaceAfter=6, leftIndent=10, rightIndent=10, backColor='#f9f9f9', borderColor='#e0e0e0', borderWidth=0.5, borderPadding=6, fontName=font_name, wordWrap='CJK', splitLongWords=1)

        buffer_text = []; buffer_code = []; in_code_block = False

        def flush_text_buffer():
            if buffer_text:
                full = " ".join(buffer_text); full = PDFProcessor._clean_markdown_link(full)
                try: flowables.append(Paragraph(html.escape(full, quote=False).replace('**', '<b>', 1).replace('**', '</b>', 1).replace('*', '<i>', 1).replace('*', '</i>', 1), style_normal))
                except: flowables.append(Paragraph(html.escape(full), style_normal))
                buffer_text.clear()

        def flush_code_buffer(is_diagram=False):
            if buffer_code:
                if is_diagram: flowables.append(XPreformatted("\n".join(buffer_code), ParagraphStyle('Dia', parent=styles['Code'], fontSize=9, leading=11, fontName=font_name)))
                else: flowables.append(Paragraph("<br/>".join([html.escape(l) for l in buffer_code]), style_code_wrap))
                buffer_code.clear()

        i = 0
        while i < len(lines):
            line = lines[i]; raw = line; stripped = line.strip()
            if '```' in line:
                if in_code_block: flush_code_buffer(False); in_code_block = False
                else: flush_text_buffer(); in_code_block = True
                i += 1; continue
            if in_code_block: buffer_code.append(raw); i += 1; continue
            if PDFProcessor._is_diagram_line(stripped):
                flush_text_buffer()
                while i < len(lines) and (PDFProcessor._is_diagram_line(lines[i].strip()) or lines[i].strip() == ""): buffer_code.append(lines[i]); i += 1
                flush_code_buffer(True); continue
            if not stripped: flush_text_buffer(); i += 1; continue
            if stripped.startswith('#'):
                flush_text_buffer(); clean = html.escape(stripped.lstrip('#').strip())
                if stripped.startswith('##'): flowables.append(Paragraph(clean, style_h2))
                else: flowables.append(Paragraph(clean, style_h1))
            elif stripped.startswith('* ') or stripped.startswith('- '):
                flush_text_buffer(); flowables.append(Paragraph("• " + html.escape(stripped[2:]), style_list))
            elif stripped == '---' or stripped == '***': flush_text_buffer(); flowables.append(Spacer(1, 12))
            else: buffer_text.append(stripped)
            i += 1
        if in_code_block: flush_code_buffer(False)
        else: flush_text_buffer()
        return flowables

    @staticmethod
    def convert_and_merge_continuously(input_files, output_path, log_func=print):
        font, _ = register_vietnamese_font(); styles = getSampleStyleSheet(); full_story = []
        style_header = ParagraphStyle('FileHeader', parent=styles['Heading2'], fontName=font, fontSize=10, textColor='grey', alignment=TA_CENTER, spaceBefore=12, spaceAfter=12)
        for idx, fp in enumerate(input_files):
            ext = os.path.splitext(fp)[1].lower().replace('.', ''); fname = os.path.basename(fp)
            log_func(f"Gộp: {fname}")
            if idx > 0: full_story.extend([Spacer(1, 15), Paragraph(f"--- {fname} ---", style_header), Spacer(1, 5)])
            if ext in ['md', 'txt']:
                t = PDFProcessor.read_text_file_safe(fp)
                if t: full_story.extend(PDFProcessor._process_markdown_structure(t, styles, font))
            elif ext in ['jpg', 'png', 'jpeg', 'bmp']:
                img = PDFProcessor._create_image_flowable(fp)
                if img: full_story.extend([img, Spacer(1, 10)])
        if full_story:
            try: SimpleDocTemplate(output_path, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm).build(full_story); return True
            except Exception as e: log_func(f"Lỗi: {e}"); return False
        return False

    @staticmethod
    def convert_single_file(input_path, output_folder=None, log_func=print):
        try:
            base = os.path.basename(input_path); ext = os.path.splitext(input_path)[1].lower().replace('.', '')
            out = os.path.join(output_folder, os.path.splitext(base)[0] + ".pdf") if output_folder else os.path.splitext(input_path)[0] + "_converted.pdf"
            font, css = register_vietnamese_font()
            
            if ext == 'docx':
                log_func("⚠️ File Word (.docx) không hỗ trợ trên Web.")
                return None 
            elif ext in ['jpg', 'png', 'jpeg', 'bmp']:
                img = PDFProcessor._create_image_flowable(input_path)
                if img: SimpleDocTemplate(out, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm).build([img]); return out
            elif ext in ['xlsx', 'xls']:
                df = pd.read_excel(input_path)
                css_t = f"<style>{css} @page {{size:A4; margin:1cm;}} table {{width:100%; border-collapse:collapse; font-size:8pt;}} td,th {{border:1px solid #333; padding:3px;}}</style>"
                html_c = f"<html><head><meta charset='UTF-8'>{css_t}</head><body>{df.to_html(index=False, border=1, classes='table')}</body></html>"
                with open(out, "wb") as f: pisa.CreatePDF(html_c, dest=f, encoding='utf-8')
                return out
            elif ext in ['md', 'txt']:
                t = PDFProcessor.read_text_file_safe(input_path)
                if t: SimpleDocTemplate(out, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm).build(PDFProcessor._process_markdown_structure(t, getSampleStyleSheet(), font)); return out
            elif ext == 'pdf': return input_path
            return None
        except Exception as e: log_func(f"Lỗi: {e}"); return None

    @staticmethod
    def split_pdf(input_path, pages_per_file):
        reader = PdfReader(input_path); total = len(reader.pages); base = os.path.splitext(input_path)[0]; res = []
        for i in range(0, total, pages_per_file):
            writer = PdfWriter()
            for p in range(i, min(i+pages_per_file, total)): writer.add_page(reader.pages[p])
            out = f"{base}_part_{i//pages_per_file + 1}.pdf"; 
            with open(out, "wb") as f: writer.write(f)
            res.append(out)
        return res

    @staticmethod
    def merge_pdfs(input_paths, output_path):
        writer = PdfWriter()
        for path in input_paths:
            reader = PdfReader(path)
            for page in reader.pages: writer.add_page(page)
        with open(output_path, "wb") as f: writer.write(f)
        return output_path

    @staticmethod
    def extract_text(input_path, output_path):
        try:
            reader = PdfReader(input_path); full = []
            for i, p in enumerate(reader.pages):
                t = p.extract_text(); full.append(t if t else f"[Trang {i+1}: Ảnh/Không có text]")
            if not full: return False
            with open(output_path, "w", encoding="utf-8") as f: f.write("\n\n--- PAGE BREAK ---\n\n".join(full))
            return True
        except: return False

    @staticmethod
    def extract_images_smart(input_path, output_folder, create_layout=False):
        try:
            reader = PdfReader(input_path); imgs = []; count = 0; base = os.path.splitext(os.path.basename(input_path))[0]
            for i, page in enumerate(reader.pages):
                if hasattr(page, 'images') and page.images:
                    for img in page.images:
                        count += 1; name = f"{base}_p{i+1}_{count}.{img.name.split('.')[-1]}"; path = os.path.join(output_folder, name)
                        with open(path, "wb") as fp: fp.write(img.data)
                        if create_layout: 
                            try: imgs.append(PILImage.open(path)) 
                            except: pass
            if create_layout and imgs:
                tw, th = 400, 400; pad = 20; cols = 3; rows = (len(imgs) + cols - 1) // cols
                sw = cols*(tw+pad)+pad; sh = rows*(th+pad)+pad; sheet = PILImage.new('RGB', (sw, sh), 'white')
                for idx, im in enumerate(imgs):
                    im.thumbnail((tw, th)); row = idx//cols; col = idx%cols
                    x = col*(tw+pad)+pad; y = row*(th+pad)+pad
                    sheet.paste(im, (x + (tw-im.size[0])//2, y + (th-im.size[1])//2))
                sheet.save(os.path.join(output_folder, f"{base}_SmartLayout.jpg"))
            return count
        except: return 0

    @staticmethod
    def clean_data_structure(input_path, output_path, log_func=print):
        try:
            log_func(f"Đang xử lý file (Mode: Optimized Structure)...")
            with open(input_path, 'r', encoding='utf-8') as f: content = f.read()
            cleaned = []
            is_transcript = False
            
            try:
                data_json = json.loads(content)
                if isinstance(data_json, dict) and "events" in data_json:
                    log_func("-> Phát hiện định dạng Transcript JSON...")
                    full_t = []
                    for ev in data_json.get("events", []):
                        if "segs" in ev:
                            for seg in ev.get("segs", []):
                                if "utf8" in seg: full_t.append(seg["utf8"].replace('\n', ' '))
                    if full_t:
                        joined = re.sub(r'\s+', ' ', "".join(full_t)).strip()
                        cleaned.append(f"[TRANSCRIPT CONTENT]:\n{joined}\n")
                        is_transcript = True
            except: pass

            if not is_transcript:
                try:
                    ds = ast.literal_eval(content)
                    if isinstance(ds, dict): ds = [ds]
                    if isinstance(ds, list):
                        for e in ds:
                            if isinstance(e, dict):
                                role = e.get("role", "UNKNOWN").upper(); parts = e.get("parts", "")
                                tmp = []
                                if isinstance(parts, list):
                                    for p in parts:
                                        val = p["text"] if isinstance(p, dict) and "text" in p else (p if isinstance(p, str) else "")
                                        if val: tmp.append(val.strip())
                                else: tmp.append(str(parts).strip())
                                txt = re.sub(r'\s+', ' ', " ".join(tmp)).strip()
                                cleaned.append(f"[{role}]:\n{txt}\n")
                except:
                    blocks = re.findall(r'\{.*?\}', content, re.DOTALL)
                    for b in blocks:
                        try:
                            e = ast.literal_eval(b)
                            if isinstance(e, dict) and ("parts" in e or "role" in e):
                                role = e.get("role", "SYSTEM").upper(); parts = e.get("parts", "")
                                tmp = []
                                if isinstance(parts, list):
                                    for p in parts:
                                        val = p["text"] if isinstance(p, dict) and "text" in p else (p if isinstance(p, str) else "")
                                        if val: tmp.append(val.strip())
                                else: tmp.append(str(parts).strip())
                                txt = re.sub(r'\s+', ' ', " ".join(tmp)).strip()
                                cleaned.append(f"[{role}]:\n{txt}\n")
                        except: continue

            if not cleaned: log_func("-> Không tìm thấy dữ liệu hợp lệ."); return False
            with open(output_path, "w", encoding="utf-8") as out:
                out.write("\n" + "-"*50 + "\n" + ("\n" + "-"*50 + "\n").join(cleaned))
            return True
        except Exception as e:
            log_func(f"Lỗi Critical (Cleaner): {e}"); traceback.print_exc(); return False

# ==============================================================================
# 3. STREAMLIT UI
# ==============================================================================

def save_uploaded_file(uploaded_file):
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            return tmp_file.name
    except Exception as e:
        return None

st.title("📄 PDF Tool Pro V25 (Web Edition)")
st.caption("✨ Phiên bản tối ưu cho Web/Mobile (Không hỗ trợ file Word docx)")
st.markdown("---")

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Ghép File", "Tách File", "Chuyển Đổi", "Trích Xuất", "Làm Sạch Data"])

with tab1:
    st.header("Ghép nhiều file PDF")
    uploaded_files = st.file_uploader("Chọn các file PDF (Kéo thả vào đây)", type=["pdf"], accept_multiple_files=True)
    if st.button("Bắt đầu Ghép") and uploaded_files:
        with st.spinner("Đang xử lý..."):
            file_paths = [save_uploaded_file(f) for f in uploaded_files]
            output_pdf = tempfile.mktemp(suffix=".pdf")
            try:
                PDFProcessor.merge_pdfs(file_paths, output_pdf)
                with open(output_pdf, "rb") as f:
                    st.success("Thành công!")
                    st.download_button("Tải File Đã Ghép", f, file_name="merged.pdf", mime="application/pdf")
            except Exception as e:
                st.error(f"Lỗi: {e}")
            finally:
                for fp in file_paths: os.remove(fp)

with tab2:
    st.header("Tách nhỏ file PDF")
    up_split = st.file_uploader("Chọn file PDF cần tách", type=["pdf"], key="split")
    pages_per = st.number_input("Số trang mỗi file con", min_value=1, value=1)
    if st.button("Tách File") and up_split:
        with st.spinner("Đang tách..."):
            fp = save_uploaded_file(up_split)
            try:
                res_files = PDFProcessor.split_pdf(fp, pages_per)
                st.success(f"Đã tách thành {len(res_files)} file!")
                import zipfile
                zip_path = tempfile.mktemp(suffix=".zip")
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for rf in res_files:
                        zipf.write(rf, os.path.basename(rf))
                with open(zip_path, "rb") as f:
                    st.download_button("Tải Tất Cả (ZIP)", f, file_name="split_files.zip", mime="application/zip")
            except Exception as e: st.error(f"Lỗi: {e}")

with tab3:
    st.header("Chuyển đổi sang PDF")
    st.info("Hỗ trợ: Excel (.xlsx), Text (.txt, .md), Ảnh (.jpg, .png). ❌ Không hỗ trợ Word (.docx)")
    up_conv = st.file_uploader("Chọn file...", accept_multiple_files=True, key="conv")
    smart_merge = st.checkbox("Gộp tất cả kết quả thành 1 file PDF duy nhất?")
    if st.button("Convert") and up_conv:
        with st.spinner("Đang chuyển đổi..."):
            file_paths = [save_uploaded_file(f) for f in up_conv]
            temp_dir = tempfile.mkdtemp()
            try:
                if smart_merge:
                    out_path = os.path.join(temp_dir, "SmartMerged.pdf")
                    PDFProcessor.convert_and_merge_continuously(file_paths, out_path)
                    with open(out_path, "rb") as f:
                        st.download_button("Tải File Gộp", f, file_name="converted_merged.pdf")
                else:
                    import zipfile
                    zip_path = tempfile.mktemp(suffix=".zip")
                    with zipfile.ZipFile(zip_path, 'w') as zipf:
                        for fp in file_paths:
                            res = PDFProcessor.convert_single_file(fp, temp_dir)
                            if res: zipf.write(res, os.path.basename(res))
                    with open(zip_path, "rb") as f:
                        st.download_button("Tải File (ZIP)", f, file_name="converted.zip")
            except Exception as e: st.error(f"Lỗi: {e}")

with tab4:
    st.header("Trích xuất Text / Ảnh từ PDF")
    up_ext = st.file_uploader("File PDF nguồn", type=["pdf"], key="ext")
    mode = st.radio("Chế độ:", ["Lấy Text (.txt)", "Lấy Ảnh (ZIP)", "Lấy Ảnh + Layout (ZIP)"])
    if st.button("Trích xuất") and up_ext:
        with st.spinner("Đang xử lý..."):
            fp = save_uploaded_file(up_ext)
            temp_dir = tempfile.mkdtemp()
            try:
                if mode == "Lấy Text (.txt)":
                    out_txt = os.path.join(temp_dir, "extracted.txt")
                    if PDFProcessor.extract_text(fp, out_txt):
                        with open(out_txt, "r", encoding="utf-8") as f:
                            st.download_button("Tải Text", f, file_name="extracted.txt")
                    else: st.warning("Không tìm thấy text (File Scan?)")
                else:
                    create_layout = "Layout" in mode
                    PDFProcessor.extract_images_smart(fp, temp_dir, create_layout)
                    import zipfile
                    zip_path = tempfile.mktemp(suffix=".zip")
                    with zipfile.ZipFile(zip_path, 'w') as zipf:
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                zipf.write(os.path.join(root, file), file)
                    with open(zip_path, "rb") as f:
                        st.download_button("Tải Ảnh (ZIP)", f, file_name="images.zip")
            except Exception as e: st.error(f"Lỗi: {e}")

with tab5:
    st.header("Làm sạch Data (Transcript/JSON)")
    st.info("Hỗ trợ: JSON Transcript (Events), Log Gemini. Tự động gộp đoạn văn, loại bỏ xuống dòng thừa.")
    up_clean = st.file_uploader("Upload file Data (.txt, .json)", key="clean")
    if st.button("Clean Data") and up_clean:
        with st.spinner("Analyzing & Cleaning..."):
            fp = save_uploaded_file(up_clean)
            out_path = tempfile.mktemp(suffix=".txt")
            log_container = st.empty()
            def web_log(msg): log_container.text(f"Log: {msg}")
            if PDFProcessor.clean_data_structure(fp, out_path, web_log):
                with open(out_path, "r", encoding="utf-8") as f:
                    st.success("Làm sạch thành công!")
                    st.download_button("Tải File Sạch (.txt)", f, file_name=f"Cleaned_{up_clean.name}.txt")
            else:
                st.error("Không tìm thấy cấu trúc dữ liệu hợp lệ.")

st.markdown("---")
st.caption("Developed by Senior Production Teamleader (V25 Web Edition)")
