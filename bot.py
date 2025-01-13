import pytesseract
from pdf2image import convert_from_path
import sys
import re
import os
from pathlib import Path

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_info_from_text(text):
    info = {}
    
    # Tìm lĩnh vực hoạt động (Điều 2)
    dieu2_match = re.search(r'Điều 2\.(.*?)Điều 3\.', text, re.DOTALL)
    if dieu2_match:
        dieu2_text = dieu2_match.group(1).strip()
        linh_vuc_points = re.findall(r'\d+\.\s*(.*?)(?=\d+\.|$)', dieu2_text, re.DOTALL)
        if linh_vuc_points:
            # Format lĩnh vực với xuống dòng và dấu gạch đầu dòng
            formatted_points = [point.strip().replace('\n', ' ') for point in linh_vuc_points]
            info['Lĩnh vực'] = '\n+ ' + '\n+ '.join(formatted_points)
        else:
            info['Lĩnh vực'] = dieu2_text.replace('Lĩnh vực hoạt động điện lực được cấp phép:', '').strip()
    
    # Tìm phạm vi hoạt động (Điều 3)
    dieu3_match = re.search(r'Điều 3\.(.*?)Điều 4\.', text, re.DOTALL)
    if dieu3_match:
        info['Phạm vi hoạt động'] = dieu3_match.group(1).strip().replace('Phạm vi hoạt động', '').strip()
    
    # Tìm Số giấy phép (có thể sai) - sửa lại regex để bắt linh hoạt hơn
    gp_patterns = [
        r'Số\s*([^,\n]+)\/GP-[^\s,\n]+',  # Pattern cho dạng Số xxx/GP-xxx
        r'Số:\s*([^,\n]+)',               # Pattern cho dạng Số: xxx
        r'Số\s+(\d+[^\n,]+)'              # Pattern cho các dạng Số xxx khác
    ]
    
    for pattern in gp_patterns:
        gp_match = re.search(pattern, text)
        if gp_match:
            so_gp = gp_match.group(0).strip()  # Lấy toàn bộ chuỗi match
            info['Số giấy phép (có thể sai)'] = so_gp.replace('Số', '').replace(':', '').strip()
            break
    
    # Tìm Ngày cấp (có thể sai) với pattern linh hoạt hơn
    date_match = re.search(r'Hà Nội,\s+ngày\s+([⁄\d]+)\s+tháng\s+(\d+)\s+năm\s+(\d+)', text)
    if date_match:
        day, month, year = date_match.groups()
        # Chuyển ký tự ⁄ thành số 1 nếu có
        day = day.replace('⁄', '1')
        info['Ngày cấp (có thể sai)'] = f"{day}/{month}/{year}"
        info['Ngày hiệu lực (có thể sai)'] = f"{day}/{month}/{year}"
    else:
        # Backup: tìm theo format cũ nếu không tìm thấy format mới
        date_match = re.search(r'ngày\s+(\d+)\s+tháng\s+(\d+)\s+năm\s+(\d+)', text)
        if date_match:
            day, month, year = date_match.groups()
            info['Ngày cấp (có thể sai)'] = f"{day}/{month}/{year}"
            info['Ngày hiệu lực (có thể sai)'] = f"{day}/{month}/{year}"
    
    # Tìm Ngày hết hạn (có thể sai) (Điều 4)
    dieu4_match = re.search(r'Điều 4\..*?ngày\s+(\d+)\s+tháng\s+(\d+)\s+năm\s+(\d+)', text, re.DOTALL)
    if dieu4_match:
        day, month, year = dieu4_match.groups()
        info['Ngày hết hạn (có thể sai)'] = f"{day}/{month}/{year}"
    
    # Tìm công suất MW
    mw_matches = re.findall(r'(\d+[,.]?\d*)\s*MW', text)
    if mw_matches:
        max_mw = max([float(mw.replace(',', '.')) for mw in mw_matches])
        info['Công suất MW'] = f"{max_mw}MW"
    
    # Tìm thông tin từ Điều 1
    dieu1_match = re.search(r'Điều 1\.(.*?)Điều 2\.', text, re.DOTALL)
    if dieu1_match:
        dieu1_text = dieu1_match.group(1)
        
        # Tìm tên tổ chức
        ten_match = re.search(r'1\.\s*Tên tổ chức:\s*(.*?)[\n\.]', dieu1_text)
        if ten_match:
            info['Tên tổ chức'] = ten_match.group(1).strip()
        
        # Tìm mã số thuế và ĐKKD
        dkkd_match = re.search(r'2\.\s*(Giấy.*?)\s*3\.', dieu1_text, re.DOTALL)
        if dkkd_match:
            dkkd_text = dkkd_match.group(1).strip()
            mst_match = re.search(r'số\s+(\d+)', dkkd_text)
            if mst_match:
                info['Mã số thuế'] = mst_match.group(1)
            info['Giấy chứng nhận đăng ký doanh nghiệp'] = dkkd_text
        
        # Tìm trụ sở chính
        tru_so_match = re.search(r'3\.\s*Trụ sở.*?:\s*(.*?)(?:Điện thoại|$)', dieu1_text, re.DOTALL)
        if tru_so_match:
            info['Trụ sở chính'] = tru_so_match.group(1).strip()
        
        # Tìm điện thoại
        phone_match = re.search(r'Điện thoại:\s*([\d\.\s]+)', dieu1_text)
        if phone_match:
            info['Điện thoại'] = phone_match.group(1).strip()
    
    return info

def read_pdf_text(pdf_path, max_pages=4):
    try:
        poppler_path = r"C:\Program Files\poppler\Library\bin"
        print(f"\nĐang xử lý file: {pdf_path}")
        
        pages = convert_from_path(
            pdf_path,
            poppler_path=poppler_path
        )
        
        full_text = ""
        pages_to_process = min(len(pages), max_pages)
        
        for i in range(pages_to_process):
            print(f"Đang xử lý trang {i+1}/{pages_to_process}...")
            text = pytesseract.image_to_string(pages[i], lang='vie')
            full_text += text + "\n"
            
        return full_text
        
    except Exception as e:
        print(f"Có lỗi khi đọc file {pdf_path}: {str(e)}")
        return None

def process_docs_folder():
    # Đường dẫn đến thư mục docs
    docs_path = Path("docs")
    
    # Lấy danh sách các file PDF và sắp xếp
    pdf_files = sorted([f for f in docs_path.glob("*.pdf")])
    
    # Mở file output.txt để ghi kết quả
    with open('output.txt', 'w', encoding='utf-8') as out_file:
        
        # Xử lý từng file PDF
        for pdf_file in pdf_files:
            try:
                text = read_pdf_text(str(pdf_file))
                if text:
                    # Lưu raw text
                    with open('raw_text.txt', 'w', encoding='utf-8') as raw_file:
                        raw_file.write(text)
                    
                    # Trích xuất thông tin
                    info = extract_info_from_text(text)
                    
                    # Ghi kết quả vào output.txt
                    out_file.write(f"==={pdf_file.name}===\n\n")
                    for key, value in info.items():
                        out_file.write(f"- {key}:\n{value}\n\n")
                    out_file.write("\n")
                    
                    print(f"Đã xử lý xong file {pdf_file.name}")
                    
            except Exception as e:
                print(f"Lỗi khi xử lý file {pdf_file.name}: {str(e)}")
                continue

if __name__ == "__main__":
    process_docs_folder()
