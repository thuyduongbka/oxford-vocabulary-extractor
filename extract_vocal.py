import os
from bs4 import BeautifulSoup
import pandas as pd # type: ignore

def extract_from_file(file_path, subject):
    # Đọc nội dung của file HTML
    with open(file_path, 'r', encoding='utf-8') as file:
        html_content = file.read()

    # Sử dụng BeautifulSoup để phân tích cú pháp HTML
    soup = BeautifulSoup(html_content, 'html.parser')

    # Lấy tên topic từ tên file
    topic = os.path.basename(file_path).replace('.html', '')

    # Tạo danh sách để lưu trữ dữ liệu
    data = []

    # Tìm tất cả các thẻ li trong ul có class là "top-g"
    list_items = soup.find('ul', class_='top-g').find_all('li')

    for li in list_items:
        word = li.a.text.strip()
        pos = li.find('span', class_='pos').text.strip()
        level = li.find('span', class_='belong-to').text.strip()
        
        # Tìm thuộc tính bắt đầu với 'data-' ngoại trừ 'data-hw'
        attribute_name = ""
        for key in li.attrs:
            if key.startswith('data-') and key != 'data-hw':
                attribute_name = key[5:-2]  # Loại bỏ '_t' ở cuối
                break
        
        # Thêm vào danh sách
        data.append([subject, topic, word, pos, level, attribute_name])

    return data

def read_html_files(folder_path):
    # Tạo danh sách để lưu trữ dữ liệu
    combined_data = []
    
    # Lặp qua tất cả các thư mục và file trong thư mục html
    for root, dirs, files in os.walk(folder_path):
        # Lấy tên subject từ tên thư mục cha
        subject = os.path.basename(root)
        for file in files:
            if file.endswith('.html'):
                file_path = os.path.join(root, file)
                combined_data.extend(extract_from_file(file_path, subject))

    return combined_data

# Đường dẫn tới thư mục chứa các thư mục html
html_folder = 'html'

# Tạo danh sách dữ liệu từ các file HTML
combined_data = read_html_files(html_folder)

# Tạo DataFrame từ dữ liệu kết hợp
df = pd.DataFrame(combined_data, columns=['Subject', 'Topic', 'Word', 'POS', 'Level', 'Sub Topic'])

# Ghi DataFrame vào file Excel
df.to_excel('vocabulary.xlsx', index=False)
