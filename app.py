from flask import Flask, request, jsonify, send_file, render_template
import zipfile
import os
import pandas as pd
import tempfile
from io import BytesIO
import re

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

def extract_date_from_filename(filename):
    match = re.search(r'(\d{4})(\d{2})(\d{2})', filename)
    if match:
        return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
    return filename

def parse_excel_file(filepath):
    try:
        df = pd.read_excel(filepath)
        data = []
        header_found = False
        header_row = 0
        
        for i in range(df.shape[0]):
            row = df.iloc[i].tolist()
            if row and isinstance(row[0], str) and '姓名' in row[0]:
                header_found = True
                header_row = i
                break
        
        if not header_found:
            return []
        
        columns = df.iloc[header_row].tolist()
        name_idx = columns.index('姓名') if '姓名' in columns else 0
        id_idx = columns.index('学号/工号') if '学号/工号' in columns else 1
        class_idx = columns.index('行政班级') if '行政班级' in columns else 5
        if '统计' in columns:
            status_idx = columns.index('统计')
        elif '签到状态' in columns:
            status_idx = columns.index('签到状态')
        else:
            status_idx = 10
        
        for i in range(header_row + 1, df.shape[0]):
            row = df.iloc[i].tolist()
            if len(row) > max(name_idx, id_idx, class_idx, status_idx):
                name = str(row[name_idx]).strip() if pd.notna(row[name_idx]) else ''
                student_id = str(row[id_idx]).strip() if pd.notna(row[id_idx]) else ''
                class_name = str(row[class_idx]).strip() if pd.notna(row[class_idx]) else ''
                status = str(row[status_idx]).strip() if pd.notna(row[status_idx]) else '未参与'
                
                if name and student_id:
                    data.append({
                        '姓名': name,
                        '学号': student_id,
                        '班级': class_name,
                        '签到状态': status
                    })
        
        return data
    except Exception as e:
        print(f"解析Excel文件失败: {e}")
        return []

def process_zip_file(zip_path):
    student_records = {}
    dates = []

    with tempfile.TemporaryDirectory() as temp_dir:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        for root, dirs, files in os.walk(temp_dir):
            for file_name in files:
                if file_name.endswith('.xlsx') and not file_name.startswith('~'):
                    file_path = os.path.join(root, file_name)
                    date_str = extract_date_from_filename(file_name)
                    dates.append(date_str)

                    try:
                        data = parse_excel_file(file_path)
                        for record in data:
                            key = (record['学号'], record['姓名'])
                            if key not in student_records:
                                student_records[key] = {
                                    '姓名': record['姓名'],
                                    '学号': record['学号'],
                                    '班级': record['班级']
                                }
                            student_records[key][date_str] = record['签到状态']
                    except Exception as e:
                        print(f"处理文件 {file_name} 失败: {e}")
                        continue

    dates.sort()

    result = []
    for key in sorted(student_records.keys()):
        record = student_records[key].copy()
        for date in dates:
            if date not in record:
                record[date] = '未参与'
        result.append(record)

    return result, dates

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': '请选择文件'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '请选择文件'}), 400
    
    if not file.filename.endswith('.zip'):
        return jsonify({'error': '请上传zip文件'}), 400
    
    try:
        with tempfile.NamedTemporaryFile(suffix='.zip', delete=False) as temp_zip:
            temp_zip.write(file.read())
            temp_zip_path = temp_zip.name
        
        try:
            data, dates = process_zip_file(temp_zip_path)
        finally:
            os.unlink(temp_zip_path)
        
        if not data:
            return jsonify({'error': '未能解析到有效数据，请检查上传的文件是否包含正确格式的Excel签到文件'}), 400
        
        return jsonify({
            'success': True,
            'data': data,
            'dates': dates,
            'columns': ['姓名', '学号', '班级'] + dates
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download', methods=['POST'])
def download_excel():
    try:
        data = request.json.get('data', [])
        dates = request.json.get('dates', [])
        
        columns = ['姓名', '学号', '班级'] + dates
        df = pd.DataFrame(data, columns=columns)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='签到汇总')
        
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            download_name='签到汇总.xlsx',
            as_attachment=True
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5001)