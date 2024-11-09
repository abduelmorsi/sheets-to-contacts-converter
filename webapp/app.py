from flask import Flask, render_template, request, send_file
import pandas as pd
import io
import requests

app = Flask(__name__)

# Add this new function to create VCF content
def create_vcf(df, name_col, phone_col):
    vcf_content = []
    for _, row in df.iterrows():
        vcf_content.extend([
            'BEGIN:VCARD',
            'VERSION:3.0',
            f'FN:{row[name_col]}',
            f'TEL;TYPE=CELL:{row[phone_col]}',
            'END:VCARD',
            ''
        ])
    return '\n'.join(vcf_content)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            # Get form data
            sheet_url = request.form['sheet_url']
            name_col = request.form['name_col'].upper()
            phone_col = request.form['phone_col'].upper()
            name_prefix = request.form.get('name_prefix', '')
            
            # Handle Google Sheets URL or uploaded CSV
            if "docs.google.com/spreadsheets" in sheet_url:
                sheet_id = sheet_url.split('/d/')[1].split('/')[0]
                csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
                response = requests.get(csv_url)
                if response.status_code != 200:
                    return "Unable to access the sheet. Make sure it's published to the web.", 400
                df = pd.read_csv(io.StringIO(response.text))
            else:
                file = request.files['csv_file']
                if file:
                    df = pd.read_csv(file)
                else:
                    return "No file uploaded", 400

            # Convert columns
            name_col_idx = ord(name_col) - ord('A')
            phone_col_idx = ord(phone_col) - ord('A')
            
            columns = df.columns.tolist()
            if name_col_idx >= len(columns) or phone_col_idx >= len(columns):
                return "Column letter is out of range", 400
                
            name_col = columns[name_col_idx]
            phone_col = columns[phone_col_idx]

            # Add prefix if specified
            if name_prefix:
                df[name_col] = name_prefix + " " + df[name_col]

            # Add new form field for format
            output_format = request.form.get('output_format', 'csv')

            if output_format == 'csv':
                # Existing CSV creation code
                output = pd.DataFrame({
                    'Name': df[name_col],
                    'Given Name': '',
                    'Additional Name': '',
                    'Family Name': '',
                    'Mobile Phone': df[phone_col]
                })
                
                buffer = io.BytesIO()
                output.to_csv(buffer, index=False, encoding='utf-8-sig')
                buffer.seek(0)
                
                return send_file(
                    buffer,
                    mimetype='text/csv',
                    as_attachment=True,
                    download_name='contacts.csv'
                )
            else:  # vcf format
                vcf_content = create_vcf(df, name_col, phone_col)
                buffer = io.BytesIO(vcf_content.encode('utf-8'))
                buffer.seek(0)
                
                return send_file(
                    buffer,
                    mimetype='text/vcard',
                    as_attachment=True,
                    download_name='contacts.vcf'
                )

        except Exception as e:
            return str(e), 400

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)