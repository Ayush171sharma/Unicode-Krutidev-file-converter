from flask import Flask, render_template, request, send_file, redirect, url_for
from werkzeug.utils import secure_filename
import os
import docx

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = './uploads'
app.config['CONVERTED_FOLDER'] = './converted'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['CONVERTED_FOLDER'], exist_ok=True)

# Load the conversion map from a text file
def load_conversion_map(conversion_type):
    conversion_map = {}
    map_file = "unicode_to_krutidev_map.txt" if conversion_type == "unicode_to_krutidev" else "krutidev_to_unicode_map.txt"
    with open(map_file, "r", encoding="utf-8") as file:
        for line in file:
            parts = line.strip().split(' ')
            if len(parts) == 2:
                original_char, converted_char = parts
                conversion_map[original_char] = converted_char
    return conversion_map

# Convert the text based on the conversion map
def convert_text(text, conversion_map):
    if text is None:
        return None
    for original_char, converted_char in conversion_map.items():
        text = text.replace(original_char, converted_char)
    return text

# Apply the conversion to each paragraph
def apply_conversion(paragraph, conversion_map, target_font):
    if paragraph.text:
        paragraph.text = convert_text(paragraph.text, conversion_map)
        for run in paragraph.runs:
            run.font.name = target_font

# Main function to process the .docx file
def process_docx(file_path, conversion_type):
    doc = docx.Document(file_path)
    conversion_map = load_conversion_map(conversion_type)
    target_font = "Kruti Dev 010" if conversion_type == "unicode_to_krutidev" else "Mangal"

    # Apply conversion to all paragraphs
    for paragraph in doc.paragraphs:
        apply_conversion(paragraph, conversion_map, target_font)

    # Convert headers and footers
    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            apply_conversion(paragraph, conversion_map, target_font)
        for paragraph in section.footer.paragraphs:
            apply_conversion(paragraph, conversion_map, target_font)

    # Convert text in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    apply_conversion(paragraph, conversion_map, target_font)

    # Convert captions
    for paragraph in doc.paragraphs:
        if paragraph.style.name.startswith('Caption'):
            apply_conversion(paragraph, conversion_map, target_font)

    # Turn off small caps in heading styles
    for style in doc.styles:
        if style.type == docx.enum.style.WD_STYLE_TYPE.PARAGRAPH and 'Heading' in style.name:
            if style.font:
                style.font.small_caps = False
            style.font.name = target_font

    output_file = os.path.join(app.config['CONVERTED_FOLDER'], os.path.basename(file_path))
    doc.save(output_file)
    return output_file

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)

    file = request.files['file']
    conversion_type = request.form.get('conversionType')

    if file.filename == '':
        return redirect(request.url)

    if file:
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        output_file = process_docx(file_path, conversion_type)

        return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
