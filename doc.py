from flask import Flask, render_template, request
from docx import Document

app = Flask(__name__)

def search_related_content(doc_file, search_string):
    """
    Search for content related to a specific string in a .docx document.

    Parameters:
        doc_file (str): Path to the existing .docx document.
        search_string (str): The string to search for.

    Returns:
        list: List of tuples containing paragraphs related to the search string.
              Each tuple contains the paragraph text and the style name.
    """
    doc = Document(doc_file)
    related_content = []
    for paragraph in doc.paragraphs:
        if search_string.lower() in paragraph.text.lower():
            related_content.append((paragraph.text, paragraph.style.name))
    return related_content

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        doc_file = request.files['docfile']
        search_string = request.form['search_string']
        doc_file_path = 'uploaded_document.docx'
        doc_file.save(doc_file_path)
        found_content = search_related_content(doc_file_path, search_string)
        return render_template('result.html', found_content=found_content)
    return render_template('index.html')

if __name__ == "__main__":
    app.run(debug=True)
