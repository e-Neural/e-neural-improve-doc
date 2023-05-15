# e-neural-improve-doc
Improve documents .docx with AI

# Requirements 
python 3.10

# Clone Project
git clone https://github.com/e-Neural/e-neural-improve-doc.git

# Install Requirements
pip install -r requirements.txt --no-cache

# Execution
## Set Environment Variable
### Linux/MacOS
export open_key=<YOUR_API_KEY>
### Windows
set open_key=<YOUR_API_KEY>

python improve-doc.py <path_original_doc> <path_save_new_document> <document_language>

# Exemple
python improve-doc.py X:/e-neural-improve-doc/teste.docx X:/e-neural-improve-doc/newDocument.docx en/ptb