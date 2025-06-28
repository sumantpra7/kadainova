import warnings
warnings.filterwarnings("ignore", category=UserWarning)



import sys
import os
import subprocess
import re
from docxtpl import DocxTemplate
from docx import Document

# ✅ Use correct path
TEMPLATE = os.path.join(os.path.dirname(__file__), "jrsu_word_file.docx")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "zxcvbnm")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --detect mode
if '--detect' in sys.argv:
    doc = Document(TEMPLATE)
    text = "\n".join(para.text for para in doc.paragraphs)
    found = re.findall(r"{{(.*?)}}", text)
    for key in sorted(set(i.strip() for i in found if i.strip())):
        print(key)
    sys.exit(0)

# Gather context from CLI args
context = {}
for arg in sys.argv[1:]:
    if arg.startswith('--') and '=' in arg:
        key, val = arg[2:].split('=', 1)
        context[key] = val

# Fill docx
doc = DocxTemplate(TEMPLATE)
doc.render(context)
filled_docx = os.path.join(OUTPUT_DIR, "output_filled.docx")
doc.save(filled_docx)

# Convert to PDF
libre_profile = "/tmp/libreoffice-profile"
os.makedirs(libre_profile, exist_ok=True)
subprocess.run([
    "libreoffice", "--headless",
    f"-env:UserInstallation=file://{libre_profile}",
    "--convert-to", "pdf",
    filled_docx, "--outdir", OUTPUT_DIR
])
