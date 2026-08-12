import pypandoc
import os
import sys

CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".pdfscan2word")
pandoc_exe = os.path.join(CONFIG_DIR, "pandoc" + (".exe" if sys.platform == "win32" else ""))
if os.path.exists(pandoc_exe):
    os.environ['PYPANDOC_PANDOC'] = pandoc_exe

markdown_content = "Here is some text with tab &#9; and emsp &emsp; and a table:\n\n| A | B |\n|---|---|\n| 1 | 2 |"
output_docx_path = "test_entities.docx"

try:
    pypandoc.convert_text(markdown_content, 'docx', format='md', outputfile=output_docx_path)
    print("Success! Entities supported.")
except Exception as e:
    print("Exception occurred:", str(e))
