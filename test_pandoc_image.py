import pypandoc
import os
import sys

CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".pdfscan2word")
pandoc_exe = os.path.join(CONFIG_DIR, "pandoc" + (".exe" if sys.platform == "win32" else ""))
if os.path.exists(pandoc_exe):
    os.environ['PYPANDOC_PANDOC'] = pandoc_exe

markdown_content = "# Hello\nThis is a test with a missing image ![alt](missing_image.png)."
output_docx_path = "test_image_output.docx"

try:
    pypandoc.convert_text(markdown_content, 'docx', format='md', outputfile=output_docx_path)
    print("Success!")
except Exception as e:
    print("Exception occurred:", str(e))
