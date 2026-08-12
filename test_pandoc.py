import os
import sys
import pypandoc

CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".pdfscan2word")
pandoc_exe = os.path.join(CONFIG_DIR, "pandoc" + (".exe" if sys.platform == "win32" else ""))
if os.path.exists(pandoc_exe):
    os.environ['PYPANDOC_PANDOC'] = pandoc_exe

markdown_content = "# Hello\nThis is a test."
output_docx_path = "test_output.docx"

try:
    pypandoc.convert_text(markdown_content, 'docx', format='md', outputfile=output_docx_path)
    print("Success! File saved at:", output_docx_path)
except Exception as e:
    print("Exception occurred:", str(e))
    import traceback
    traceback.print_exc()
