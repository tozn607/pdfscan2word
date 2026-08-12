import pypandoc
import os
import sys

CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".pdfscan2word")
pandoc_exe = os.path.join(CONFIG_DIR, "pandoc" + (".exe" if sys.platform == "win32" else ""))
if os.path.exists(pandoc_exe):
    os.environ['PYPANDOC_PANDOC'] = pandoc_exe

output_docx_path = "test_readonly.docx"
# Create an empty file
with open(output_docx_path, 'w') as f:
    f.write('test')
# Make it read-only
os.chmod(output_docx_path, 0o400)

try:
    pypandoc.convert_text("# test", 'docx', format='md', outputfile=output_docx_path)
    print("Success! Overwrote readonly file.")
except Exception as e:
    print("Exception occurred:", type(e).__name__, "-", str(e))
finally:
    # Cleanup
    os.chmod(output_docx_path, 0o600)
    os.remove(output_docx_path)
