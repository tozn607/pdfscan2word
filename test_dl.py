import pypandoc
import os
import sys

CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".pdfscan2word_test")
os.makedirs(CONFIG_DIR, exist_ok=True)
pypandoc.download_pandoc(targetfolder=CONFIG_DIR, download_folder=CONFIG_DIR)

print("Downloaded files:")
for f in os.listdir(CONFIG_DIR):
    print(f)
