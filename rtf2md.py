#%%
import sys
import os, os.path
import comtypes.client
from tqdm import tqdm
import pypandoc

wdFormatPDF = 17
wdFormatHTML = 8

# %%
def conv_rtf2md(input_dir, output_dir):

    for subdir, dirs, files in tqdm(os.walk(input_dir), desc="Walking through .RTF files", colour='blue'):
        for file in tqdm(iterable=files, desc="Converting .RTF files to MD", colour='yellow'):
            in_file = os.path.join(subdir, file)
            # output_file = file.split('.')[0]
            output_file = os.path.splitext(file)[0]

            # to Markdown, feeding absolute path as input
            out_file_md = output_dir + output_file +'.md'
            pypandoc.convert_file(source_file=in_file, to='md', outputfile=out_file_md)

# %%
if __name__ == "__main__":
    input_dir = os.getcwd() + "\\test_in"
    output_dir = os.getcwd() + "\\test_out\\"

    conv_rtf2md(input_dir, output_dir)