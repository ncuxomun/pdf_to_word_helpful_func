#%%
import sys
import os, os.path
import comtypes.client
from tqdm import tqdm

wdFormatPDF = 17
wdFormatHTML = 8

# %%
def conv_rft2pdf(input_dir, output_dir):

    for subdir, dirs, files in tqdm(os.walk(input_dir), desc="Walking through .RTF files", colour='blue'):
        for file in tqdm(iterable=files, desc="Converting .RTF files", colour='green'):
            in_file = os.path.join(subdir, file)
            # output_file = file.split('.')[0]
            output_file = os.path.splitext(file)[0]

            # Opening Word file
            word = comtypes.client.CreateObject('Word.Application')
            doc = word.Documents.Open(in_file)
            
            # to PDF
            out_file_pdf = output_dir + output_file +'.pdf'
            doc.SaveAs(out_file_pdf, FileFormat=wdFormatPDF)

            # # to HTML
            out_file_html = output_dir + output_file +'.html'
            doc.SaveAs(out_file_html, FileFormat=wdFormatHTML)
            
            # Close and Quit Word
            doc.Close()
            word.Quit()


        # for subdir, dirs, files in os.walk(input_dir):
        #     for file in files:
        #         in_file = os.path.join(subdir, file)
        #         # output_file = file.split('.')[0]
        #         output_file = os.path.splitext(file)[0]
                
        #         out_file = output_dir+output_file+'.pdf'
        #         word = comtypes.client.CreateObject('Word.Application')

        #         doc = word.Documents.Open(in_file)
        #         doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        #         doc.Close()
        #         word.Quit()

# %%
if __name__ == "__main__":
    input_dir = os.getcwd() + "\\test_in"
    output_dir = os.getcwd() + "\\test_out\\"

    conv_rft2pdf(input_dir, output_dir)